using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FileBatchPrinterGUI
{
    // ==================== 硬件ID生成器 ====================
    public static class DeviceIdGenerator
    {
        public static string GetDeviceId()
        {
            string cpuId = GetCpuId();
            string diskId = GetDiskId();
            string mac = GetMacAddress();
            string rawId = $"{cpuId}|{diskId}|{mac}";
            return ComputeSHA256(rawId).Substring(0, 16);
        }

        private static string GetCpuId()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT ProcessorId FROM Win32_Processor"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        string id = obj["ProcessorId"]?.ToString();
                        if (!string.IsNullOrEmpty(id)) return id;
                    }
                }
            }
            catch { }
            return "CPU_UNKNOWN";
        }

        private static string GetDiskId()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_DiskDrive WHERE Index=0"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        string id = obj["SerialNumber"]?.ToString();
                        if (!string.IsNullOrEmpty(id)) return id;
                    }
                }
            }
            catch { }
            return "DISK_UNKNOWN";
        }

        private static string GetMacAddress()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=true"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        string mac = obj["MACAddress"]?.ToString();
                        if (!string.IsNullOrEmpty(mac)) return mac.Replace(":", "").ToUpper();
                    }
                }
            }
            catch { }
            return "MAC_UNKNOWN";
        }

        private static string ComputeSHA256(string input)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(input));
                StringBuilder sb = new StringBuilder();
                foreach (byte b in bytes) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }
    }

    // ==================== 授权结果类 ====================
    public class LicenseCheckResult
    {
        public bool Success { get; set; }
        public string DeviceCode { get; set; } = "";
        public string DeviceName { get; set; } = "";
        public DateTime ExpireDate { get; set; }
        public int RemainingDays { get; set; }
        public bool IsActive { get; set; }
        public string Code { get; set; } = "";
        public string Error { get; set; } = "";
    }

    // ==================== Supabase 授权客户端 ====================
    public class SupabaseLicenseClient
    {
        private readonly HttpClient _httpClient;
        private readonly string _supabaseUrl;
        private readonly string _apiKey;
        private readonly string _deviceCode;

        public SupabaseLicenseClient(string supabaseUrl, string apiKey, string deviceCode)
        {
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(8);
            _supabaseUrl = supabaseUrl.TrimEnd('/');
            _apiKey = apiKey;
            _deviceCode = deviceCode;
        }

        public string GetDeviceCode() => _deviceCode;

        public async Task<LicenseCheckResult> CheckLicenseAsync()
        {
            try
            {
                string url = $"{_supabaseUrl}/rest/v1/device_licenses?device_code=eq.{Uri.EscapeDataString(_deviceCode)}&select=*";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Add("apikey", _apiKey);
                request.Headers.Add("Authorization", $"Bearer {_apiKey}");

                var response = await _httpClient.SendAsync(request);
                if (!response.IsSuccessStatusCode)
                {
                    return new LicenseCheckResult { Success = false, Code = "API_ERROR", Error = $"API请求失败: {response.StatusCode}" };
                }

                var json = await response.Content.ReadAsStringAsync();
                var devices = JsonSerializer.Deserialize<List<DeviceRecord>>(json);

                if (devices == null || devices.Count == 0)
                {
                    return new LicenseCheckResult { Success = false, Code = "NOT_FOUND", Error = "设备未授权", DeviceCode = _deviceCode };
                }

                var device = devices[0];

                if (!device.is_active)
                {
                    return new LicenseCheckResult { Success = false, Code = "INACTIVE", Error = "设备已被禁用", DeviceName = device.device_name };
                }

                if (DateTime.TryParse(device.expire_date, out var expireDate))
                {
                    var remainingDays = (expireDate - DateTime.Today).Days;
                    if (remainingDays < 0)
                    {
                        return new LicenseCheckResult { Success = false, Code = "EXPIRED", Error = $"授权已于 {expireDate:yyyy-MM-dd} 过期", ExpireDate = expireDate };
                    }

                    return new LicenseCheckResult
                    {
                        Success = true,
                        DeviceCode = device.device_code,
                        DeviceName = device.device_name,
                        ExpireDate = expireDate,
                        RemainingDays = remainingDays,
                        IsActive = device.is_active
                    };
                }

                return new LicenseCheckResult { Success = false, Code = "DATA_ERROR", Error = "授权数据格式错误" };
            }
            catch (HttpRequestException ex)
            {
                return new LicenseCheckResult { Success = false, Code = "NETWORK_ERROR", Error = $"网络连接失败: {ex.Message}" };
            }
            catch (Exception ex)
            {
                return new LicenseCheckResult { Success = false, Code = "UNKNOWN_ERROR", Error = ex.Message };
            }
        }

        private class DeviceRecord
        {
            public string device_code { get; set; } = "";
            public string device_name { get; set; } = "";
            public string expire_date { get; set; } = "";
            public bool is_active { get; set; } = true;
            public string notes { get; set; } = "";
        }
    }

    // ==================== 主窗体 ====================
    public partial class Form1 : Form
    {
        private List<string> filesList = new List<string>();
        private readonly List<string> supportPrintExts = new List<string>
        {
            ".txt", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
            ".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tiff"
        };

        private TextBox txtDirectory;
        private Button btnScan;
        private TextBox txtSearch;
        private Button btnSearch;
        private CheckedListBox clbFiles;
        private Button btnSelectAll;
        private Button btnSelectByExt;
        private Button btnPrint;
        private Label lblStatus;
        private Label lblSelectedCount;
        private Label lblAuthorVersion;
        private Label lblEmail;
        private Label lblExpireDate;

        private SupabaseLicenseClient _licenseClient;
        private LicenseCheckResult _currentLicense;

        // Supabase 配置
        private const string SUPABASE_URL = "https://ddibitljconqinagnzcl.supabase.co";
        private const string SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRkaWJpdGxqY29ucWluYWduemNsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzIyMjEyMDYsImV4cCI6MjA4Nzc5NzIwNn0.limm3ilAaT-D0P-smIOC6RQDXQXBOTOOaW4FwEZa8rA";

        public Form1()
        {
            // 初始化 HTTP 协议（解决 .NET Framework 的 TLS 问题）
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;

            string deviceCode = DeviceIdGenerator.GetDeviceId();
            _licenseClient = new SupabaseLicenseClient(SUPABASE_URL, SUPABASE_ANON_KEY, deviceCode);

            // 同步运行异步方法
            var task = Task.Run(async () => await CheckLicenseAsync());
            if (!task.Result)
            {
                Environment.Exit(0);
                return;
            }

            LoadCustomIcon();
            SetupUI();
        }

        private async Task<bool> CheckLicenseAsync()
        {
            var result = await _licenseClient.CheckLicenseAsync();

            if (result.Success)
            {
                _currentLicense = result;
                return true;
            }

            // 在主线程显示消息框
            if (InvokeRequired)
            {
                return (bool)Invoke(new Func<Task<bool>>(async () => await CheckLicenseAsync()));
            }

            switch (result.Code)
            {
                case "NETWORK_ERROR":
                    MessageBox.Show("无法连接授权服务器，请检查网络后重试。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case "NOT_FOUND":
                    MessageBox.Show($"设备未授权！\n\n设备码：{_licenseClient.GetDeviceCode()}\n\n请联系管理员在 Supabase 中添加此设备。\n\n管理员操作：登录 Supabase，在 device_licenses 表中插入记录，设置 expire_date 为授权截止日期。",
                        "需要授权", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                case "EXPIRED":
                    MessageBox.Show($"授权已于 {result.ExpireDate:yyyy-MM-dd} 过期，请联系管理员续期。", "授权过期", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
                case "INACTIVE":
                    MessageBox.Show("设备已被禁用，请联系管理员。", "授权禁用", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                default:
                    MessageBox.Show($"验证失败：{result.Error}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
            return false;
        }

        private void LoadCustomIcon()
        {
            try
            {
                string iconPath = Path.Combine(Application.StartupPath, "Print.ico");
                if (File.Exists(iconPath)) this.Icon = new Icon(iconPath);
            }
            catch { }
        }

        private void SetupUI()
        {
            this.Text = "检验表批量打印工具";
            this.Size = new Size(800, 590);
            this.StartPosition = FormStartPosition.CenterScreen;

            // 目录行
            Label lblDir = new Label { Text = "目录路径:", Location = new Point(12, 15), Size = new Size(70, 23) };
            txtDirectory = new TextBox { Location = new Point(88, 12), Size = new Size(540, 23) };
            btnScan = new Button { Text = "扫描", Location = new Point(634, 10), Size = new Size(75, 27) };
            btnScan.Click += BtnScan_Click;

            // 搜索行
            Label lblSearch = new Label { Text = "关键词:", Location = new Point(12, 50), Size = new Size(70, 23) };
            txtSearch = new TextBox { Location = new Point(88, 47), Size = new Size(460, 23) };
            btnSearch = new Button { Text = "搜索", Location = new Point(554, 45), Size = new Size(75, 27) };
            btnSearch.Click += BtnSearch_Click;

            // 文件列表标题行
            Label lblFiles = new Label { Text = "文件列表（勾选要打印的文件）:", Location = new Point(12, 85), Size = new Size(200, 23) };
            lblEmail = new Label { Text = "Email: guoqiang.w@cn.interplex.com", Location = new Point(500, 85), Size = new Size(250, 23), ForeColor = Color.Green };

            // 文件列表
            clbFiles = new CheckedListBox
            {
                Location = new Point(12, 110),
                Size = new Size(760, 300),
                CheckOnClick = true,
                HorizontalScrollbar = true
            };
            clbFiles.Format += (s, e) => { if (e.ListItem is string path) e.Value = Path.GetFileName(path); };
            clbFiles.ItemCheck += ClbFiles_ItemCheck;

            // 底部一行
            int bottomY = 420;
            btnSelectAll = new Button { Text = "全选", Location = new Point(12, bottomY), Size = new Size(75, 27) };
            btnSelectAll.Click += BtnSelectAll_Click;

            btnSelectByExt = new Button { Text = "按后缀选择", Location = new Point(95, bottomY), Size = new Size(90, 27) };
            btnSelectByExt.Click += BtnSelectByExt_Click;

            lblSelectedCount = new Label { Text = "已选择 0 个文件", Location = new Point(195, bottomY + 2), Size = new Size(120, 23) };
            lblAuthorVersion = new Label { Text = "作者:王国强 Rev.A01", Location = new Point(325, bottomY + 2), Size = new Size(150, 23), ForeColor = Color.DarkBlue, Font = new Font("Segoe UI", 9, FontStyle.Bold) };

            DateTime expireDate = _currentLicense?.ExpireDate ?? DateTime.Today.AddDays(30);
            lblExpireDate = new Label { Text = $"授权至:{expireDate:yyyy-MM-dd}", Location = new Point(490, bottomY + 2), Size = new Size(120, 23), ForeColor = Color.Red };

            btnPrint = new Button { Text = "批量打印", Location = new Point(690, bottomY), Size = new Size(90, 27) };
            btnPrint.Click += BtnPrint_Click;

            // 状态栏
            lblStatus = new Label { Text = "就绪", Location = new Point(12, bottomY + 40), Size = new Size(760, 50), BorderStyle = BorderStyle.Fixed3D };

            this.Controls.AddRange(new Control[] { lblDir, txtDirectory, btnScan, lblSearch, txtSearch, btnSearch, lblFiles, lblEmail, clbFiles, btnSelectAll, btnSelectByExt, lblSelectedCount, lblAuthorVersion, lblExpireDate, btnPrint, lblStatus });
        }

        private void ClbFiles_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)(() => { lblSelectedCount.Text = $"已选择 {clbFiles.CheckedItems.Count} 个文件"; }));
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            if (clbFiles.Items.Count == 0) return;
            bool allChecked = clbFiles.CheckedItems.Count == clbFiles.Items.Count;
            for (int i = 0; i < clbFiles.Items.Count; i++) clbFiles.SetItemChecked(i, !allChecked);
            btnSelectAll.Text = allChecked ? "全选" : "取消全选";
        }

        // 按后缀名批量选择
        private void BtnSelectByExt_Click(object sender, EventArgs e)
        {
            if (clbFiles.Items.Count == 0)
            {
                MessageBox.Show("请先扫描目录！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 使用自定义输入对话框（避免依赖 Microsoft.VisualBasic）
            string input = ShowInputDialog("请输入要选择的后缀名（多个用逗号或空格分隔）\n\n示例: pdf,jpg,ppt,xlsx", "按后缀选择文件");

            if (string.IsNullOrWhiteSpace(input)) return;

            // 分割后缀名
            string[] extensions = input.Split(new char[] { ',', ' ', ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < extensions.Length; i++)
            {
                string ext = extensions[i].Trim().ToLower();
                if (!ext.StartsWith(".")) ext = "." + ext;
                extensions[i] = ext;
            }

            int selectedCount = 0;
            for (int i = 0; i < clbFiles.Items.Count; i++)
            {
                string filePath = clbFiles.Items[i] as string;
                if (string.IsNullOrEmpty(filePath)) continue;

                string fileExt = Path.GetExtension(filePath).ToLower();
                if (extensions.Contains(fileExt))
                {
                    clbFiles.SetItemChecked(i, true);
                    selectedCount++;
                }
            }

            lblSelectedCount.Text = $"已选择 {clbFiles.CheckedItems.Count} 个文件";
            lblStatus.Text = $"按后缀选择完成，共选中 {selectedCount} 个文件。";
        }

        // 自定义输入对话框（避免依赖 Microsoft.VisualBasic）
        private string ShowInputDialog(string prompt, string title)
        {
            Form dialog = new Form();
            dialog.Text = title;
            dialog.Size = new Size(400, 160);
            dialog.StartPosition = FormStartPosition.CenterParent;
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialog.MaximizeBox = false;
            dialog.MinimizeBox = false;

            Label lblPrompt = new Label();
            lblPrompt.Text = prompt;
            lblPrompt.Location = new Point(12, 15);
            lblPrompt.Size = new Size(360, 50);
            lblPrompt.AutoSize = false;

            TextBox txtInput = new TextBox();
            txtInput.Location = new Point(12, 70);
            txtInput.Size = new Size(360, 23);

            Button btnOK = new Button();
            btnOK.Text = "确定";
            btnOK.Location = new Point(200, 100);
            btnOK.Size = new Size(80, 25);
            btnOK.DialogResult = DialogResult.OK;

            Button btnCancel = new Button();
            btnCancel.Text = "取消";
            btnCancel.Location = new Point(290, 100);
            btnCancel.Size = new Size(80, 25);
            btnCancel.DialogResult = DialogResult.Cancel;

            dialog.Controls.Add(lblPrompt);
            dialog.Controls.Add(txtInput);
            dialog.Controls.Add(btnOK);
            dialog.Controls.Add(btnCancel);

            dialog.AcceptButton = btnOK;
            dialog.CancelButton = btnCancel;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                return txtInput.Text.Trim();
            }
            return "";
        }

        private void BtnScan_Click(object sender, EventArgs e)
        {
            string dirPath = txtDirectory.Text.Trim();
            if (string.IsNullOrEmpty(dirPath)) { MessageBox.Show("请输入目录路径！"); return; }
            if (!Directory.Exists(dirPath)) { MessageBox.Show("目录不存在！"); return; }

            filesList.Clear();
            clbFiles.Items.Clear();
            lblStatus.Text = "正在扫描...";

            try
            {
                string[] files = Directory.GetFiles(dirPath, "*.*", SearchOption.AllDirectories);
                int count = 0;
                foreach (string file in files)
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (supportPrintExts.Contains(ext))
                    {
                        filesList.Add(file);
                        clbFiles.Items.Add(file);
                        count++;
                    }
                    if (count % 10 == 0) Application.DoEvents();
                }
                lblStatus.Text = $"扫描完成！共找到 {filesList.Count} 个可打印文件。";
                lblSelectedCount.Text = "已选择 0 个文件";
                btnSelectAll.Text = "全选";
            }
            catch (Exception ex)
            {
                filesList.Clear();
                clbFiles.Items.Clear();
                lblStatus.Text = $"扫描出错: {ex.Message}";
            }
        }

        // 搜索：支持 Excel 竖向复制
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            if (filesList.Count == 0)
            {
                MessageBox.Show("请先扫描目录！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string keywords = txtSearch.Text.Trim();
            if (string.IsNullOrEmpty(keywords))
            {
                clbFiles.Items.Clear();
                foreach (string file in filesList) clbFiles.Items.Add(file);
                lblStatus.Text = "已显示全部文件。";
                return;
            }

            string[] keywordList = keywords.Split(new char[] { ' ', '\r', '\n', ',', ';', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            keywordList = keywordList.Distinct().ToArray();

            List<string> results = new List<string>();
            foreach (string file in filesList)
            {
                bool match = false;
                foreach (string keyword in keywordList)
                {
                    if (file.IndexOf(keyword.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        match = true;
                        break;
                    }
                }
                if (match) results.Add(file);
            }

            clbFiles.Items.Clear();
            foreach (string file in results) clbFiles.Items.Add(file);

            lblStatus.Text = $"搜索完成，关键词数：{keywordList.Length}，匹配文件数：{results.Count}";
            lblSelectedCount.Text = $"已选择 {clbFiles.CheckedItems.Count} 个文件";
        }

        private void PrintTextFile(string filePath)
        {
            string[] lines = File.ReadAllLines(filePath, Encoding.Default);
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += (sender, e) =>
            {
                float yPos = 0;
                int count = 0;
                float leftMargin = e.MarginBounds.Left;
                float topMargin = e.MarginBounds.Top;
                while (count < lines.Length)
                {
                    yPos = topMargin + (count * 12);
                    e.Graphics.DrawString(lines[count], new Font("宋体", 10), Brushes.Black, leftMargin, yPos);
                    count++;
                    if (yPos > e.MarginBounds.Bottom) { e.HasMorePages = true; return; }
                }
                e.HasMorePages = false;
            };
            pd.Print();
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (clbFiles.CheckedItems.Count == 0)
            {
                MessageBox.Show("请先勾选要打印的文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int success = 0, fail = 0;
            lblStatus.Text = "开始提交打印任务...";
            Application.DoEvents();

            foreach (var item in clbFiles.CheckedItems)
            {
                string file = item as string;
                try
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (ext == ".txt")
                    {
                        PrintTextFile(file);
                    }
                    else
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = file,
                            Verb = "print",
                            CreateNoWindow = true,
                            UseShellExecute = true,
                            WindowStyle = ProcessWindowStyle.Hidden
                        });
                    }
                    success++;
                    System.Threading.Thread.Sleep(1500);
                }
                catch (Exception ex)
                {
                    fail++;
                    lblStatus.Text = $"打印失败: {Path.GetFileName(file)} - {ex.Message}";
                    Application.DoEvents();
                }
            }

            lblStatus.Text = $"打印任务提交完成！成功: {success}, 失败: {fail}。";
            MessageBox.Show($"打印完成！成功 {success} 个，失败 {fail} 个。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    // ==================== 程序入口 ====================
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
