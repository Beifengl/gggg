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
            string rawId = cpuId + "|" + diskId + "|" + mac;
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
            _httpClient.Timeout = TimeSpan.FromSeconds(10);
            _supabaseUrl = supabaseUrl.TrimEnd('/');
            _apiKey = apiKey;
            _deviceCode = deviceCode;
        }

        public string GetDeviceCode() => _deviceCode;

        public async Task<LicenseCheckResult> CheckLicenseAsync()
        {
            try
            {
                string url = _supabaseUrl + "/rest/v1/device_licenses?device_code=eq." + Uri.EscapeDataString(_deviceCode) + "&select=*";

                using (var request = new HttpRequestMessage(HttpMethod.Get, url))
                {
                    request.Headers.Add("apikey", _apiKey);
                    request.Headers.Add("Authorization", "Bearer " + _apiKey);

                    var response = await _httpClient.SendAsync(request);
                    string json = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                    {
                        return new LicenseCheckResult { Success = false, Code = "API_ERROR", Error = "API请求失败: " + response.StatusCode };
                    }

                    if (string.IsNullOrWhiteSpace(json) || json == "[]")
                    {
                        return new LicenseCheckResult { Success = false, Code = "NOT_FOUND", Error = "设备未授权", DeviceCode = _deviceCode };
                    }

                    // 简单解析 JSON
                    string deviceCode = "";
                    string deviceName = "";
                    string expireDateStr = "";
                    bool isActive = true;

                    int idx = json.IndexOf("device_code");
                    if (idx > 0)
                    {
                        int start = json.IndexOf("\"", idx + 12) + 1;
                        int end = json.IndexOf("\"", start);
                        if (start > 0 && end > start) deviceCode = json.Substring(start, end - start);
                    }

                    idx = json.IndexOf("device_name");
                    if (idx > 0)
                    {
                        int start = json.IndexOf("\"", idx + 12) + 1;
                        int end = json.IndexOf("\"", start);
                        if (start > 0 && end > start) deviceName = json.Substring(start, end - start);
                    }

                    idx = json.IndexOf("expire_date");
                    if (idx > 0)
                    {
                        int start = json.IndexOf("\"", idx + 12) + 1;
                        int end = json.IndexOf("\"", start);
                        if (start > 0 && end > start) expireDateStr = json.Substring(start, end - start);
                    }

                    idx = json.IndexOf("is_active");
                    if (idx > 0)
                    {
                        int start = idx + 11;
                        while (start < json.Length && json[start] == ' ') start++;
                        if (start < json.Length)
                        {
                            isActive = json[start] == 't';
                        }
                    }

                    if (!isActive)
                    {
                        return new LicenseCheckResult { Success = false, Code = "INACTIVE", Error = "设备已被禁用", DeviceName = deviceName };
                    }

                    if (DateTime.TryParse(expireDateStr, out var expireDate))
                    {
                        var remainingDays = (expireDate - DateTime.Today).Days;
                        if (remainingDays < 0)
                        {
                            return new LicenseCheckResult { Success = false, Code = "EXPIRED", Error = "授权已于 " + expireDate.ToString("yyyy-MM-dd") + " 过期", ExpireDate = expireDate };
                        }

                        return new LicenseCheckResult
                        {
                            Success = true,
                            DeviceCode = deviceCode,
                            DeviceName = deviceName,
                            ExpireDate = expireDate,
                            RemainingDays = remainingDays,
                            IsActive = isActive
                        };
                    }

                    return new LicenseCheckResult { Success = false, Code = "DATA_ERROR", Error = "授权数据格式错误" };
                }
            }
            catch (HttpRequestException ex)
            {
                return new LicenseCheckResult { Success = false, Code = "NETWORK_ERROR", Error = "网络连接失败: " + ex.Message };
            }
            catch (Exception ex)
            {
                return new LicenseCheckResult { Success = false, Code = "UNKNOWN_ERROR", Error = ex.Message };
            }
        }
    }

    // ==================== 授权窗口 ====================
    public class LicenseDialog : Form
    {
        private Label lblDeviceCode;
        private Label lblStatus;
        private Label lblExpireDate;
        private Button btnRefresh;
        private Button btnClose;
        private SupabaseLicenseClient _licenseClient;
        private string _deviceCode;
        private bool _isValid = false;
        private LicenseCheckResult _currentLicense;

        public bool IsValid => _isValid;
        public LicenseCheckResult CurrentLicense => _currentLicense;

        public LicenseDialog(string supabaseUrl, string apiKey, string deviceCode)
        {
            _deviceCode = deviceCode;
            _licenseClient = new SupabaseLicenseClient(supabaseUrl, apiKey, deviceCode);

            SetupUI();
            this.Load += async (s, e) => await CheckLicenseAsync();
        }

        private void SetupUI()
        {
            this.Text = "授权验证";
            this.Size = new Size(450, 280);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Label lblTitle = new Label();
            lblTitle.Text = "设备码:";
            lblTitle.Location = new Point(20, 25);
            lblTitle.Size = new Size(60, 23);
            lblTitle.Font = new Font("微软雅黑", 9, FontStyle.Bold);

            lblDeviceCode = new Label();
            lblDeviceCode.Text = _deviceCode;
            lblDeviceCode.Location = new Point(90, 25);
            lblDeviceCode.Size = new Size(320, 23);
            lblDeviceCode.Font = new Font("Consolas", 9, FontStyle.Regular);
            lblDeviceCode.BackColor = Color.WhiteSmoke;
            lblDeviceCode.BorderStyle = BorderStyle.FixedSingle;
            lblDeviceCode.TextAlign = ContentAlignment.MiddleLeft;

            Label lblStatusTitle = new Label();
            lblStatusTitle.Text = "状态:";
            lblStatusTitle.Location = new Point(20, 65);
            lblStatusTitle.Size = new Size(60, 23);
            lblStatusTitle.Font = new Font("微软雅黑", 9, FontStyle.Bold);

            lblStatus = new Label();
            lblStatus.Text = "正在验证...";
            lblStatus.Location = new Point(90, 65);
            lblStatus.Size = new Size(320, 23);
            lblStatus.Font = new Font("微软雅黑", 9, FontStyle.Regular);
            lblStatus.ForeColor = Color.Blue;

            Label lblExpireTitle = new Label();
            lblExpireTitle.Text = "有效期:";
            lblExpireTitle.Location = new Point(20, 105);
            lblExpireTitle.Size = new Size(60, 23);
            lblExpireTitle.Font = new Font("微软雅黑", 9, FontStyle.Bold);

            lblExpireDate = new Label();
            lblExpireDate.Text = "---";
            lblExpireDate.Location = new Point(90, 105);
            lblExpireDate.Size = new Size(320, 23);
            lblExpireDate.Font = new Font("微软雅黑", 9, FontStyle.Regular);

            Label lblTip = new Label();
            lblTip.Text = "提示：如果设备未授权，请联系管理员在 Supabase 中添加此设备码。";
            lblTip.Location = new Point(20, 150);
            lblTip.Size = new Size(400, 40);
            lblTip.Font = new Font("微软雅黑", 8, FontStyle.Italic);
            lblTip.ForeColor = Color.Gray;

            btnRefresh = new Button();
            btnRefresh.Text = "刷新";
            btnRefresh.Location = new Point(220, 210);
            btnRefresh.Size = new Size(80, 30);
            btnRefresh.FlatStyle = FlatStyle.Flat;
            btnRefresh.BackColor = Color.LightGray;
            btnRefresh.Click += async (s, e) => await CheckLicenseAsync();

            btnClose = new Button();
            btnClose.Text = "关闭";
            btnClose.Location = new Point(330, 210);
            btnClose.Size = new Size(80, 30);
            btnClose.FlatStyle = FlatStyle.Flat;
            btnClose.BackColor = Color.LightGray;
            btnClose.Click += (s, e) => this.Close();

            this.Controls.Add(lblTitle);
            this.Controls.Add(lblDeviceCode);
            this.Controls.Add(lblStatusTitle);
            this.Controls.Add(lblStatus);
            this.Controls.Add(lblExpireTitle);
            this.Controls.Add(lblExpireDate);
            this.Controls.Add(lblTip);
            this.Controls.Add(btnRefresh);
            this.Controls.Add(btnClose);
        }

        private async Task CheckLicenseAsync()
        {
            btnRefresh.Enabled = false;
            lblStatus.Text = "正在验证...";
            lblStatus.ForeColor = Color.Blue;
            Application.DoEvents();

            var result = await _licenseClient.CheckLicenseAsync();

            if (result.Success)
            {
                _isValid = true;
                _currentLicense = result;
                lblStatus.Text = "✓ 授权有效";
                lblStatus.ForeColor = Color.Green;
                lblExpireDate.Text = result.ExpireDate.ToString("yyyy-MM-dd") + " (剩余 " + result.RemainingDays + " 天)";
                lblExpireDate.ForeColor = result.RemainingDays <= 7 ? Color.Orange : Color.Black;
                btnRefresh.Text = "进入系统";
            }
            else
            {
                _isValid = false;
                switch (result.Code)
                {
                    case "NETWORK_ERROR":
                        lblStatus.Text = "✗ 网络连接失败";
                        lblStatus.ForeColor = Color.Red;
                        lblExpireDate.Text = "请检查网络后点击刷新";
                        break;
                    case "NOT_FOUND":
                        lblStatus.Text = "✗ 设备未授权";
                        lblStatus.ForeColor = Color.Red;
                        lblExpireDate.Text = "请联系管理员添加此设备码";
                        break;
                    case "EXPIRED":
                        lblStatus.Text = "✗ 授权已过期";
                        lblStatus.ForeColor = Color.Red;
                        lblExpireDate.Text = "已于 " + result.ExpireDate.ToString("yyyy-MM-dd") + " 过期";
                        break;
                    case "INACTIVE":
                        lblStatus.Text = "✗ 设备已被禁用";
                        lblStatus.ForeColor = Color.Red;
                        lblExpireDate.Text = "请联系管理员启用";
                        break;
                    default:
                        lblStatus.Text = "✗ " + result.Error;
                        lblStatus.ForeColor = Color.Red;
                        lblExpireDate.Text = "请点击刷新重试";
                        break;
                }
                btnRefresh.Text = "刷新";
            }

            btnRefresh.Enabled = true;
        }
    }

    // ==================== 主窗体 ====================
    public partial class Form1 : Form
    {
        private List<string> filesList = new List<string>();
        private readonly List<string> supportPrintExts = new List<string>();

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

        private LicenseCheckResult _currentLicense;

        private const string SUPABASE_URL = "https://ddibitljconqinagnzcl.supabase.co";
        private const string SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRkaWJpdGxqY29ucWluYWduemNsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzIyMjEyMDYsImV4cCI6MjA4Nzc5NzIwNn0.limm3ilAaT-D0P-smIOC6RQDXQXBOTOOaW4FwEZa8rA";

        public Form1()
        {
            // 初始化支持的后缀列表
            supportPrintExts.AddRange(new string[] { ".txt", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tiff" });

            string deviceCode = DeviceIdGenerator.GetDeviceId();
            using (var licenseDialog = new LicenseDialog(SUPABASE_URL, SUPABASE_ANON_KEY, deviceCode))
            {
                licenseDialog.ShowDialog();
                if (!licenseDialog.IsValid)
                {
                    Environment.Exit(0);
                    return;
                }
                _currentLicense = licenseDialog.CurrentLicense;
            }

            InitializeComponent();
            SetupUI();
        }

        private void InitializeComponent()
        {
            this.Text = "检验表批量打印工具";
            this.Size = new Size(800, 590);
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void SetupUI()
        {
            Label lblDir = new Label();
            lblDir.Text = "目录路径:";
            lblDir.Location = new Point(12, 15);
            lblDir.Size = new Size(70, 23);
            
            txtDirectory = new TextBox();
            txtDirectory.Location = new Point(88, 12);
            txtDirectory.Size = new Size(540, 23);
            
            btnScan = new Button();
            btnScan.Text = "扫描";
            btnScan.Location = new Point(634, 10);
            btnScan.Size = new Size(75, 27);
            btnScan.Click += BtnScan_Click;

            Label lblSearch = new Label();
            lblSearch.Text = "关键词:";
            lblSearch.Location = new Point(12, 50);
            lblSearch.Size = new Size(70, 23);
            
            txtSearch = new TextBox();
            txtSearch.Location = new Point(88, 47);
            txtSearch.Size = new Size(460, 23);
            
            btnSearch = new Button();
            btnSearch.Text = "搜索";
            btnSearch.Location = new Point(554, 45);
            btnSearch.Size = new Size(75, 27);
            btnSearch.Click += BtnSearch_Click;

            Label lblFiles = new Label();
            lblFiles.Text = "文件列表（勾选要打印的文件）:";
            lblFiles.Location = new Point(12, 85);
            lblFiles.Size = new Size(200, 23);
            
            lblEmail = new Label();
            lblEmail.Text = "Email: guoqiang.w@cn.interplex.com";
            lblEmail.Location = new Point(500, 85);
            lblEmail.Size = new Size(250, 23);
            lblEmail.ForeColor = Color.Green;

            clbFiles = new CheckedListBox();
            clbFiles.Location = new Point(12, 110);
            clbFiles.Size = new Size(760, 300);
            clbFiles.CheckOnClick = true;
            clbFiles.HorizontalScrollbar = true;
            clbFiles.Format += (s, e) => { if (e.ListItem is string path) e.Value = Path.GetFileName(path); };
            clbFiles.ItemCheck += ClbFiles_ItemCheck;

            int bottomY = 420;
            
            btnSelectAll = new Button();
            btnSelectAll.Text = "全选";
            btnSelectAll.Location = new Point(12, bottomY);
            btnSelectAll.Size = new Size(75, 27);
            btnSelectAll.Click += BtnSelectAll_Click;

            btnSelectByExt = new Button();
            btnSelectByExt.Text = "按后缀选择";
            btnSelectByExt.Location = new Point(95, bottomY);
            btnSelectByExt.Size = new Size(90, 27);
            btnSelectByExt.Click += BtnSelectByExt_Click;

            lblSelectedCount = new Label();
            lblSelectedCount.Text = "已选择 0 个文件";
            lblSelectedCount.Location = new Point(195, bottomY + 2);
            lblSelectedCount.Size = new Size(120, 23);

            lblAuthorVersion = new Label();
            lblAuthorVersion.Text = "作者:王国强 Rev.A01";
            lblAuthorVersion.Location = new Point(325, bottomY + 2);
            lblAuthorVersion.Size = new Size(150, 23);
            lblAuthorVersion.ForeColor = Color.DarkBlue;
            lblAuthorVersion.Font = new Font("Segoe UI", 9, FontStyle.Bold);

            lblExpireDate = new Label();
            lblExpireDate.Text = "授权至:" + _currentLicense.ExpireDate.ToString("yyyy-MM-dd") + " (剩余 " + _currentLicense.RemainingDays + " 天)";
            lblExpireDate.Location = new Point(480, bottomY + 2);
            lblExpireDate.Size = new Size(180, 23);
            lblExpireDate.ForeColor = _currentLicense.RemainingDays <= 7 ? Color.Orange : Color.Green;

            btnPrint = new Button();
            btnPrint.Text = "批量打印";
            btnPrint.Location = new Point(690, bottomY);
            btnPrint.Size = new Size(90, 27);
            btnPrint.Click += BtnPrint_Click;

            lblStatus = new Label();
            lblStatus.Text = "就绪";
            lblStatus.Location = new Point(12, bottomY + 40);
            lblStatus.Size = new Size(760, 50);
            lblStatus.BorderStyle = BorderStyle.Fixed3D;

            this.Controls.Add(lblDir);
            this.Controls.Add(txtDirectory);
            this.Controls.Add(btnScan);
            this.Controls.Add(lblSearch);
            this.Controls.Add(txtSearch);
            this.Controls.Add(btnSearch);
            this.Controls.Add(lblFiles);
            this.Controls.Add(lblEmail);
            this.Controls.Add(clbFiles);
            this.Controls.Add(btnSelectAll);
            this.Controls.Add(btnSelectByExt);
            this.Controls.Add(lblSelectedCount);
            this.Controls.Add(lblAuthorVersion);
            this.Controls.Add(lblExpireDate);
            this.Controls.Add(btnPrint);
            this.Controls.Add(lblStatus);
        }

        private void ClbFiles_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)(() => { lblSelectedCount.Text = "已选择 " + clbFiles.CheckedItems.Count + " 个文件"; }));
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            if (clbFiles.Items.Count == 0) return;
            bool allChecked = clbFiles.CheckedItems.Count == clbFiles.Items.Count;
            for (int i = 0; i < clbFiles.Items.Count; i++) clbFiles.SetItemChecked(i, !allChecked);
            btnSelectAll.Text = allChecked ? "全选" : "取消全选";
        }

        private void BtnSelectByExt_Click(object sender, EventArgs e)
        {
            if (clbFiles.Items.Count == 0)
            {
                MessageBox.Show("请先扫描目录！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string input = ShowInputDialog("请输入要选择的后缀名（多个用逗号或空格分隔）\n\n示例: pdf,jpg,ppt,xlsx", "按后缀选择文件");

            if (string.IsNullOrWhiteSpace(input)) return;

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

            lblSelectedCount.Text = "已选择 " + clbFiles.CheckedItems.Count + " 个文件";
            lblStatus.Text = "按后缀选择完成，共选中 " + selectedCount + " 个文件。";
        }

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
                lblStatus.Text = "扫描完成！共找到 " + filesList.Count + " 个可打印文件。";
                lblSelectedCount.Text = "已选择 0 个文件";
                btnSelectAll.Text = "全选";
            }
            catch (Exception ex)
            {
                filesList.Clear();
                clbFiles.Items.Clear();
                lblStatus.Text = "扫描出错: " + ex.Message;
            }
        }

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

            lblStatus.Text = "搜索完成，关键词数：" + keywordList.Length + "，匹配文件数：" + results.Count;
            lblSelectedCount.Text = "已选择 " + clbFiles.CheckedItems.Count + " 个文件";
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
                    lblStatus.Text = "打印失败: " + Path.GetFileName(file) + " - " + ex.Message;
                    Application.DoEvents();
                }
            }

            lblStatus.Text = "打印任务提交完成！成功: " + success + ", 失败: " + fail + "。";
            MessageBox.Show("打印完成！成功 " + success + " 个，失败 " + fail + " 个。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
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