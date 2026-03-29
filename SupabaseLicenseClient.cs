using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace FileBatchPrinterGUI
{
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
                var devices = JsonSerializer.Deserialize<DeviceRecord[]>(json);

                if (devices == null || devices.Length == 0)
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

        public class DeviceRecord
        {
            public string device_code { get; set; } = "";
            public string device_name { get; set; } = "";
            public string expire_date { get; set; } = "";
            public bool is_active { get; set; } = true;
            public string notes { get; set; } = "";
        }
    }
}