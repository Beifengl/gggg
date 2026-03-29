using System;
using System.Management;
using System.Security.Cryptography;
using System.Text;

namespace FileBatchPrinterGUI
{
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
}