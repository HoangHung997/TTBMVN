using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace ExcelAddIn1.Funtion
{
    public sealed class HardwareFingerprint
    {
        public string MachineGuid { get; set; }
        public string BaseBoardSerial { get; set; }
        public string BiosSerial { get; set; }
        public List<string> DiskSerials { get; } = new List<string>();
        public string MachineName { get; set; }

        public string ToStableText()
        {
            string disks = string.Join(",", DiskSerials.OrderBy(x => x, StringComparer.OrdinalIgnoreCase));
            return "MG=" + NormalizePart(MachineGuid)
                + "|BB=" + NormalizePart(BaseBoardSerial)
                + "|BIOS=" + NormalizePart(BiosSerial)
                + "|DISK=" + NormalizePart(disks)
                + "|NAME=" + NormalizePart(MachineName);
        }

        private static string NormalizePart(string value)
        {
            return (value ?? string.Empty).Trim().ToUpperInvariant();
        }
    }

    public static class HardwareFingerprintProvider
    {
        public static HardwareFingerprint GetFingerprint()
        {
            var fingerprint = new HardwareFingerprint
            {
                MachineGuid = ReadMachineGuid(),
                BaseBoardSerial = ReadFirstWmiValue("SELECT SerialNumber FROM Win32_BaseBoard", "SerialNumber"),
                BiosSerial = ReadFirstWmiValue("SELECT SerialNumber FROM Win32_BIOS", "SerialNumber"),
                MachineName = Environment.MachineName
            };

            foreach (string diskSerial in ReadWmiValues("SELECT SerialNumber FROM Win32_DiskDrive", "SerialNumber"))
            {
                string normalized = NormalizeHardwareValue(diskSerial);
                if (!string.IsNullOrWhiteSpace(normalized) && !fingerprint.DiskSerials.Contains(normalized, StringComparer.OrdinalIgnoreCase))
                    fingerprint.DiskSerials.Add(normalized);
            }

            return fingerprint;
        }

        public static string GetMachineId()
        {
            HardwareFingerprint fingerprint = GetFingerprint();
            string raw = fingerprint.ToStableText();
            if (string.IsNullOrWhiteSpace(raw.Replace("MG=", string.Empty).Replace("|BB=", string.Empty).Replace("|BIOS=", string.Empty).Replace("|DISK=", string.Empty).Replace("|NAME=", string.Empty)))
                raw = Environment.MachineName + "|" + Environment.OSVersion.VersionString;

            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(raw));
                string hex = BitConverter.ToString(hash).Replace("-", string.Empty).Substring(0, 16).ToUpperInvariant();
                return FormatMachineId(hex);
            }
        }

        public static string NormalizeMachineId(string machineId)
        {
            var builder = new StringBuilder();
            foreach (char c in (machineId ?? string.Empty).ToUpperInvariant())
            {
                if ((c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9'))
                    builder.Append(c);
            }

            return builder.ToString();
        }

        private static string FormatMachineId(string hex)
        {
            hex = NormalizeMachineId(hex);
            if (hex.Length <= 4)
                return hex;

            var groups = new List<string>();
            for (int i = 0; i < hex.Length; i += 4)
                groups.Add(hex.Substring(i, Math.Min(4, hex.Length - i)));

            return string.Join("-", groups);
        }

        private static string ReadMachineGuid()
        {
            string value = ReadMachineGuid(RegistryView.Registry64);
            if (!string.IsNullOrWhiteSpace(value))
                return value;

            return ReadMachineGuid(RegistryView.Registry32);
        }

        private static string ReadMachineGuid(RegistryView view)
        {
            try
            {
                using (RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                using (RegistryKey key = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Cryptography"))
                {
                    return NormalizeHardwareValue(Convert.ToString(key?.GetValue("MachineGuid")));
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ReadFirstWmiValue(string query, string propertyName)
        {
            return ReadWmiValues(query, propertyName).FirstOrDefault() ?? string.Empty;
        }

        private static IEnumerable<string> ReadWmiValues(string query, string propertyName)
        {
            var result = new List<string>();
            try
            {
                using (var searcher = new ManagementObjectSearcher(query))
                using (ManagementObjectCollection objects = searcher.Get())
                {
                    foreach (ManagementObject item in objects)
                    {
                        string value = NormalizeHardwareValue(Convert.ToString(item[propertyName]));
                        if (!string.IsNullOrWhiteSpace(value))
                            result.Add(value);
                    }
                }
            }
            catch
            {
            }

            return result;
        }

        private static string NormalizeHardwareValue(string value)
        {
            string text = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            string upper = text.ToUpperInvariant();
            string compact = new string(upper.Where(char.IsLetterOrDigit).ToArray());
            if (string.IsNullOrWhiteSpace(compact))
                return string.Empty;

            string[] invalidValues =
            {
                "TOBEFILLEDBYOEM",
                "DEFAULTSTRING",
                "SYSTEMSERIALNUMBER",
                "NONE",
                "UNKNOWN",
                "NOTSPECIFIED",
                "NOTAVAILABLE",
                "0000000000000000",
                "00000000"
            };

            if (invalidValues.Contains(compact, StringComparer.OrdinalIgnoreCase))
                return string.Empty;

            if (compact.All(c => c == '0'))
                return string.Empty;

            return compact;
        }
    }
}
