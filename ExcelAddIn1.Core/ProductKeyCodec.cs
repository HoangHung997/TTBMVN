using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace ExcelAddIn1.Core
{
    public static class ProductKeyCodec
    {
        public const string ProductKeyPrefix = "TTB26";
        public const string SigningKeyFileName = "license-signing-key-v2.bin";

        private const byte FormatVersion = 2;
        private const int PayloadLength = 11;
        private const int SignatureLength = 64;
        private const string ProductKeyAlphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
        private const string PublicKeyBlobBase64 =
            "RUNTMSAAAACn+AQAhlYIYP9OtZz7v5885jUOwrRNJ6Mgdq/uXWeDD1fG3yXCu7OJswj9XBG/6rVM555zVyGG6hB9dzqIIhsG";

        private static readonly DateTime LicenseEpoch = new DateTime(2020, 1, 1);

        public static string Generate(string machineId, DateTime expiryDate, byte[] privateKeyBlob)
        {
            if (privateKeyBlob == null || privateKeyBlob.Length == 0)
                throw new ArgumentException("Thieu private key de ky license.", nameof(privateKeyBlob));

            byte[] payload = BuildPayload(machineId, expiryDate);
            byte[] signature = Sign(payload, privateKeyBlob);
            if (signature.Length != SignatureLength)
                throw new CryptographicException("Chu ky ECDSA khong dung do dai P-256.");

            byte[] keyBytes = new byte[payload.Length + signature.Length];
            Buffer.BlockCopy(payload, 0, keyBytes, 0, payload.Length);
            Buffer.BlockCopy(signature, 0, keyBytes, payload.Length, signature.Length);
            return ProductKeyPrefix + "-" + GroupEvery(Base32Encode(keyBytes), 5);
        }

        public static bool TryValidate(string licenseKey, string machineId, out DateTime expiryDate)
        {
            return TryValidate(
                licenseKey,
                machineId,
                Convert.FromBase64String(PublicKeyBlobBase64),
                out expiryDate);
        }

        public static bool TryValidate(
            string licenseKey,
            string machineId,
            byte[] publicKeyBlob,
            out DateTime expiryDate)
        {
            expiryDate = DateTime.MinValue;
            if (string.IsNullOrWhiteSpace(licenseKey) || publicKeyBlob == null || publicKeyBlob.Length == 0)
                return false;

            string[] parts = licenseKey.Trim().Split('-');
            if (parts.Length < 2 ||
                !string.Equals(parts[0], ProductKeyPrefix, StringComparison.OrdinalIgnoreCase))
                return false;

            string body = ConcatParts(parts, 1);
            if (!TryBase32Decode(body, out byte[] keyBytes) ||
                keyBytes.Length != PayloadLength + SignatureLength)
                return false;

            byte[] payload = new byte[PayloadLength];
            byte[] signature = new byte[SignatureLength];
            Buffer.BlockCopy(keyBytes, 0, payload, 0, payload.Length);
            Buffer.BlockCopy(keyBytes, payload.Length, signature, 0, signature.Length);

            if (payload[0] != FormatVersion || !Verify(payload, signature, publicKeyBlob))
                return false;

            ushort expiryDays = (ushort)((payload[1] << 8) | payload[2]);
            expiryDate = LicenseEpoch.AddDays(expiryDays).Date;

            byte[] machineBytes = new byte[8];
            Buffer.BlockCopy(payload, 3, machineBytes, 0, machineBytes.Length);
            return string.Equals(
                BytesToHex(machineBytes),
                NormalizeMachineIdForLicense(machineId),
                StringComparison.OrdinalIgnoreCase);
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

        private static byte[] BuildPayload(string machineId, DateTime expiryDate)
        {
            string normalizedMachineId = NormalizeMachineIdForLicense(machineId);
            ushort expiryDays = ToExpiryDays(expiryDate);
            byte[] payload = new byte[PayloadLength];
            payload[0] = FormatVersion;
            payload[1] = (byte)(expiryDays >> 8);
            payload[2] = (byte)(expiryDays & 0xFF);
            byte[] machineBytes = HexToBytes(normalizedMachineId);
            Buffer.BlockCopy(machineBytes, 0, payload, 3, machineBytes.Length);
            return payload;
        }

        private static byte[] Sign(byte[] payload, byte[] privateKeyBlob)
        {
            using (CngKey key = CngKey.Import(privateKeyBlob, CngKeyBlobFormat.EccPrivateBlob))
            using (var signer = new ECDsaCng(key) { HashAlgorithm = CngAlgorithm.Sha256 })
                return signer.SignData(payload);
        }

        private static bool Verify(byte[] payload, byte[] signature, byte[] publicKeyBlob)
        {
            try
            {
                using (CngKey key = CngKey.Import(publicKeyBlob, CngKeyBlobFormat.EccPublicBlob))
                using (var verifier = new ECDsaCng(key) { HashAlgorithm = CngAlgorithm.Sha256 })
                    return verifier.VerifyData(payload, signature);
            }
            catch (CryptographicException)
            {
                return false;
            }
        }

        private static string NormalizeMachineIdForLicense(string machineId)
        {
            string normalized = NormalizeMachineId(machineId);
            if (normalized.Length == 16)
                return normalized;

            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(normalized));
                return BytesToHex(hash).Substring(0, 16);
            }
        }

        private static ushort ToExpiryDays(DateTime expiryDate)
        {
            double days = (expiryDate.Date - LicenseEpoch).TotalDays;
            if (days < 0)
                return 0;
            if (days > ushort.MaxValue)
                return ushort.MaxValue;
            return Convert.ToUInt16(days);
        }

        private static string Base32Encode(byte[] data)
        {
            if (data == null || data.Length == 0)
                return string.Empty;

            var builder = new StringBuilder();
            int buffer = 0;
            int bitsLeft = 0;
            foreach (byte value in data)
            {
                buffer = (buffer << 8) | value;
                bitsLeft += 8;
                while (bitsLeft >= 5)
                {
                    int index = (buffer >> (bitsLeft - 5)) & 31;
                    bitsLeft -= 5;
                    builder.Append(ProductKeyAlphabet[index]);
                }

                buffer &= bitsLeft == 0 ? 0 : (1 << bitsLeft) - 1;
            }

            if (bitsLeft > 0)
                builder.Append(ProductKeyAlphabet[(buffer << (5 - bitsLeft)) & 31]);
            return builder.ToString();
        }

        private static bool TryBase32Decode(string text, out byte[] data)
        {
            data = null;
            string normalized = NormalizeMachineId(text);
            if (string.IsNullOrWhiteSpace(normalized))
                return false;

            var bytes = new List<byte>();
            int buffer = 0;
            int bitsLeft = 0;
            foreach (char c in normalized)
            {
                int value = ProductKeyAlphabet.IndexOf(c);
                if (value < 0)
                    return false;

                buffer = (buffer << 5) | value;
                bitsLeft += 5;
                if (bitsLeft >= 8)
                {
                    bytes.Add((byte)((buffer >> (bitsLeft - 8)) & 0xFF));
                    bitsLeft -= 8;
                    buffer &= bitsLeft == 0 ? 0 : (1 << bitsLeft) - 1;
                }
            }

            data = bytes.ToArray();
            return true;
        }

        private static byte[] HexToBytes(string hex)
        {
            hex = NormalizeMachineId(hex);
            if (hex.Length % 2 != 0)
                hex = "0" + hex;

            byte[] bytes = new byte[hex.Length / 2];
            for (int i = 0; i < bytes.Length; i++)
                bytes[i] = byte.Parse(hex.Substring(i * 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            return bytes;
        }

        private static string BytesToHex(byte[] bytes)
        {
            return BitConverter.ToString(bytes).Replace("-", string.Empty).ToUpperInvariant();
        }

        private static string GroupEvery(string value, int groupSize)
        {
            value = NormalizeMachineId(value);
            var parts = new List<string>();
            for (int i = 0; i < value.Length; i += groupSize)
                parts.Add(value.Substring(i, Math.Min(groupSize, value.Length - i)));
            return string.Join("-", parts);
        }

        private static string ConcatParts(string[] parts, int startIndex)
        {
            var builder = new StringBuilder();
            for (int i = startIndex; i < parts.Length; i++)
                builder.Append(parts[i]);
            return builder.ToString();
        }
    }
}
