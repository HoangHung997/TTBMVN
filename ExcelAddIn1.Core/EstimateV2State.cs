using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace ExcelAddIn1.Core
{
    public sealed class EstimateV2WorkItemState
    {
        public EstimateV2WorkItemState(
            string workItemId,
            string normCode,
            string variantCode,
            string packageId,
            string dataVersion,
            string kind,
            string fingerprint,
            bool isOrphaned)
        {
            WorkItemId = NormalizeId(workItemId);
            NormCode = Clean(normCode);
            VariantCode = Clean(variantCode);
            PackageId = Clean(packageId);
            DataVersion = Clean(dataVersion);
            Kind = Clean(kind);
            Fingerprint = NormalizeFingerprint(fingerprint);
            IsOrphaned = isOrphaned;
        }

        public string WorkItemId { get; }
        public string NormCode { get; }
        public string VariantCode { get; }
        public string PackageId { get; }
        public string DataVersion { get; }
        public string Kind { get; }
        public string Fingerprint { get; }
        public bool IsOrphaned { get; }

        public bool HasNormBinding => NormCode.Length > 0;

        public EstimateV2WorkItemState WithBinding(
            string normCode,
            string variantCode,
            string packageId,
            string dataVersion)
        {
            return new EstimateV2WorkItemState(
                WorkItemId,
                normCode,
                variantCode,
                packageId,
                dataVersion,
                Kind,
                Fingerprint,
                IsOrphaned);
        }

        public EstimateV2WorkItemState WithoutBinding()
        {
            return new EstimateV2WorkItemState(
                WorkItemId,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                Kind,
                Fingerprint,
                IsOrphaned);
        }

        public EstimateV2WorkItemState WithFingerprint(string fingerprint)
        {
            return new EstimateV2WorkItemState(
                WorkItemId,
                NormCode,
                VariantCode,
                PackageId,
                DataVersion,
                Kind,
                fingerprint,
                IsOrphaned);
        }

        public EstimateV2WorkItemState WithOrphaned(bool isOrphaned)
        {
            return new EstimateV2WorkItemState(
                WorkItemId,
                NormCode,
                VariantCode,
                PackageId,
                DataVersion,
                Kind,
                Fingerprint,
                isOrphaned);
        }

        public static string CreateId()
        {
            return Guid.NewGuid().ToString("N");
        }

        public static bool IsValidId(string value)
        {
            Guid parsed;
            return Guid.TryParseExact((value ?? string.Empty).Trim(), "N", out parsed);
        }

        private static string NormalizeId(string value)
        {
            Guid parsed;
            string text = (value ?? string.Empty).Trim();
            if (!Guid.TryParseExact(text, "N", out parsed))
                throw new ArgumentException("WorkItemId phai la GUID dang N (32 ky tu).", nameof(value));
            return parsed.ToString("N");
        }

        private static string NormalizeFingerprint(string value)
        {
            string text = Clean(value).ToUpperInvariant();
            if (text.Length == 0)
                return string.Empty;
            if (text.Length != 64 || text.Any(ch =>
                !((ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'F'))))
            {
                throw new ArgumentException("Fingerprint phai la SHA-256 hex 64 ky tu.", nameof(value));
            }
            return text;
        }

        private static string Clean(string value)
        {
            return (value ?? string.Empty).Trim();
        }
    }

    public sealed class EstimateV2State
    {
        public const int CurrentSchemaVersion = 1;

        public EstimateV2State(
            IEnumerable<EstimateV2WorkItemState> workItems,
            DateTime updatedUtc,
            int schemaVersion = CurrentSchemaVersion)
        {
            if (schemaVersion != CurrentSchemaVersion)
                throw new ArgumentOutOfRangeException(nameof(schemaVersion));
            SchemaVersion = schemaVersion;
            UpdatedUtc = NormalizeUtc(updatedUtc);

            var items = (workItems ?? Enumerable.Empty<EstimateV2WorkItemState>())
                .Where(item => item != null)
                .OrderBy(item => item.WorkItemId, StringComparer.Ordinal)
                .ToList();
            string duplicate = items
                .GroupBy(item => item.WorkItemId, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .FirstOrDefault();
            if (duplicate != null)
                throw new ArgumentException("Trung WorkItemId: " + duplicate, nameof(workItems));

            WorkItems = new ReadOnlyCollection<EstimateV2WorkItemState>(items);
        }

        public int SchemaVersion { get; }
        public DateTime UpdatedUtc { get; }
        public IReadOnlyList<EstimateV2WorkItemState> WorkItems { get; }

        public EstimateV2WorkItemState Find(string workItemId)
        {
            string id = (workItemId ?? string.Empty).Trim();
            return WorkItems.FirstOrDefault(item =>
                string.Equals(item.WorkItemId, id, StringComparison.OrdinalIgnoreCase));
        }

        public EstimateV2State Upsert(EstimateV2WorkItemState item, DateTime updatedUtc)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item));
            var result = WorkItems
                .Where(existing => !string.Equals(
                    existing.WorkItemId,
                    item.WorkItemId,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();
            result.Add(item);
            return new EstimateV2State(result, updatedUtc, SchemaVersion);
        }

        public EstimateV2State Remove(string workItemId, DateTime updatedUtc)
        {
            string id = (workItemId ?? string.Empty).Trim();
            return new EstimateV2State(
                WorkItems.Where(item => !string.Equals(
                    item.WorkItemId,
                    id,
                    StringComparison.OrdinalIgnoreCase)),
                updatedUtc,
                SchemaVersion);
        }

        public static EstimateV2State Empty(DateTime updatedUtc)
        {
            return new EstimateV2State(
                Enumerable.Empty<EstimateV2WorkItemState>(),
                updatedUtc);
        }

        private static DateTime NormalizeUtc(DateTime value)
        {
            if (value.Kind == DateTimeKind.Utc)
                return value;
            if (value.Kind == DateTimeKind.Unspecified)
                return DateTime.SpecifyKind(value, DateTimeKind.Utc);
            return value.ToUniversalTime();
        }
    }

    public static class EstimateV2Fingerprint
    {
        public static string Compute(
            string workCode,
            string description,
            string unit,
            string kind)
        {
            string payload = string.Join("|", new[]
            {
                Normalize(workCode),
                Normalize(description),
                Normalize(unit),
                Normalize(kind)
            });
            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(payload));
                var builder = new StringBuilder(bytes.Length * 2);
                foreach (byte value in bytes)
                    builder.Append(value.ToString("X2", CultureInfo.InvariantCulture));
                return builder.ToString();
            }
        }

        private static string Normalize(string value)
        {
            string text = (value ?? string.Empty).Trim();
            if (text.Length == 0)
                return string.Empty;
            var builder = new StringBuilder(text.Length);
            bool previousWhitespace = false;
            foreach (char ch in text)
            {
                if (char.IsWhiteSpace(ch))
                {
                    if (!previousWhitespace)
                        builder.Append(' ');
                    previousWhitespace = true;
                }
                else
                {
                    builder.Append(char.ToUpperInvariant(ch));
                    previousWhitespace = false;
                }
            }
            return builder.ToString();
        }
    }

    public static class EstimateV2StateSerializer
    {
        public const string NamespaceUri = "urn:ttbmvn:estimate-v2:state:v1";
        private static readonly XNamespace Ns = NamespaceUri;

        public static string Serialize(EstimateV2State state)
        {
            if (state == null)
                throw new ArgumentNullException(nameof(state));

            var root = new XElement(
                Ns + "estimateState",
                new XAttribute("schemaVersion", state.SchemaVersion),
                new XAttribute("updatedUtc", state.UpdatedUtc.ToString("o", CultureInfo.InvariantCulture)),
                state.WorkItems.Select(item =>
                    new XElement(
                        Ns + "workItem",
                        new XAttribute("id", item.WorkItemId),
                        new XAttribute("normCode", item.NormCode),
                        new XAttribute("variantCode", item.VariantCode),
                        new XAttribute("packageId", item.PackageId),
                        new XAttribute("dataVersion", item.DataVersion),
                        new XAttribute("kind", item.Kind),
                        new XAttribute("fingerprint", item.Fingerprint),
                        new XAttribute("orphaned", item.IsOrphaned ? "1" : "0"))));
            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root)
                .ToString(SaveOptions.DisableFormatting);
        }

        public static EstimateV2State Deserialize(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
                throw new InvalidDataException("Estimate V2 state trong.");

            XDocument document;
            try
            {
                document = XDocument.Parse(payload, LoadOptions.None);
            }
            catch (Exception ex) when (ex is System.Xml.XmlException || ex is ArgumentException)
            {
                throw new InvalidDataException("Estimate V2 state XML khong hop le.", ex);
            }

            XElement root = document.Root;
            if (root == null || root.Name != Ns + "estimateState")
                throw new InvalidDataException("Estimate V2 state sai namespace/root.");

            int schemaVersion;
            if (!int.TryParse(
                RequiredAttribute(root, "schemaVersion"),
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out schemaVersion) ||
                schemaVersion != EstimateV2State.CurrentSchemaVersion)
            {
                throw new InvalidDataException("Estimate V2 state schemaVersion khong duoc ho tro.");
            }

            DateTime updatedUtc;
            if (!DateTime.TryParse(
                RequiredAttribute(root, "updatedUtc"),
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out updatedUtc))
            {
                throw new InvalidDataException("Estimate V2 state updatedUtc khong hop le.");
            }

            try
            {
                var items = root.Elements(Ns + "workItem")
                    .Select(element => new EstimateV2WorkItemState(
                        RequiredAttribute(element, "id"),
                        OptionalAttribute(element, "normCode"),
                        OptionalAttribute(element, "variantCode"),
                        OptionalAttribute(element, "packageId"),
                        OptionalAttribute(element, "dataVersion"),
                        OptionalAttribute(element, "kind"),
                        OptionalAttribute(element, "fingerprint"),
                        string.Equals(
                            OptionalAttribute(element, "orphaned"),
                            "1",
                            StringComparison.Ordinal)))
                    .ToList();

                if (root.Elements().Any(element => element.Name != Ns + "workItem"))
                    throw new InvalidDataException("Estimate V2 state co node khong duoc ho tro.");

                return new EstimateV2State(items, updatedUtc, schemaVersion);
            }
            catch (ArgumentException ex)
            {
                throw new InvalidDataException("Estimate V2 state workItem khong hop le.", ex);
            }
        }

        private static string RequiredAttribute(XElement element, string name)
        {
            XAttribute attribute = element.Attribute(name);
            if (attribute == null || string.IsNullOrWhiteSpace(attribute.Value))
                throw new InvalidDataException("Thieu attribute " + name + ".");
            return attribute.Value;
        }

        private static string OptionalAttribute(XElement element, string name)
        {
            XAttribute attribute = element.Attribute(name);
            return attribute == null ? string.Empty : attribute.Value;
        }
    }
}
