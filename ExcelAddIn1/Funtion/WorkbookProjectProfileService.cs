using ExcelAddIn1.Core;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class WorkbookProjectProfileService
    {
        public const string ManifestPropertyName = "TTBMVN.ProjectProfile.Manifest";
        private const string PartPropertyPrefix = "TTBMVN.ProjectProfile.Part.";
        private const int ChunkLength = 240;
        private const int StoreVersion = 1;
        private const int MaxPartCount = 1000;

        public static bool Save(Excel.Workbook workbook, ProjectProfile profile)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            string payload = ProjectProfileSerializer.Serialize(profile);
            return SavePayload(workbook, payload);
        }

        public static bool RestorePayload(Excel.Workbook workbook, string payload)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (payload == null)
                return Clear(workbook);

            ProjectProfileSerializer.Deserialize(payload);
            return SavePayload(workbook, payload);
        }

        public static bool Clear(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            object properties = null;
            try
            {
                properties = workbook.CustomDocumentProperties;
                bool found = false;
                DeleteProperties(properties, name =>
                {
                    bool match = string.Equals(
                        name,
                        ManifestPropertyName,
                        StringComparison.OrdinalIgnoreCase) ||
                        name.StartsWith(PartPropertyPrefix, StringComparison.OrdinalIgnoreCase);
                    found |= match;
                    return match;
                });
                return found;
            }
            finally
            {
                ReleaseComObject(properties);
            }
        }

        private static bool SavePayload(Excel.Workbook workbook, string payload)
        {
            string checksum = ProjectProfileSerializer.ComputeChecksum(payload);
            string existingPayload;
            if (TryReadPayload(workbook, out existingPayload) &&
                string.Equals(
                    ProjectProfileSerializer.ComputeChecksum(existingPayload),
                    checksum,
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string generation = Guid.NewGuid().ToString("N");
            IReadOnlyList<string> chunks = Split(payload);
            object properties = null;
            bool manifestCommitted = false;
            try
            {
                properties = workbook.CustomDocumentProperties;
                for (int index = 0; index < chunks.Count; index++)
                    SetStringProperty(properties, BuildPartName(generation, index), chunks[index]);

                string manifest = string.Join("|", new[]
                {
                    StoreVersion.ToString(CultureInfo.InvariantCulture),
                    generation,
                    chunks.Count.ToString(CultureInfo.InvariantCulture),
                    checksum
                });
                SetStringProperty(properties, ManifestPropertyName, manifest);
                manifestCommitted = true;

                DeleteProperties(properties, name =>
                    name.StartsWith(PartPropertyPrefix, StringComparison.OrdinalIgnoreCase) &&
                    !name.StartsWith(PartPropertyPrefix + generation + ".", StringComparison.OrdinalIgnoreCase));
                return true;
            }
            catch
            {
                if (!manifestCommitted && properties != null)
                {
                    try
                    {
                        DeleteProperties(properties, name =>
                            name.StartsWith(PartPropertyPrefix + generation + ".", StringComparison.OrdinalIgnoreCase));
                    }
                    catch
                    {
                    }
                }
                throw;
            }
            finally
            {
                ReleaseComObject(properties);
            }
        }

        public static bool TryLoad(Excel.Workbook workbook, out ProjectProfile profile)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            string payload;
            if (!TryReadPayload(workbook, out payload))
            {
                profile = null;
                return false;
            }

            profile = ProjectProfileSerializer.Deserialize(payload);
            return true;
        }

        public static ProjectProfile LoadRequired(Excel.Workbook workbook)
        {
            ProjectProfile profile;
            if (!TryLoad(workbook, out profile))
                throw new InvalidOperationException("Workbook chua co project profile.");
            return profile;
        }

        public static bool TryReadPayload(Excel.Workbook workbook, out string payload)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            object properties = null;
            try
            {
                properties = workbook.CustomDocumentProperties;
                string manifest;
                if (!TryGetStringProperty(properties, ManifestPropertyName, out manifest))
                {
                    payload = null;
                    return false;
                }

                string generation;
                int partCount;
                string expectedChecksum;
                ParseManifest(manifest, out generation, out partCount, out expectedChecksum);

                var builder = new System.Text.StringBuilder(partCount * ChunkLength);
                for (int index = 0; index < partCount; index++)
                {
                    string chunk;
                    string partName = BuildPartName(generation, index);
                    if (!TryGetStringProperty(properties, partName, out chunk))
                        throw new InvalidDataException("Project profile thieu chunk " + index + ".");
                    builder.Append(chunk);
                }

                payload = builder.ToString();
                string actualChecksum = ProjectProfileSerializer.ComputeChecksum(payload);
                if (!string.Equals(actualChecksum, expectedChecksum, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Project profile checksum khong khop.");
                return true;
            }
            finally
            {
                ReleaseComObject(properties);
            }
        }

        private static IReadOnlyList<string> Split(string payload)
        {
            var chunks = new List<string>();
            for (int offset = 0; offset < payload.Length; offset += ChunkLength)
                chunks.Add(payload.Substring(offset, Math.Min(ChunkLength, payload.Length - offset)));

            if (chunks.Count == 0)
                chunks.Add(string.Empty);
            if (chunks.Count > MaxPartCount)
                throw new InvalidOperationException("Project profile vuot qua kich thuoc luu tru cho phep.");
            return chunks.AsReadOnly();
        }

        private static string BuildPartName(string generation, int index)
        {
            return PartPropertyPrefix + generation + "." + index.ToString("D4", CultureInfo.InvariantCulture);
        }

        private static void ParseManifest(
            string manifest,
            out string generation,
            out int partCount,
            out string checksum)
        {
            string[] parts = (manifest ?? string.Empty).Split('|');
            int storeVersion;
            Guid parsedGeneration;
            if (parts.Length != 4 ||
                !int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out storeVersion) ||
                storeVersion != StoreVersion ||
                !Guid.TryParseExact(parts[1], "N", out parsedGeneration) ||
                !int.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out partCount) ||
                partCount < 1 || partCount > MaxPartCount ||
                parts[3].Length != 64)
            {
                throw new InvalidDataException("Project profile manifest khong hop le.");
            }

            generation = parsedGeneration.ToString("N");
            checksum = parts[3];
        }

        private static bool TryGetStringProperty(
            object propertiesObject,
            string propertyName,
            out string value)
        {
            dynamic properties = propertiesObject;
            for (int index = 1; index <= properties.Count; index++)
            {
                object propertyObject = null;
                try
                {
                    dynamic property = properties[index];
                    propertyObject = property;
                    if (!string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    value = Convert.ToString(property.Value, CultureInfo.InvariantCulture) ?? string.Empty;
                    return true;
                }
                finally
                {
                    ReleaseComObject(propertyObject);
                }
            }

            value = null;
            return false;
        }

        private static void SetStringProperty(
            object propertiesObject,
            string propertyName,
            string value)
        {
            dynamic properties = propertiesObject;
            object existingObject = null;
            try
            {
                for (int index = 1; index <= properties.Count; index++)
                {
                    object propertyObject = null;
                    try
                    {
                        dynamic property = properties[index];
                        propertyObject = property;
                        if (!string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        existingObject = propertyObject;
                        propertyObject = null;
                        break;
                    }
                    finally
                    {
                        ReleaseComObject(propertyObject);
                    }
                }

                if (existingObject == null)
                    existingObject = properties.Add(propertyName, false, MsoDocProperties.msoPropertyTypeString, value, Type.Missing);
                else
                {
                    dynamic existing = existingObject;
                    existing.Value = value;
                }
            }
            finally
            {
                ReleaseComObject(existingObject);
            }
        }

        private static void DeleteProperties(object propertiesObject, Func<string, bool> predicate)
        {
            dynamic properties = propertiesObject;
            for (int index = properties.Count; index >= 1; index--)
            {
                object propertyObject = null;
                try
                {
                    dynamic property = properties[index];
                    propertyObject = property;
                    if (predicate(property.Name ?? string.Empty))
                        property.Delete();
                }
                finally
                {
                    ReleaseComObject(propertyObject);
                }
            }
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
