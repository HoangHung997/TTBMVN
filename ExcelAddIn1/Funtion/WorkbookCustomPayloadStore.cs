using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    internal static class WorkbookCustomPayloadStore
    {
        private const int ChunkLength = 240;
        private const int StoreVersion = 1;
        private const int MaxPartCount = 2000;

        internal static bool Save(
            Excel.Workbook workbook,
            string manifestPropertyName,
            string partPropertyPrefix,
            string payload,
            string label)
        {
            Required(workbook, manifestPropertyName, partPropertyPrefix, label);
            if (payload == null)
                throw new ArgumentNullException(nameof(payload));
            string checksum = ExcelAddIn1.Core.PriceProfileSerializer.ComputePayloadChecksum(payload);
            string existing;
            if (TryRead(
                workbook,
                manifestPropertyName,
                partPropertyPrefix,
                label,
                out existing) &&
                string.Equals(
                    ExcelAddIn1.Core.PriceProfileSerializer.ComputePayloadChecksum(existing),
                    checksum,
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string generation = Guid.NewGuid().ToString("N");
            IReadOnlyList<string> chunks = Split(payload, label);
            object properties = null;
            bool committed = false;
            try
            {
                properties = workbook.CustomDocumentProperties;
                for (int index = 0; index < chunks.Count; index++)
                    AddStringProperty(properties, PartName(partPropertyPrefix, generation, index), chunks[index]);
                string manifest = string.Join("|", new[]
                {
                    StoreVersion.ToString(CultureInfo.InvariantCulture),
                    generation,
                    chunks.Count.ToString(CultureInfo.InvariantCulture),
                    checksum
                });
                SetStringProperty(properties, manifestPropertyName, manifest);
                committed = true;
                try
                {
                    DeleteProperties(properties, name =>
                        name.StartsWith(partPropertyPrefix, StringComparison.OrdinalIgnoreCase) &&
                        !name.StartsWith(partPropertyPrefix + generation + ".", StringComparison.OrdinalIgnoreCase));
                }
                catch (Exception ex)
                {
                    RuntimeLogger.Log(ex, "Cleanup stale " + label + " payload parts");
                }
                return true;
            }
            catch
            {
                if (!committed && properties != null)
                {
                    try
                    {
                        DeleteProperties(properties, name =>
                            name.StartsWith(partPropertyPrefix + generation + ".", StringComparison.OrdinalIgnoreCase));
                    }
                    catch
                    {
                    }
                }
                throw;
            }
            finally
            {
                Release(properties);
            }
        }

        internal static bool TryRead(
            Excel.Workbook workbook,
            string manifestPropertyName,
            string partPropertyPrefix,
            string label,
            out string payload)
        {
            Required(workbook, manifestPropertyName, partPropertyPrefix, label);
            object properties = null;
            try
            {
                properties = workbook.CustomDocumentProperties;
                Dictionary<string, string> propertyValues = ReadStringProperties(properties);
                string manifest;
                if (!propertyValues.TryGetValue(manifestPropertyName, out manifest))
                {
                    payload = null;
                    return false;
                }
                string generation;
                int partCount;
                string checksum;
                ParseManifest(manifest, label, out generation, out partCount, out checksum);
                var builder = new System.Text.StringBuilder(partCount * ChunkLength);
                for (int index = 0; index < partCount; index++)
                {
                    string chunk;
                    if (!propertyValues.TryGetValue(
                        PartName(partPropertyPrefix, generation, index),
                        out chunk))
                    {
                        throw new InvalidDataException(label + " thieu chunk " + index + ".");
                    }
                    builder.Append(chunk);
                }
                payload = builder.ToString();
                string actual = ExcelAddIn1.Core.PriceProfileSerializer.ComputePayloadChecksum(payload);
                if (!string.Equals(actual, checksum, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException(label + " checksum khong khop.");
                return true;
            }
            finally
            {
                Release(properties);
            }
        }

        internal static bool Clear(
            Excel.Workbook workbook,
            string manifestPropertyName,
            string partPropertyPrefix,
            string label)
        {
            Required(workbook, manifestPropertyName, partPropertyPrefix, label);
            object properties = null;
            try
            {
                properties = workbook.CustomDocumentProperties;
                bool found = false;
                DeleteProperties(properties, name =>
                {
                    bool match = string.Equals(name, manifestPropertyName, StringComparison.OrdinalIgnoreCase) ||
                        name.StartsWith(partPropertyPrefix, StringComparison.OrdinalIgnoreCase);
                    found |= match;
                    return match;
                });
                return found;
            }
            finally
            {
                Release(properties);
            }
        }

        private static IReadOnlyList<string> Split(string payload, string label)
        {
            var chunks = new List<string>();
            for (int offset = 0; offset < payload.Length; offset += ChunkLength)
                chunks.Add(payload.Substring(offset, Math.Min(ChunkLength, payload.Length - offset)));
            if (chunks.Count == 0)
                chunks.Add(string.Empty);
            if (chunks.Count > MaxPartCount)
                throw new InvalidOperationException(label + " vuot kich thuoc workbook cho phep.");
            return chunks.AsReadOnly();
        }

        private static void Required(
            Excel.Workbook workbook,
            string manifestPropertyName,
            string partPropertyPrefix,
            string label)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(manifestPropertyName) ||
                string.IsNullOrWhiteSpace(partPropertyPrefix) ||
                string.IsNullOrWhiteSpace(label))
            {
                throw new ArgumentException("Workbook payload store identity trong.");
            }
        }

        private static string PartName(string prefix, string generation, int index)
        {
            return prefix + generation + "." + index.ToString("D4", CultureInfo.InvariantCulture);
        }

        private static void ParseManifest(
            string manifest,
            string label,
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
                throw new InvalidDataException(label + " manifest khong hop le.");
            }
            generation = parsedGeneration.ToString("N");
            checksum = parts[3];
        }

        private static bool TryGetStringProperty(object propertiesObject, string name, out string value)
        {
            dynamic properties = propertiesObject;
            for (int index = 1; index <= properties.Count; index++)
            {
                object propertyObject = null;
                try
                {
                    dynamic property = properties[index];
                    propertyObject = property;
                    if (!string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                        continue;
                    value = Convert.ToString(property.Value, CultureInfo.InvariantCulture) ?? string.Empty;
                    return true;
                }
                finally
                {
                    Release(propertyObject);
                }
            }
            value = null;
            return false;
        }

        private static Dictionary<string, string> ReadStringProperties(object propertiesObject)
        {
            dynamic properties = propertiesObject;
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 1; index <= properties.Count; index++)
            {
                object propertyObject = null;
                try
                {
                    dynamic property = properties[index];
                    propertyObject = property;
                    string name = Convert.ToString(property.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                    if (name.Length > 0)
                    {
                        values[name] = Convert.ToString(
                            property.Value,
                            CultureInfo.InvariantCulture) ?? string.Empty;
                    }
                }
                finally
                {
                    Release(propertyObject);
                }
            }
            return values;
        }

        private static void AddStringProperty(object propertiesObject, string name, string value)
        {
            dynamic properties = propertiesObject;
            object propertyObject = null;
            try
            {
                propertyObject = properties.Add(
                    name,
                    false,
                    MsoDocProperties.msoPropertyTypeString,
                    value,
                    Type.Missing);
            }
            finally
            {
                Release(propertyObject);
            }
        }

        private static void SetStringProperty(object propertiesObject, string name, string value)
        {
            dynamic properties = propertiesObject;
            object existing = null;
            try
            {
                for (int index = 1; index <= properties.Count; index++)
                {
                    object propertyObject = null;
                    try
                    {
                        dynamic property = properties[index];
                        propertyObject = property;
                        if (!string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                            continue;
                        existing = propertyObject;
                        propertyObject = null;
                        break;
                    }
                    finally
                    {
                        Release(propertyObject);
                    }
                }
                if (existing == null)
                    existing = properties.Add(name, false, MsoDocProperties.msoPropertyTypeString, value, Type.Missing);
                else
                {
                    dynamic property = existing;
                    property.Value = value;
                }
            }
            finally
            {
                Release(existing);
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
                    Release(propertyObject);
                }
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
