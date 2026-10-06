using ExcelAddIn1.Core;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class EstimateV2ColumnLayout
    {
        public EstimateV2ColumnLayout(
            int workCodeColumn,
            int normDisplayColumn,
            int descriptionColumn,
            int unitColumn,
            int quantityColumn,
            int technicalIdColumn,
            int technicalNormColumn,
            int technicalKindColumn,
            int technicalFingerprintColumn)
        {
            WorkCodeColumn = RequiredColumn(workCodeColumn, nameof(workCodeColumn));
            NormDisplayColumn = RequiredColumn(normDisplayColumn, nameof(normDisplayColumn));
            DescriptionColumn = RequiredColumn(descriptionColumn, nameof(descriptionColumn));
            UnitColumn = RequiredColumn(unitColumn, nameof(unitColumn));
            QuantityColumn = RequiredColumn(quantityColumn, nameof(quantityColumn));
            TechnicalIdColumn = RequiredColumn(technicalIdColumn, nameof(technicalIdColumn));
            TechnicalNormColumn = RequiredColumn(technicalNormColumn, nameof(technicalNormColumn));
            TechnicalKindColumn = RequiredColumn(technicalKindColumn, nameof(technicalKindColumn));
            TechnicalFingerprintColumn = RequiredColumn(
                technicalFingerprintColumn,
                nameof(technicalFingerprintColumn));

            int firstTechnical = new[]
            {
                TechnicalIdColumn,
                TechnicalNormColumn,
                TechnicalKindColumn,
                TechnicalFingerprintColumn
            }.Min();
            int lastVisible = new[]
            {
                WorkCodeColumn,
                NormDisplayColumn,
                DescriptionColumn,
                UnitColumn,
                QuantityColumn
            }.Max();
            if (firstTechnical <= lastVisible)
                throw new ArgumentException("Cot ky thuat phai nam sau cac cot du lieu hien thi.");
        }

        public int WorkCodeColumn { get; }
        public int NormDisplayColumn { get; }
        public int DescriptionColumn { get; }
        public int UnitColumn { get; }
        public int QuantityColumn { get; }
        public int TechnicalIdColumn { get; }
        public int TechnicalNormColumn { get; }
        public int TechnicalKindColumn { get; }
        public int TechnicalFingerprintColumn { get; }

        private static int RequiredColumn(int value, string name)
        {
            if (value < 1 || value > 16384)
                throw new ArgumentOutOfRangeException(name);
            return value;
        }
    }

    public sealed class EstimateV2ReconcileResult
    {
        public EstimateV2ReconcileResult(
            int workItemCount,
            int createdCount,
            int recoveredCount,
            int duplicateIdCount,
            int restoredNormDisplayCount,
            int orphanedCount,
            bool stateChanged,
            IEnumerable<string> messages)
        {
            WorkItemCount = workItemCount;
            CreatedCount = createdCount;
            RecoveredCount = recoveredCount;
            DuplicateIdCount = duplicateIdCount;
            RestoredNormDisplayCount = restoredNormDisplayCount;
            OrphanedCount = orphanedCount;
            StateChanged = stateChanged;
            Messages = (messages ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
        }

        public int WorkItemCount { get; }
        public int CreatedCount { get; }
        public int RecoveredCount { get; }
        public int DuplicateIdCount { get; }
        public int RestoredNormDisplayCount { get; }
        public int OrphanedCount { get; }
        public bool StateChanged { get; }
        public IReadOnlyList<string> Messages { get; }
    }

    public static class WorkbookEstimateV2StateService
    {
        public const string CustomXmlNamespace = EstimateV2StateSerializer.NamespaceUri;
        public const string HeaderId = "__TTB_ID";
        public const string HeaderNorm = "__TTB_NORM";
        public const string HeaderKind = "__TTB_KIND";
        public const string HeaderFingerprint = "__TTB_HASH";
        public const string DefaultKind = "WORKITEM";

        public static bool TryLoad(Excel.Workbook workbook, out EstimateV2State state)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            object partsObject = null;
            try
            {
                partsObject = workbook.CustomXMLParts;
                dynamic parts = partsObject;
                for (int index = parts.Count; index >= 1; index--)
                {
                    object partObject = null;
                    try
                    {
                        dynamic part = parts[index];
                        partObject = part;
                        string ns = Convert.ToString(part.NamespaceURI, CultureInfo.InvariantCulture) ?? string.Empty;
                        if (!string.Equals(ns, CustomXmlNamespace, StringComparison.Ordinal))
                            continue;
                        string xml = Convert.ToString(part.XML, CultureInfo.InvariantCulture) ?? string.Empty;
                        try
                        {
                            state = EstimateV2StateSerializer.Deserialize(xml);
                            return true;
                        }
                        catch (Exception ex)
                        {
                            RuntimeLogger.Log(ex, "Read Estimate V2 CustomXML part");
                        }
                    }
                    finally
                    {
                        Release(partObject);
                    }
                }

                state = null;
                return false;
            }
            finally
            {
                Release(partsObject);
            }
        }

        public static EstimateV2State LoadOrCreate(Excel.Workbook workbook)
        {
            EstimateV2State state;
            return TryLoad(workbook, out state)
                ? state
                : EstimateV2State.Empty(DateTime.UtcNow);
        }

        public static void Save(Excel.Workbook workbook, EstimateV2State state)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (state == null)
                throw new ArgumentNullException(nameof(state));

            string xml = EstimateV2StateSerializer.Serialize(state);
            object partsObject = null;
            object createdObject = null;
            string createdId = string.Empty;
            try
            {
                partsObject = workbook.CustomXMLParts;
                dynamic parts = partsObject;
                dynamic created = parts.Add(xml, Type.Missing);
                createdObject = created;
                createdId = Convert.ToString(created.Id, CultureInfo.InvariantCulture) ?? string.Empty;

                for (int index = parts.Count; index >= 1; index--)
                {
                    object partObject = null;
                    try
                    {
                        dynamic part = parts[index];
                        partObject = part;
                        string id = Convert.ToString(part.Id, CultureInfo.InvariantCulture) ?? string.Empty;
                        string ns = Convert.ToString(part.NamespaceURI, CultureInfo.InvariantCulture) ?? string.Empty;
                        if (!string.Equals(ns, CustomXmlNamespace, StringComparison.Ordinal) ||
                            string.Equals(id, createdId, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }
                        part.Delete();
                    }
                    finally
                    {
                        Release(partObject);
                    }
                }
            }
            catch
            {
                if (createdObject != null)
                {
                    try
                    {
                        dynamic created = createdObject;
                        created.Delete();
                    }
                    catch
                    {
                    }
                }
                throw;
            }
            finally
            {
                Release(createdObject);
                Release(partsObject);
            }
        }

        public static bool Clear(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            object partsObject = null;
            bool removed = false;
            try
            {
                partsObject = workbook.CustomXMLParts;
                dynamic parts = partsObject;
                for (int index = parts.Count; index >= 1; index--)
                {
                    object partObject = null;
                    try
                    {
                        dynamic part = parts[index];
                        partObject = part;
                        string ns = Convert.ToString(part.NamespaceURI, CultureInfo.InvariantCulture) ?? string.Empty;
                        if (!string.Equals(ns, CustomXmlNamespace, StringComparison.Ordinal))
                            continue;
                        part.Delete();
                        removed = true;
                    }
                    finally
                    {
                        Release(partObject);
                    }
                }
                return removed;
            }
            finally
            {
                Release(partsObject);
            }
        }

        public static EstimateV2ColumnLayout EnsureTechnicalColumns(
            Excel.Worksheet worksheet,
            int headerRow,
            int workCodeColumn,
            int normDisplayColumn,
            int descriptionColumn,
            int unitColumn,
            int quantityColumn,
            int preferredFirstTechnicalColumn)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (headerRow < 1)
                throw new ArgumentOutOfRangeException(nameof(headerRow));

            int existing = FindTechnicalStartColumn(worksheet, headerRow);
            int technicalStart = existing > 0
                ? existing
                : FindSafeTechnicalStartColumn(
                    worksheet,
                    Math.Max(
                        preferredFirstTechnicalColumn,
                        new[]
                        {
                            workCodeColumn,
                            normDisplayColumn,
                            descriptionColumn,
                            unitColumn,
                            quantityColumn
                        }.Max() + 1));

            WriteTechnicalHeader(worksheet, headerRow, technicalStart, HeaderId);
            WriteTechnicalHeader(worksheet, headerRow, technicalStart + 1, HeaderNorm);
            WriteTechnicalHeader(worksheet, headerRow, technicalStart + 2, HeaderKind);
            WriteTechnicalHeader(worksheet, headerRow, technicalStart + 3, HeaderFingerprint);
            HideColumns(worksheet, technicalStart, technicalStart + 3);

            return new EstimateV2ColumnLayout(
                workCodeColumn,
                normDisplayColumn,
                descriptionColumn,
                unitColumn,
                quantityColumn,
                technicalStart,
                technicalStart + 1,
                technicalStart + 2,
                technicalStart + 3);
        }

        public static EstimateV2ReconcileResult Reconcile(
            Excel.Workbook workbook,
            Excel.Worksheet worksheet,
            int firstDataRow,
            int lastDataRow,
            EstimateV2ColumnLayout columns)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (columns == null)
                throw new ArgumentNullException(nameof(columns));
            if (firstDataRow < 1 || lastDataRow < firstDataRow)
                throw new ArgumentOutOfRangeException(nameof(firstDataRow));

            HideColumns(
                worksheet,
                columns.TechnicalIdColumn,
                columns.TechnicalFingerprintColumn);

            EstimateV2State state = LoadOrCreate(workbook);
            var stateById = state.WorkItems.ToDictionary(
                item => item.WorkItemId,
                StringComparer.OrdinalIgnoreCase);
            var fingerprintCandidates = state.WorkItems
                .Where(item => item.Fingerprint.Length > 0)
                .GroupBy(item => item.Fingerprint, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => group.Key,
                    group => group.ToList(),
                    StringComparer.OrdinalIgnoreCase);

            var seenIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var output = state.WorkItems.ToDictionary(
                item => item.WorkItemId,
                StringComparer.OrdinalIgnoreCase);
            var messages = new List<string>();
            int createdCount = 0;
            int recoveredCount = 0;
            int duplicateCount = 0;
            int restoredNormCount = 0;
            bool changed = false;
            int workItemCount = 0;

            for (int row = firstDataRow; row <= lastDataRow; row++)
            {
                string workCode = ReadText(worksheet, row, columns.WorkCodeColumn);
                string description = ReadText(worksheet, row, columns.DescriptionColumn);
                string unit = ReadText(worksheet, row, columns.UnitColumn);
                string quantity = ReadText(worksheet, row, columns.QuantityColumn);
                string visibleNorm = ReadText(worksheet, row, columns.NormDisplayColumn);

                if (!IsWorkItemRow(workCode, description, unit, quantity))
                    continue;

                workItemCount++;
                string fingerprint = EstimateV2Fingerprint.Compute(
                    workCode,
                    description,
                    unit,
                    DefaultKind);
                string rowId = ReadText(worksheet, row, columns.TechnicalIdColumn);
                EstimateV2WorkItemState existing = null;

                bool validId = EstimateV2WorkItemState.IsValidId(rowId);
                if (validId)
                    stateById.TryGetValue(rowId, out existing);

                if (!validId || existing == null)
                {
                    string recoveredId = TryRecoverByFingerprint(
                        fingerprint,
                        fingerprintCandidates,
                        seenIds);
                    if (recoveredId.Length > 0)
                    {
                        rowId = recoveredId;
                        existing = stateById[rowId];
                        recoveredCount++;
                        changed = true;
                        messages.Add("Phuc hoi WorkItemId tai dong " + row + " tu fingerprint.");
                    }
                    else
                    {
                        rowId = EstimateV2WorkItemState.CreateId();
                        string importedNorm = visibleNorm;
                        existing = new EstimateV2WorkItemState(
                            rowId,
                            importedNorm,
                            string.Empty,
                            string.Empty,
                            string.Empty,
                            DefaultKind,
                            fingerprint,
                            false);
                        stateById[rowId] = existing;
                        output[rowId] = existing;
                        createdCount++;
                        changed = true;
                    }
                }
                else if (seenIds.Contains(rowId))
                {
                    string originalId = rowId;
                    rowId = EstimateV2WorkItemState.CreateId();
                    existing = new EstimateV2WorkItemState(
                        rowId,
                        existing.NormCode,
                        existing.VariantCode,
                        existing.PackageId,
                        existing.DataVersion,
                        existing.Kind.Length == 0 ? DefaultKind : existing.Kind,
                        fingerprint,
                        false);
                    stateById[rowId] = existing;
                    output[rowId] = existing;
                    duplicateCount++;
                    createdCount++;
                    changed = true;
                    messages.Add(
                        "Dong " + row + " trung WorkItemId " + originalId +
                        "; da cap ID moi va giu binding dinh muc.");
                }

                seenIds.Add(rowId);

                string kind = existing.Kind.Length == 0 ? DefaultKind : existing.Kind;
                EstimateV2WorkItemState normalized = new EstimateV2WorkItemState(
                    rowId,
                    existing.NormCode.Length > 0 ? existing.NormCode : visibleNorm,
                    existing.VariantCode,
                    existing.PackageId,
                    existing.DataVersion,
                    kind,
                    fingerprint,
                    false);
                if (!Equivalent(existing, normalized))
                {
                    output[rowId] = normalized;
                    stateById[rowId] = normalized;
                    existing = normalized;
                    changed = true;
                }

                WriteTextIfDifferent(
                    worksheet,
                    row,
                    columns.TechnicalIdColumn,
                    rowId,
                    ref changed);
                WriteTextIfDifferent(
                    worksheet,
                    row,
                    columns.TechnicalNormColumn,
                    existing.NormCode,
                    ref changed);
                WriteTextIfDifferent(
                    worksheet,
                    row,
                    columns.TechnicalKindColumn,
                    existing.Kind,
                    ref changed);
                WriteTextIfDifferent(
                    worksheet,
                    row,
                    columns.TechnicalFingerprintColumn,
                    existing.Fingerprint,
                    ref changed);

                if (existing.NormCode.Length > 0 &&
                    !string.Equals(visibleNorm, existing.NormCode, StringComparison.OrdinalIgnoreCase))
                {
                    WriteCellValue(
                        worksheet,
                        row,
                        columns.NormDisplayColumn,
                        existing.NormCode);
                    restoredNormCount++;
                    changed = true;
                }
            }

            int orphanedCount = 0;
            foreach (EstimateV2WorkItemState item in output.Values.ToList())
            {
                bool orphaned = !seenIds.Contains(item.WorkItemId);
                if (orphaned)
                    orphanedCount++;
                if (item.IsOrphaned == orphaned)
                    continue;
                output[item.WorkItemId] = item.WithOrphaned(orphaned);
                changed = true;
            }

            EstimateV2State next = new EstimateV2State(
                output.Values,
                DateTime.UtcNow);
            if (changed || !TryLoad(workbook, out _))
                Save(workbook, next);

            return new EstimateV2ReconcileResult(
                workItemCount,
                createdCount,
                recoveredCount,
                duplicateCount,
                restoredNormCount,
                orphanedCount,
                changed,
                messages);
        }

        public static bool BindNorm(
            Excel.Workbook workbook,
            string workItemId,
            string normCode,
            string variantCode,
            string packageId,
            string dataVersion)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            EstimateV2State state = LoadOrCreate(workbook);
            EstimateV2WorkItemState item = state.Find(workItemId);
            if (item == null)
                throw new InvalidOperationException("Khong tim thay WorkItemId de gan dinh muc.");
            string code = (normCode ?? string.Empty).Trim();
            if (code.Length == 0)
                throw new ArgumentException("Ma dinh muc khong duoc trong.", nameof(normCode));
            EstimateV2WorkItemState updated = item.WithBinding(
                code,
                variantCode,
                packageId,
                dataVersion).WithOrphaned(false);
            if (Equivalent(item, updated))
                return false;
            Save(workbook, state.Upsert(updated, DateTime.UtcNow));
            return true;
        }

        public static bool UnbindNorm(Excel.Workbook workbook, string workItemId)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            EstimateV2State state = LoadOrCreate(workbook);
            EstimateV2WorkItemState item = state.Find(workItemId);
            if (item == null || !item.HasNormBinding)
                return false;
            Save(workbook, state.Upsert(item.WithoutBinding(), DateTime.UtcNow));
            return true;
        }

        private static string TryRecoverByFingerprint(
            string fingerprint,
            IReadOnlyDictionary<string, List<EstimateV2WorkItemState>> candidates,
            ISet<string> seenIds)
        {
            List<EstimateV2WorkItemState> matches;
            if (!candidates.TryGetValue(fingerprint, out matches))
                return string.Empty;
            List<EstimateV2WorkItemState> available = matches
                .Where(item => !seenIds.Contains(item.WorkItemId))
                .ToList();
            return available.Count == 1 ? available[0].WorkItemId : string.Empty;
        }

        private static bool IsWorkItemRow(
            string workCode,
            string description,
            string unit,
            string quantity)
        {
            if (workCode.Length == 0 && description.Length == 0)
                return false;
            // Cac dong nhom/tieu de trong Gia DT TC thuong khong co ma, don vi va khoi luong.
            return workCode.Length > 0 || unit.Length > 0 || quantity.Length > 0;
        }

        private static bool Equivalent(
            EstimateV2WorkItemState left,
            EstimateV2WorkItemState right)
        {
            return left != null &&
                right != null &&
                string.Equals(left.WorkItemId, right.WorkItemId, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(left.NormCode, right.NormCode, StringComparison.Ordinal) &&
                string.Equals(left.VariantCode, right.VariantCode, StringComparison.Ordinal) &&
                string.Equals(left.PackageId, right.PackageId, StringComparison.Ordinal) &&
                string.Equals(left.DataVersion, right.DataVersion, StringComparison.Ordinal) &&
                string.Equals(left.Kind, right.Kind, StringComparison.Ordinal) &&
                string.Equals(left.Fingerprint, right.Fingerprint, StringComparison.OrdinalIgnoreCase) &&
                left.IsOrphaned == right.IsOrphaned;
        }

        private static int FindTechnicalStartColumn(Excel.Worksheet worksheet, int headerRow)
        {
            Excel.Range used = null;
            try
            {
                used = worksheet.UsedRange;
                int lastColumn = Math.Min(
                    16381,
                    Math.Max(1, used.Column + used.Columns.Count - 1));
                for (int column = 1; column <= lastColumn; column++)
                {
                    if (!string.Equals(
                        ReadText(worksheet, headerRow, column),
                        HeaderId,
                        StringComparison.Ordinal))
                    {
                        continue;
                    }
                    if (column + 3 > 16384)
                        return 0;
                    if (string.Equals(ReadText(worksheet, headerRow, column + 1), HeaderNorm, StringComparison.Ordinal) &&
                        string.Equals(ReadText(worksheet, headerRow, column + 2), HeaderKind, StringComparison.Ordinal) &&
                        string.Equals(ReadText(worksheet, headerRow, column + 3), HeaderFingerprint, StringComparison.Ordinal))
                    {
                        return column;
                    }
                }
                return 0;
            }
            finally
            {
                Release(used);
            }
        }

        private static int FindSafeTechnicalStartColumn(
            Excel.Worksheet worksheet,
            int preferredStart)
        {
            Excel.Range used = null;
            try
            {
                used = worksheet.UsedRange;
                int usedLast = Math.Max(1, used.Column + used.Columns.Count - 1);
                int start = Math.Max(preferredStart, usedLast + 1);
                if (start + 3 > 16384)
                    throw new InvalidOperationException("Khong con cot trong de luu metadata Du toan V2.");
                return start;
            }
            finally
            {
                Release(used);
            }
        }

        private static void WriteTechnicalHeader(
            Excel.Worksheet worksheet,
            int row,
            int column,
            string value)
        {
            WriteCellValue(worksheet, row, column, value);
        }

        private static void HideColumns(
            Excel.Worksheet worksheet,
            int firstColumn,
            int lastColumn)
        {
            Excel.Range range = null;
            try
            {
                range = worksheet.Range[
                    worksheet.Columns[firstColumn],
                    worksheet.Columns[lastColumn]];
                range.EntireColumn.Hidden = true;
            }
            finally
            {
                Release(range);
            }
        }

        private static string ReadText(
            Excel.Worksheet worksheet,
            int row,
            int column)
        {
            Excel.Range cell = null;
            try
            {
                cell = worksheet.Cells[row, column] as Excel.Range;
                object value = cell?.Value2;
                return (Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty).Trim();
            }
            finally
            {
                Release(cell);
            }
        }

        private static void WriteTextIfDifferent(
            Excel.Worksheet worksheet,
            int row,
            int column,
            string value,
            ref bool changed)
        {
            string current = ReadText(worksheet, row, column);
            string desired = value ?? string.Empty;
            if (string.Equals(current, desired, StringComparison.Ordinal))
                return;
            WriteCellValue(worksheet, row, column, desired);
            changed = true;
        }

        private static void WriteCellValue(
            Excel.Worksheet worksheet,
            int row,
            int column,
            object value)
        {
            Excel.Range cell = null;
            try
            {
                cell = worksheet.Cells[row, column] as Excel.Range;
                if (cell == null)
                    throw new InvalidOperationException("Khong truy cap duoc o Excel.");
                cell.Value2 = value;
            }
            finally
            {
                Release(cell);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
