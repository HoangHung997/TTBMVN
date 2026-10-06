using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class WorksheetRoleService
    {
        public const string PropertyName = "TTBMVN.WorksheetRole";

        public static string GetRoleId(Excel.Worksheet worksheet)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            Excel.CustomProperties properties = null;
            try
            {
                properties = worksheet.CustomProperties;
                for (int index = 1; index <= properties.Count; index++)
                {
                    Excel.CustomProperty property = null;
                    try
                    {
                        property = properties.Item[index];
                        if (string.Equals(property.Name, PropertyName, StringComparison.OrdinalIgnoreCase))
                            return Convert.ToString(property.Value) ?? string.Empty;
                    }
                    finally
                    {
                        ReleaseComObject(property);
                    }
                }

                return string.Empty;
            }
            finally
            {
                ReleaseComObject(properties);
            }
        }

        public static bool TryGetRole(Excel.Worksheet worksheet, out WorksheetRole role)
        {
            return WorksheetRoleCatalog.TryParse(GetRoleId(worksheet), out role);
        }

        public static void SetRole(Excel.Worksheet worksheet, WorksheetRole role)
        {
            SetRoleId(worksheet, WorksheetRoleCatalog.ToId(role));
        }

        public static void SetRoleId(Excel.Worksheet worksheet, string roleId)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            WorksheetRole parsedRole;
            if (!WorksheetRoleCatalog.TryParse(roleId, out parsedRole))
                throw new ArgumentException("Vai tro sheet khong hop le: " + roleId, nameof(roleId));

            SetRolePropertyValue(worksheet, WorksheetRoleCatalog.ToId(parsedRole));
        }

        private static void SetRolePropertyValue(Excel.Worksheet worksheet, string roleId)
        {
            Excel.CustomProperties properties = null;
            Excel.CustomProperty firstMatch = null;
            try
            {
                properties = worksheet.CustomProperties;
                for (int index = properties.Count; index >= 1; index--)
                {
                    Excel.CustomProperty property = null;
                    try
                    {
                        property = properties.Item[index];
                        if (!string.Equals(property.Name, PropertyName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        if (firstMatch == null)
                        {
                            firstMatch = property;
                            property = null;
                        }
                        else
                        {
                            property.Delete();
                        }
                    }
                    finally
                    {
                        ReleaseComObject(property);
                    }
                }

                if (firstMatch == null)
                    firstMatch = properties.Add(PropertyName, roleId);
                else
                    firstMatch.Value = roleId;
            }
            finally
            {
                ReleaseComObject(firstMatch);
                ReleaseComObject(properties);
            }
        }

        public static void ClearRole(Excel.Worksheet worksheet)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            Excel.CustomProperties properties = null;
            try
            {
                properties = worksheet.CustomProperties;
                for (int index = properties.Count; index >= 1; index--)
                {
                    Excel.CustomProperty property = null;
                    try
                    {
                        property = properties.Item[index];
                        if (string.Equals(property.Name, PropertyName, StringComparison.OrdinalIgnoreCase))
                            property.Delete();
                    }
                    finally
                    {
                        ReleaseComObject(property);
                    }
                }
            }
            finally
            {
                ReleaseComObject(properties);
            }
        }

        public static IReadOnlyList<WorksheetRoleAssignment> ReadAssignments(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            var assignments = new List<WorksheetRoleAssignment>();
            Excel.Sheets worksheets = null;
            try
            {
                worksheets = workbook.Worksheets;
                for (int index = 1; index <= worksheets.Count; index++)
                {
                    Excel.Worksheet worksheet = null;
                    try
                    {
                        worksheet = worksheets.Item[index] as Excel.Worksheet;
                        if (worksheet == null)
                            continue;

                        string roleId = GetRoleId(worksheet);
                        if (roleId.Length > 0)
                        {
                            assignments.Add(new WorksheetRoleAssignment(
                                worksheet.CodeName,
                                worksheet.Name,
                                roleId));
                        }
                    }
                    finally
                    {
                        ReleaseComObject(worksheet);
                    }
                }
            }
            finally
            {
                ReleaseComObject(worksheets);
            }

            return assignments.AsReadOnly();
        }

        public static WorksheetRoleValidationResult Validate(Excel.Workbook workbook)
        {
            return WorksheetRoleValidator.Validate(ReadAssignments(workbook));
        }

        public static Excel.Worksheet ResolveWorksheetRequired(
            Excel.Workbook workbook,
            WorksheetRole role)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            WorksheetRoleAssignment assignment = WorksheetRoleResolver.ResolveUnique(
                ReadAssignments(workbook),
                role);
            Excel.Sheets worksheets = null;
            try
            {
                worksheets = workbook.Worksheets;
                Excel.Worksheet nameFallback = null;
                for (int index = 1; index <= worksheets.Count; index++)
                {
                    Excel.Worksheet worksheet = null;
                    try
                    {
                        worksheet = worksheets.Item[index] as Excel.Worksheet;
                        if (worksheet == null)
                            continue;
                        if (string.Equals(worksheet.CodeName, assignment.SheetKey, StringComparison.OrdinalIgnoreCase))
                        {
                            Excel.Worksheet result = worksheet;
                            worksheet = null;
                            ReleaseComObject(nameFallback);
                            return result;
                        }
                        if (nameFallback == null &&
                            string.Equals(worksheet.Name, assignment.SheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            nameFallback = worksheet;
                            worksheet = null;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(worksheet);
                    }
                }
                if (nameFallback != null)
                    return nameFallback;
                throw new InvalidOperationException(
                    "Khong tim thay sheet cho role " + WorksheetRoleCatalog.ToId(role) + ".");
            }
            finally
            {
                ReleaseComObject(worksheets);
            }
        }

        public static WorksheetRoleValidationResult ApplyMapping(
            Excel.Workbook workbook,
            IEnumerable<WorksheetRoleMappingEntry> mapping)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            var mappingList = (mapping ?? Enumerable.Empty<WorksheetRoleMappingEntry>()).ToList();
            IReadOnlyList<WorkbookSheetDescriptor> sheets =
                WorkbookSheetChangeCoordinator.CaptureSnapshot(workbook);
            WorksheetRoleMappingValidationResult mappingValidation =
                WorksheetRoleMappingValidator.Validate(mappingList, sheets);
            if (!mappingValidation.IsValid)
                throw new ArgumentException(string.Join(" ", mappingValidation.Issues.Select(issue => issue.Message)), nameof(mapping));

            var oldRoles = new List<Tuple<string, string>>();
            Excel.Sheets worksheets = null;
            try
            {
                worksheets = workbook.Worksheets;
                for (int index = 1; index <= worksheets.Count; index++)
                {
                    Excel.Worksheet worksheet = null;
                    try
                    {
                        worksheet = worksheets.Item[index] as Excel.Worksheet;
                        if (worksheet == null)
                            continue;
                        string roleId = GetRoleId(worksheet);
                        if (roleId.Length > 0)
                            oldRoles.Add(Tuple.Create(worksheet.Name, roleId));
                    }
                    finally
                    {
                        ReleaseComObject(worksheet);
                    }
                }

                try
                {
                    ClearAllRoles(worksheets);
                    foreach (WorksheetRoleMappingEntry entry in mappingList)
                    {
                        WorkbookSheetDescriptor descriptor = sheets.First(sheet =>
                            string.Equals(sheet.Key, entry.SheetKey, StringComparison.OrdinalIgnoreCase));
                        Excel.Worksheet worksheet = null;
                        try
                        {
                            worksheet = worksheets.Item[descriptor.Name] as Excel.Worksheet;
                            SetRole(worksheet, entry.Role);
                        }
                        finally
                        {
                            ReleaseComObject(worksheet);
                        }
                    }

                    WorksheetRoleValidationResult result = Validate(workbook);
                    if (!result.IsValid)
                        throw new InvalidOperationException("Mapping da ghi nhung validation role khong dat.");
                    return result;
                }
                catch
                {
                    try
                    {
                        ClearAllRoles(worksheets);
                        foreach (Tuple<string, string> oldRole in oldRoles)
                        {
                            Excel.Worksheet worksheet = null;
                            try
                            {
                                worksheet = worksheets.Item[oldRole.Item1] as Excel.Worksheet;
                                SetRolePropertyValue(worksheet, oldRole.Item2);
                            }
                            finally
                            {
                                ReleaseComObject(worksheet);
                            }
                        }
                    }
                    catch (Exception rollbackException)
                    {
                        RuntimeLogger.Log(rollbackException, "Rollback worksheet role mapping");
                    }
                    throw;
                }
            }
            finally
            {
                ReleaseComObject(worksheets);
            }
        }

        public static void RestoreAssignments(
            Excel.Workbook workbook,
            IEnumerable<WorksheetRoleAssignment> assignments)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            var snapshot = (assignments ?? Enumerable.Empty<WorksheetRoleAssignment>())
                .Where(assignment => assignment != null)
                .ToList();
            Excel.Sheets worksheets = null;
            try
            {
                worksheets = workbook.Worksheets;
                ClearAllRoles(worksheets);
                foreach (WorksheetRoleAssignment assignment in snapshot)
                {
                    Excel.Worksheet worksheet = FindWorksheet(
                        worksheets,
                        assignment.SheetKey,
                        assignment.SheetName);
                    if (worksheet == null)
                    {
                        throw new InvalidOperationException(
                            "Khong the khoi phuc vai tro sheet vi sheet khong con ton tai: " +
                            assignment.SheetName + ".");
                    }

                    try
                    {
                        SetRolePropertyValue(worksheet, assignment.RoleId);
                    }
                    finally
                    {
                        ReleaseComObject(worksheet);
                    }
                }
            }
            finally
            {
                ReleaseComObject(worksheets);
            }
        }

        public static void ClearAssignments(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            Excel.Sheets worksheets = null;
            try
            {
                worksheets = workbook.Worksheets;
                ClearAllRoles(worksheets);
            }
            finally
            {
                ReleaseComObject(worksheets);
            }
        }

        public static Excel.Worksheet ResolveRequired(Excel.Workbook workbook, WorksheetRole role)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            string roleId = WorksheetRoleCatalog.ToId(role);
            Excel.Worksheet match = null;
            Excel.Sheets worksheets = null;
            try
            {
                worksheets = workbook.Worksheets;
                for (int index = 1; index <= worksheets.Count; index++)
                {
                    Excel.Worksheet worksheet = worksheets.Item[index] as Excel.Worksheet;
                    if (worksheet == null)
                        continue;

                    WorksheetRole currentRole;
                    if (!TryGetRole(worksheet, out currentRole) || currentRole != role)
                    {
                        ReleaseComObject(worksheet);
                        continue;
                    }

                    if (match != null)
                    {
                        ReleaseComObject(worksheet);
                        ReleaseComObject(match);
                        match = null;
                        throw new InvalidOperationException(
                            "Vai tro " + roleId + " dang bi gan trung cho nhieu sheet.");
                    }

                    match = worksheet;
                }
            }
            finally
            {
                ReleaseComObject(worksheets);
            }

            if (match == null)
                throw new InvalidOperationException("Khong tim thay sheet co vai tro " + roleId + ".");

            return match;
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private static void ClearAllRoles(Excel.Sheets worksheets)
        {
            for (int index = 1; index <= worksheets.Count; index++)
            {
                Excel.Worksheet worksheet = null;
                try
                {
                    worksheet = worksheets.Item[index] as Excel.Worksheet;
                    if (worksheet != null)
                        ClearRole(worksheet);
                }
                finally
                {
                    ReleaseComObject(worksheet);
                }
            }
        }

        private static Excel.Worksheet FindWorksheet(
            Excel.Sheets worksheets,
            string sheetKey,
            string sheetName)
        {
            for (int index = 1; index <= worksheets.Count; index++)
            {
                Excel.Worksheet worksheet = worksheets.Item[index] as Excel.Worksheet;
                if (worksheet == null)
                    continue;
                if ((!string.IsNullOrWhiteSpace(sheetKey) && string.Equals(
                        worksheet.CodeName,
                        sheetKey,
                        StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrWhiteSpace(sheetName) && string.Equals(
                        worksheet.Name,
                        sheetName,
                        StringComparison.OrdinalIgnoreCase)))
                {
                    return worksheet;
                }
                ReleaseComObject(worksheet);
            }
            return null;
        }
    }
}
