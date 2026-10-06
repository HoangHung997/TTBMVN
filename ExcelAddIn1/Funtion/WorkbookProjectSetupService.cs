using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public enum ProjectSetupCommitPhase
    {
        ProfileSaved
    }

    public sealed class ProjectSetupCommitResult
    {
        internal ProjectSetupCommitResult(bool profileChanged)
        {
            ProfileChanged = profileChanged;
        }

        public bool ProfileChanged { get; }
    }

    public static class WorkbookProjectSetupService
    {
        public static ProjectSetupCommitResult Commit(
            Excel.Workbook workbook,
            ProjectSetupPlan plan)
        {
            return Commit(workbook, plan, null);
        }

        public static ProjectSetupCommitResult Commit(
            Excel.Workbook workbook,
            ProjectSetupPlan plan,
            Action<ProjectSetupCommitPhase> phaseCallback)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (plan == null)
                throw new ArgumentNullException(nameof(plan));

            string previousPayload;
            bool hadProfile = WorkbookProjectProfileService.TryReadPayload(
                workbook,
                out previousPayload);
            IReadOnlyList<WorksheetRoleAssignment> previousAssignments =
                WorksheetRoleService.ReadAssignments(workbook);
            bool profileChanged = false;

            using (new ExcelWriteContext(workbook.Application))
            {
                try
                {
                    profileChanged = WorkbookProjectProfileService.Save(
                        workbook,
                        plan.Profile);
                    phaseCallback?.Invoke(ProjectSetupCommitPhase.ProfileSaved);
                    WorksheetRoleService.ApplyMapping(workbook, plan.SheetMapping);
                    return new ProjectSetupCommitResult(profileChanged);
                }
                catch (Exception commitException)
                {
                    Exception rollbackException = null;
                    try
                    {
                        WorksheetRoleService.RestoreAssignments(
                            workbook,
                            previousAssignments);
                        WorkbookProjectProfileService.RestorePayload(
                            workbook,
                            hadProfile ? previousPayload : null);
                    }
                    catch (Exception ex)
                    {
                        rollbackException = ex;
                        RuntimeLogger.Log(ex, "Rollback project setup");
                    }

                    if (rollbackException != null)
                    {
                        throw new AggregateException(
                            "Khong luu duoc project setup va rollback khong hoan tat.",
                            commitException,
                            rollbackException);
                    }
                    throw;
                }
            }
        }
    }
}
