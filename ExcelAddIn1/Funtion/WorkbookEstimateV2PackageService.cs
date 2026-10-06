using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class EstimateV2PackagePreview
    {
        internal EstimateV2PackagePreview(EstimateV2PackageMigrationPlan plan, string profilePayload,
            ProjectProfile targetProfile, IEnumerable<string> differences)
        {
            Plan = plan; ProfilePayload = profilePayload; TargetProfile = targetProfile;
            Differences = differences.ToList().AsReadOnly();
        }
        public EstimateV2PackageMigrationPlan Plan { get; }
        public string ProfilePayload { get; }
        public ProjectProfile TargetProfile { get; }
        public IReadOnlyList<string> Differences { get; }
    }

    public static class WorkbookEstimateV2PackageService
    {
        public static IReadOnlyList<RegulationPackage> ListAvailable(out string[] issues)
        {
            var messages = new List<string>();
            try { RegulationPackageBootstrapService.LoadAvailablePackages(); }
            catch (Exception ex) { messages.Add(ex.Message); }
            var result = new List<RegulationPackage>();
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            string root = Path.Combine(store.RootDirectory, "packages");
            if (Directory.Exists(root))
                foreach (string package in Directory.GetDirectories(root))
                    foreach (string version in Directory.GetDirectories(package))
                        foreach (string checksum in Directory.GetDirectories(version))
                            try
                            {
                                result.Add(store.LoadRequired(Path.GetFileName(package),
                                    Path.GetFileName(version), Path.GetFileName(checksum)));
                            }
                            catch (Exception ex) { messages.Add(Path.GetFileName(package) + ": " + ex.Message); }
            issues = messages.Distinct().ToArray();
            return result.AsReadOnly();
        }

        public static EstimateV2PackagePreview Preview(Excel.Workbook workbook, RegulationPackage target,
            DateTime preparedDate, DateTime priceDate)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (target == null) throw new ArgumentNullException(nameof(target));
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            RegulationPackageBundle bundle = store.LoadBundleRequired(target.PackageId,
                target.DataVersion, target.PackageChecksum);
            EstimateV2State state = LoadState(workbook);
            ProjectProfile profile;
            bool hasProfile = WorkbookProjectProfileService.TryLoad(workbook, out profile);
            string payload;
            WorkbookProjectProfileService.TryReadPayload(workbook, out payload);
            var differences = new List<string>();
            if (hasProfile)
            {
                RegulationPackagePinVerificationResult source = RegulationPackagePinService.Verify(profile, store);
                if (source.IsAvailable)
                    differences.AddRange(RegulationPackageDiffer.Compare(source.Package, target).Changes
                        .Select(change => change.Key + ": " + change.OldValue + " → " + change.NewValue));
                else
                    differences.Add("Gói nguồn không đọc được: " + source.Message +
                        " Chỉ có thể đối chiếu mã định mức/variant; chưa so sánh được hao phí cũ.");
            }
            else
            {
                profile = new ProjectProfile
                {
                    ProjectId = Guid.NewGuid().ToString("N"),
                    PriceProfileId = "WORKBOOK-PRICES",
                    PreparedDate = preparedDate.Date,
                    PriceDate = priceDate.Date
                };
            }
            ProjectProfile next = RegulationPackagePinService.Pin(profile, target);
            if (next.PreparedDate < target.EffectiveFrom ||
                (target.EffectiveTo.HasValue && next.PreparedDate > target.EffectiveTo))
                differences.Add("Ngày lập hồ sơ nằm ngoài khoảng hiệu lực của gói; cần kiểm tra điều khoản chuyển tiếp.");
            differences.Add("Chuyển binding đang dùng sang gói đích; giữ nguyên mã định mức và variant.");
            differences.Add("Phải cập nhật VL-NC-M, DG và THKP sau khi đổi gói. Công thức cũ được giữ để đối chiếu.");
            return new EstimateV2PackagePreview(EstimateV2PackageMigrationPlan.Create(state, bundle),
                payload, next, differences);
        }

        public static string Apply(Excel.Workbook workbook, EstimateV2PackagePreview preview,
            bool confirmed, Action afterStateSaved = null)
        {
            if (!confirmed) return string.Empty;
            if (preview == null || !preview.Plan.CanApply)
                throw new InvalidOperationException("Preview còn lỗi định mức/variant; không thể đổi gói.");
            if (workbook.ReadOnly || string.IsNullOrEmpty(workbook.Path))
                throw new InvalidOperationException("Hãy lưu workbook có quyền ghi trước khi đổi gói để tạo bản dự phòng.");
            EstimateV2State before = LoadState(workbook);
            string profilePayload;
            WorkbookProjectProfileService.TryReadPayload(workbook, out profilePayload);
            if (!preview.Plan.Matches(before) || !string.Equals(profilePayload, preview.ProfilePayload, StringComparison.Ordinal))
                throw new InvalidOperationException("Workbook đã thay đổi sau preview. Hãy xem lại tác động.");
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            store.LoadBundleRequired(preview.Plan.Package.PackageId, preview.Plan.Package.DataVersion,
                preview.Plan.Package.PackageChecksum);
            string backupDirectory = Path.Combine(AppPaths.LocalAppDataRoot, "V2Backups");
            Directory.CreateDirectory(backupDirectory);
            string backup = Path.Combine(backupDirectory, "before-package-" + Guid.NewGuid().ToString("N") +
                Path.GetExtension(workbook.Name));
            workbook.SaveCopyAs(backup);
            bool hadState = HasState(workbook);
            try
            {
                using (new ExcelWriteContext(workbook.Application))
                {
                    WorkbookEstimateV2StateService.Save(workbook, preview.Plan.TargetState);
                    afterStateSaved?.Invoke();
                    WorkbookProjectProfileService.Save(workbook, preview.TargetProfile);
                }
            }
            catch (Exception changeError)
            {
                try
                {
                    if (hadState) WorkbookEstimateV2StateService.Save(workbook, before);
                    else ClearState(workbook);
                    WorkbookProjectProfileService.RestorePayload(workbook, profilePayload);
                }
                catch (Exception rollbackError)
                {
                    throw new AggregateException("Rollback thất bại. Bản dự phòng: " + backup,
                        changeError, rollbackError);
                }
                throw new InvalidOperationException("Đã rollback trạng thái và package. Bản dự phòng: " + backup, changeError);
            }
            return backup;
        }

        private static EstimateV2State LoadState(Excel.Workbook workbook)
        {
            EstimateV2State state;
            if (WorkbookEstimateV2StateService.TryLoad(workbook, out state)) return state;
            if (HasState(workbook)) throw new InvalidOperationException("Custom XML V2 bị hỏng; cần phục hồi trước khi đổi gói.");
            return EstimateV2State.Empty(DateTime.MinValue);
        }

        private static bool HasState(Excel.Workbook workbook)
        {
            Microsoft.Office.Core.CustomXMLParts parts = null;
            Microsoft.Office.Core.CustomXMLParts selected = null;
            try
            {
                parts = workbook.CustomXMLParts;
                selected = parts.SelectByNamespace(WorkbookEstimateV2StateService.CustomXmlNamespace);
                return selected.Count > 0;
            }
            finally { Release(selected); Release(parts); }
        }

        private static void ClearState(Excel.Workbook workbook)
        {
            Microsoft.Office.Core.CustomXMLParts parts = null;
            Microsoft.Office.Core.CustomXMLParts selected = null;
            try
            {
                parts = workbook.CustomXMLParts;
                selected = parts.SelectByNamespace(WorkbookEstimateV2StateService.CustomXmlNamespace);
                for (int i = selected.Count; i > 0; i--)
                {
                    Microsoft.Office.Core.CustomXMLPart part = selected[i];
                    try { part.Delete(); } finally { Release(part); }
                }
            }
            finally { Release(selected); Release(parts); }
        }
        private static void Release(object item)
        { if (item != null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item); }
    }
}
