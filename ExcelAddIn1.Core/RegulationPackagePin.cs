using System;
using System.IO;

namespace ExcelAddIn1.Core
{
    public enum RegulationPackagePinStatus
    {
        Available,
        UnpinnedLegacyProfile,
        Missing,
        Corrupt,
        IdentityMismatch,
        InvalidProfile
    }

    public sealed class RegulationPackagePinVerificationResult
    {
        internal RegulationPackagePinVerificationResult(
            RegulationPackagePinStatus status,
            RegulationPackage package,
            string message)
        {
            Status = status;
            Package = package;
            Message = message ?? string.Empty;
        }

        public bool IsAvailable => Status == RegulationPackagePinStatus.Available;
        public RegulationPackagePinStatus Status { get; }
        public RegulationPackage Package { get; }
        public string Message { get; }
    }

    public static class RegulationPackagePinService
    {
        public static bool IsPinned(ProjectProfile profile)
        {
            return profile != null &&
                   !string.IsNullOrWhiteSpace(profile.RegulationPackageVersion) &&
                   !string.IsNullOrWhiteSpace(profile.RegulationPackageChecksum);
        }

        public static ProjectProfile Pin(ProjectProfile profile, RegulationPackage package)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));
            RegulationPackageValidationResult packageValidation =
                RegulationPackageValidator.Validate(package);
            if (!packageValidation.IsValid)
                throw new ArgumentException(string.Join(" ", packageValidation.Errors), nameof(package));

            ProjectProfile pinned = profile.Clone();
            pinned.SchemaVersion = ProjectProfile.CurrentSchemaVersion;
            pinned.RegulationPackageId = package.PackageId;
            pinned.RegulationPackageVersion = package.DataVersion;
            pinned.RegulationPackageChecksum = package.PackageChecksum;
            ProjectProfileValidationResult profileValidation = ProjectProfileValidator.Validate(pinned);
            if (!profileValidation.IsValid)
                throw new ArgumentException(string.Join(" ", profileValidation.Errors), nameof(profile));
            return pinned;
        }

        public static RegulationPackagePinVerificationResult Verify(
            ProjectProfile profile,
            RegulationPackageStore store)
        {
            if (store == null)
                throw new ArgumentNullException(nameof(store));
            ProjectProfileValidationResult validation = ProjectProfileValidator.Validate(profile);
            if (!validation.IsValid)
            {
                return new RegulationPackagePinVerificationResult(
                    RegulationPackagePinStatus.InvalidProfile,
                    null,
                    string.Join(" ", validation.Errors));
            }
            if (!IsPinned(profile))
            {
                return new RegulationPackagePinVerificationResult(
                    RegulationPackagePinStatus.UnpinnedLegacyProfile,
                    null,
                    "Workbook cu chua pin package version/checksum; can chon va pin truoc khi tinh.");
            }

            RegulationPackage package;
            try
            {
                package = store.LoadRequired(
                    profile.RegulationPackageId,
                    profile.RegulationPackageVersion,
                    profile.RegulationPackageChecksum);
            }
            catch (DirectoryNotFoundException)
            {
                return new RegulationPackagePinVerificationResult(
                    RegulationPackagePinStatus.Missing,
                    null,
                    "May nay chua cai dung package da pin trong workbook.");
            }
            catch (FileNotFoundException)
            {
                return new RegulationPackagePinVerificationResult(
                    RegulationPackagePinStatus.Corrupt,
                    null,
                    "Package da pin bi thieu file.");
            }
            catch (InvalidDataException ex)
            {
                return new RegulationPackagePinVerificationResult(
                    RegulationPackagePinStatus.Corrupt,
                    null,
                    "Package da pin khong qua kiem tra: " + ex.Message);
            }
            catch (IOException ex)
            {
                return new RegulationPackagePinVerificationResult(
                    RegulationPackagePinStatus.Corrupt,
                    null,
                    "Khong doc duoc package da pin: " + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                return new RegulationPackagePinVerificationResult(
                    RegulationPackagePinStatus.Corrupt,
                    null,
                    "Khong co quyen doc package da pin: " + ex.Message);
            }

            if (!string.Equals(package.PackageId, profile.RegulationPackageId, StringComparison.Ordinal) ||
                !string.Equals(package.DataVersion, profile.RegulationPackageVersion, StringComparison.Ordinal) ||
                !string.Equals(
                    package.PackageChecksum,
                    profile.RegulationPackageChecksum,
                    StringComparison.OrdinalIgnoreCase))
            {
                return new RegulationPackagePinVerificationResult(
                    RegulationPackagePinStatus.IdentityMismatch,
                    null,
                    "Package load duoc khong khop identity da pin.");
            }

            return new RegulationPackagePinVerificationResult(
                RegulationPackagePinStatus.Available,
                package,
                "Da nap dung package version/checksum ma workbook da pin.");
        }
    }
}
