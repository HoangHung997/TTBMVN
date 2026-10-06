using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum RegulationUpdateProviderStatus
    {
        Disabled = 0,
        Available = 1,
        Unavailable = 2
    }

    public sealed class RegulationUpdateProviderResult
    {
        public RegulationUpdateProviderResult(
            RegulationUpdateProviderStatus status,
            string message,
            IEnumerable<string> packageLocations)
        {
            Status = status;
            Message = message ?? string.Empty;
            PackageLocations = new ReadOnlyCollection<string>(
                (packageLocations ?? Enumerable.Empty<string>()).ToList());
        }

        public RegulationUpdateProviderStatus Status { get; }
        public string Message { get; }
        public IReadOnlyList<string> PackageLocations { get; }
    }

    public interface IRegulationUpdateProvider
    {
        string ProviderId { get; }
        bool RequiresNetwork { get; }
        RegulationUpdateProviderResult CheckForUpdates();
    }

    public sealed class DisabledOnlineRegulationUpdateProvider : IRegulationUpdateProvider
    {
        public string ProviderId => "online-disabled";
        public bool RequiresNetwork => true;

        public RegulationUpdateProviderResult CheckForUpdates()
        {
            return new RegulationUpdateProviderResult(
                RegulationUpdateProviderStatus.Disabled,
                "Kenh cap nhat online chua duoc bat.",
                null);
        }
    }

    public sealed class OfflineUpdateInspection
    {
        internal OfflineUpdateInspection(
            OfflineUpdateVerificationResult verification,
            RegulationPackage sourcePackage,
            RegulationPackageDiffResult diff)
        {
            Verification = verification;
            SourcePackage = sourcePackage;
            Diff = diff;
        }

        public OfflineUpdateVerificationResult Verification { get; }
        public RegulationPackage SourcePackage { get; }
        public RegulationPackageDiffResult Diff { get; }
    }

    public enum OfflineUpdateInstallStatus
    {
        InstalledAndActivated = 0,
        AlreadyInstalledAndActivated = 1
    }

    public sealed class OfflineUpdateInstallResult
    {
        internal OfflineUpdateInstallResult(
            OfflineUpdateInstallStatus status,
            RegulationPackage package,
            string installDirectory)
        {
            Status = status;
            Package = package;
            InstallDirectory = installDirectory;
        }

        public OfflineUpdateInstallStatus Status { get; }
        public RegulationPackage Package { get; }
        public string InstallDirectory { get; }
    }

    public sealed class OfflineUpdatePolicyException : InvalidOperationException
    {
        public OfflineUpdatePolicyException(
            OfflineUpdateDisposition disposition,
            string message)
            : base(message)
        {
            Disposition = disposition;
        }

        public OfflineUpdateDisposition Disposition { get; }
    }

    public sealed class OfflineUpdateCenterService
    {
        private readonly RegulationPackageStore packageStore;
        private readonly RegulationPackageActivationStore activationStore;
        private readonly IReadOnlyList<OfflineUpdateTrustedKey> trustedKeys;
        private readonly Version currentAppVersion;

        public OfflineUpdateCenterService(
            string storeRoot,
            IEnumerable<OfflineUpdateTrustedKey> trustedKeys,
            Version currentAppVersion)
        {
            packageStore = new RegulationPackageStore(storeRoot);
            activationStore = new RegulationPackageActivationStore(storeRoot);
            this.trustedKeys = new ReadOnlyCollection<OfflineUpdateTrustedKey>(
                (trustedKeys ?? Enumerable.Empty<OfflineUpdateTrustedKey>()).ToList());
            if (this.trustedKeys.Count == 0)
                throw new ArgumentException("Trust store khong duoc rong.", nameof(trustedKeys));
            this.currentAppVersion = currentAppVersion ?? throw new ArgumentNullException(nameof(currentAppVersion));
        }

        public IReadOnlyList<RegulationPackage> ListInstalled()
        {
            return packageStore.ListInstalled();
        }

        public IReadOnlyList<RegulationPackage> ListPreferred()
        {
            return activationStore.SelectPreferred(ListInstalled());
        }

        public bool IsPreferred(RegulationPackage package)
        {
            IReadOnlyList<RegulationPackage> installed = ListInstalled();
            return activationStore.IsPreferred(package, installed);
        }

        public OfflineUpdateInspection Inspect(
            string updatePath,
            RegulationPackage currentWorkbookPackage)
        {
            OfflineUpdateVerificationResult verification = OfflineUpdatePackageVerifier.Verify(
                updatePath,
                trustedKeys,
                currentAppVersion,
                ListInstalled());
            RegulationPackageDiffResult diff = currentWorkbookPackage == null ||
                string.Equals(
                    currentWorkbookPackage.PackageChecksum,
                    verification.Package.PackageChecksum,
                    StringComparison.OrdinalIgnoreCase)
                ? null
                : RegulationPackageDiffer.Compare(currentWorkbookPackage, verification.Package);
            return new OfflineUpdateInspection(verification, currentWorkbookPackage, diff);
        }

        public OfflineUpdateInstallResult InstallAndActivate(OfflineUpdateInspection inspection)
        {
            if (inspection == null)
                throw new ArgumentNullException(nameof(inspection));
            OfflineUpdateVerificationResult verified = OfflineUpdatePackageVerifier.Verify(
                inspection.Verification.ArchivePath,
                trustedKeys,
                currentAppVersion,
                ListInstalled());
            if (!string.Equals(
                verified.ArchiveChecksum,
                inspection.Verification.ArchiveChecksum,
                StringComparison.OrdinalIgnoreCase))
            {
                throw new OfflineUpdateVerificationException(
                    OfflineUpdateVerificationFailure.PayloadMismatch,
                    "Goi update da thay doi sau khi xem truoc.");
            }
            if (verified.Disposition != OfflineUpdateDisposition.Ready &&
                verified.Disposition != OfflineUpdateDisposition.AlreadyInstalled)
            {
                throw new OfflineUpdatePolicyException(
                    verified.Disposition,
                    verified.DispositionReason);
            }

            if (verified.Disposition == OfflineUpdateDisposition.AlreadyInstalled)
            {
                RegulationPackage installed = ListInstalled().Single(package =>
                    SameIdentity(package, verified.Package));
                activationStore.SetPreferred(installed, ListInstalled());
                return new OfflineUpdateInstallResult(
                    OfflineUpdateInstallStatus.AlreadyInstalledAndActivated,
                    installed,
                    GetInstallDirectory(installed));
            }

            string extractionRoot = Path.Combine(
                packageStore.RootDirectory,
                ".update-extract",
                Guid.NewGuid().ToString("N"));
            try
            {
                OfflineUpdatePackageVerifier.ExtractVerifiedBundle(verified, extractionRoot);
                RegulationPackageInstallResult installed = packageStore.ImportFromDirectory(extractionRoot);
                IReadOnlyList<RegulationPackage> all = ListInstalled();
                activationStore.SetPreferred(installed.Package, all);
                return new OfflineUpdateInstallResult(
                    installed.Status == RegulationPackageInstallStatus.AlreadyInstalled
                        ? OfflineUpdateInstallStatus.AlreadyInstalledAndActivated
                        : OfflineUpdateInstallStatus.InstalledAndActivated,
                    installed.Package,
                    installed.InstallDirectory);
            }
            finally
            {
                if (Directory.Exists(extractionRoot))
                    Directory.Delete(extractionRoot, true);
                string parent = Path.GetDirectoryName(extractionRoot);
                if (Directory.Exists(parent) && !Directory.EnumerateFileSystemEntries(parent).Any())
                    Directory.Delete(parent);
            }
        }

        public void ActivateInstalled(RegulationPackage package)
        {
            IReadOnlyList<RegulationPackage> installed = ListInstalled();
            activationStore.SetPreferred(package, installed);
        }

        public RegulationUpdateProviderResult CheckProvider(IRegulationUpdateProvider provider)
        {
            if (provider == null)
                throw new ArgumentNullException(nameof(provider));
            try
            {
                return provider.CheckForUpdates();
            }
            catch (Exception ex)
            {
                return new RegulationUpdateProviderResult(
                    RegulationUpdateProviderStatus.Unavailable,
                    "Khong truy cap duoc kenh " + provider.ProviderId + ": " + ex.Message,
                    null);
            }
        }

        private string GetInstallDirectory(RegulationPackage package)
        {
            return Path.Combine(
                packageStore.RootDirectory,
                "packages",
                package.PackageId,
                package.DataVersion,
                package.PackageChecksum);
        }

        private static bool SameIdentity(RegulationPackage left, RegulationPackage right)
        {
            return string.Equals(left.PackageId, right.PackageId, StringComparison.Ordinal) &&
                   string.Equals(left.DataVersion, right.DataVersion, StringComparison.Ordinal) &&
                   string.Equals(left.PackageChecksum, right.PackageChecksum, StringComparison.OrdinalIgnoreCase);
        }
    }
}
