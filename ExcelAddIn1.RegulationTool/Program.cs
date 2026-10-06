using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelAddIn1.Core;

namespace ExcelAddIn1.RegulationTool
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            try
            {
                Console.OutputEncoding = new UTF8Encoding(false);
                if (args.Length == 3 && string.Equals(args[0], "build", StringComparison.OrdinalIgnoreCase))
                {
                    Print(RegulationPackageBundleBuilder.Build(args[1], args[2]));
                    return 0;
                }
                if (args.Length == 2 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase))
                {
                    Print(RegulationPackageBundleReader.Read(args[1]));
                    return 0;
                }
                if (args.Length == 3 && string.Equals(args[0], "diff", StringComparison.OrdinalIgnoreCase))
                {
                    PrintDiff(
                        RegulationPackageBundleReader.Read(args[1]),
                        RegulationPackageBundleReader.Read(args[2]));
                    return 0;
                }
                if (args.Length == 4 && string.Equals(args[0], "keygen", StringComparison.OrdinalIgnoreCase))
                {
                    OfflineUpdateTrustedKey key = OfflineUpdatePublisher.CreateKey(
                        args[1], args[2], args[3]);
                    Console.WriteLine("KeyId=" + key.KeyId);
                    Console.WriteLine("PrivateKeyStore=Windows CAPI user key container");
                    Console.WriteLine("PublicKey=" + Path.GetFullPath(args[2]));
                    return 0;
                }
                if (args.Length == 6 && string.Equals(args[0], "build-update", StringComparison.OrdinalIgnoreCase))
                {
                    OfflineUpdateManifest manifest = OfflineUpdatePublisher.Build(
                        args[1], args[2], args[3], args[4], args[5]);
                    PrintUpdate(manifest, args[2]);
                    return 0;
                }
                if ((args.Length == 4 || args.Length == 5) &&
                    string.Equals(args[0], "verify-update", StringComparison.OrdinalIgnoreCase))
                {
                    OfflineUpdateTrustedKey key = OfflineUpdatePublicKeySerializer.Deserialize(
                        File.ReadAllText(args[2], new UTF8Encoding(false, true)));
                    IEnumerable<RegulationPackage> installed = args.Length == 5
                        ? new RegulationPackageStore(args[4]).ListInstalled()
                        : new RegulationPackage[0];
                    OfflineUpdateVerificationResult result = OfflineUpdatePackageVerifier.Verify(
                        args[1], new[] { key }, new Version(args[3]), installed);
                    PrintVerification(result);
                    return result.Disposition == OfflineUpdateDisposition.VersionConflict ||
                           result.Disposition == OfflineUpdateDisposition.DowngradeRejected ||
                           result.Disposition == OfflineUpdateDisposition.AppUpgradeRequired
                        ? 3
                        : 0;
                }

                Console.Error.WriteLine("Usage:");
                Console.Error.WriteLine("  ExcelAddIn1.RegulationTool build <source-directory> <bundle-directory>");
                Console.Error.WriteLine("  ExcelAddIn1.RegulationTool validate <bundle-directory>");
                Console.Error.WriteLine("  ExcelAddIn1.RegulationTool diff <old-bundle-directory> <new-bundle-directory>");
                Console.Error.WriteLine("  ExcelAddIn1.RegulationTool keygen <key-container> <public-key-file> <key-id>");
                Console.Error.WriteLine("  ExcelAddIn1.RegulationTool build-update <bundle-directory> <output.ttbupdate> <minimum-app-version> <key-container> <public-key-file>");
                Console.Error.WriteLine("  ExcelAddIn1.RegulationTool verify-update <file.ttbupdate> <public-key-file> <current-app-version> [package-store]");
                return 2;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.GetType().Name + ": " + ex.Message);
                return 1;
            }
        }

        private static void PrintUpdate(OfflineUpdateManifest manifest, string outputFile)
        {
            Console.WriteLine("UpdateId=" + manifest.UpdateId);
            Console.WriteLine("UpdateVersion=" + manifest.UpdateVersion);
            Console.WriteLine("Package=" + manifest.PackageId + "@" + manifest.PackageVersion);
            Console.WriteLine("PackageChecksum=" + manifest.PackageChecksum);
            Console.WriteLine("SigningKeyId=" + manifest.SigningKeyId);
            Console.WriteLine("Output=" + Path.GetFullPath(outputFile));
        }

        private static void PrintVerification(OfflineUpdateVerificationResult result)
        {
            Console.WriteLine("UpdateId=" + result.Manifest.UpdateId);
            Console.WriteLine("Package=" + result.Package.PackageId + "@" + result.Package.DataVersion);
            Console.WriteLine("ArchiveChecksum=" + result.ArchiveChecksum);
            Console.WriteLine("SigningKeyId=" + result.Manifest.SigningKeyId);
            Console.WriteLine("Disposition=" + result.Disposition);
            Console.WriteLine("Reason=" + result.DispositionReason);
        }

        private static void Print(RegulationPackageBundle bundle)
        {
            Console.WriteLine("PackageId=" + bundle.Package.PackageId);
            Console.WriteLine("DataVersion=" + bundle.Package.DataVersion);
            Console.WriteLine("PackageChecksum=" + bundle.Package.PackageChecksum);
            foreach (RegulationPackageModuleManifest module in bundle.Package.Modules)
            {
                Console.WriteLine(
                    "Module=" + module.Kind +
                    ";Records=" + module.RecordCount +
                    ";Checksum=" + module.ContentChecksum);
            }
        }

        private static void PrintDiff(RegulationPackageBundle source, RegulationPackageBundle target)
        {
            RegulationDataBundleDiffResult diff = RegulationDataBundleDiffer.Compare(source, target);
            Console.WriteLine("Source=" + source.Package.PackageId + "@" + source.Package.DataVersion);
            Console.WriteLine("Target=" + target.Package.PackageId + "@" + target.Package.DataVersion);
            Console.WriteLine(
                "RecordChanges=" + diff.Changes.Count +
                ";Added=" + diff.AddedCount +
                ";Removed=" + diff.RemovedCount +
                ";Changed=" + diff.ChangedCount);
            Console.WriteLine("Module\tKey\tChange\tField\tOldValue\tNewValue");
            foreach (RegulationDataRecordChange change in diff.Changes)
            {
                foreach (RegulationDataFieldChange field in change.FieldChanges)
                {
                    Console.WriteLine(string.Join("\t", new[]
                    {
                        change.ModuleKind.ToString(),
                        Escape(change.Key),
                        change.Kind.ToString(),
                        Escape(field.Field),
                        Escape(field.OldValue),
                        Escape(field.NewValue)
                    }));
                }
            }
        }

        private static string Escape(string value)
        {
            return (value ?? string.Empty)
                .Replace("\\", "\\\\")
                .Replace("\t", "\\t")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n");
        }
    }
}
