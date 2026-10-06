using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelAddIn1.Funtion
{
    public static class RegulationPackageBootstrapService
    {
        private const string ProductPackagePrefix = "BQP-RPBM-";

        public static IReadOnlyList<RegulationPackage> LoadAvailablePackages()
        {
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            InstallBundledPackages(store);
            return new ReadOnlyCollection<RegulationPackage>(store.ListInstalled()
                .Where(package => package.PackageId.StartsWith(
                    ProductPackagePrefix,
                    StringComparison.OrdinalIgnoreCase))
                .ToList());
        }

        public static IReadOnlyList<RegulationPackage> LoadPreferredPackages()
        {
            IReadOnlyList<RegulationPackage> available = LoadAvailablePackages();
            return new RegulationPackageActivationStore(AppPaths.RegulationPackageDirectory)
                .SelectPreferred(available);
        }

        private static void InstallBundledPackages(RegulationPackageStore store)
        {
            string assemblyDirectory = Path.GetDirectoryName(
                Assembly.GetExecutingAssembly().Location);
            string root = Path.Combine(
                assemblyDirectory ?? AppDomain.CurrentDomain.BaseDirectory,
                "RegulationPackages");
            if (!Directory.Exists(root))
                return;

            foreach (string packageDirectory in Directory.GetDirectories(root))
            {
                string bundle = Path.Combine(packageDirectory, "bundle");
                if (!File.Exists(Path.Combine(
                    bundle,
                    RegulationPackageLayout.ManifestFileName)))
                {
                    continue;
                }

                RegulationPackage candidate = RegulationPackageBundleReader.Read(bundle).Package;
                if (!store.IsRemoved(candidate)) store.ImportFromDirectory(bundle);
            }
        }
    }
}
