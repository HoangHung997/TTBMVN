using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ExcelAddIn1.Core;

namespace ExcelAddIn1.Tests
{
    internal static class Program
    {
        private static int failed;

        private static int Main()
        {
            Run("CalculateVolume", TestCalculateVolume);
            Run("CalculateHeightFromVolume", TestCalculateHeightFromVolume);
            Run("ExcelColumnAddress", TestExcelColumnAddress);
            Run("DaodatFeasibility", TestDaodatFeasibility);
            Run("ProductKeyCodec", TestProductKeyCodec);
            Run("WorksheetRoleCatalog", TestWorksheetRoleCatalog);
            Run("WorksheetRoleValidation", TestWorksheetRoleValidation);
            Run("WorksheetRoleResolver", TestWorksheetRoleResolver);
            Run("ProjectProfileRoundTrip", TestProjectProfileRoundTrip);
            Run("ProjectProfileMigration", TestProjectProfileMigration);
            Run("ProjectProfileValidation", TestProjectProfileValidation);
            Run("ProjectSetupPlanning", TestProjectSetupPlanning);
            Run("WorkbookSheetListState", TestWorkbookSheetListState);
            Run("WorksheetRoleMapping", TestWorksheetRoleMapping);
            Run("RegulationPackageRoundTrip", TestRegulationPackageRoundTrip);
            Run("RegulationPackageValidation", TestRegulationPackageValidation);
            Run("RegulationPackageChecksum", TestRegulationPackageChecksum);
            Run("RegulationDataModuleRoundTrip", TestRegulationDataModuleRoundTrip);
            Run("RegulationDataModuleValidation", TestRegulationDataModuleValidation);
            Run("Bqp2021PackageBundle", TestBqp2021PackageBundle);
            Run("Bqp2021PackageRejectsUnverified", TestBqp2021PackageRejectsUnverified);
            Run("Bqp2025PackageBundle", TestBqp2025PackageBundle);
            Run("Bqp2021To2025RecordDiff", TestBqp2021To2025RecordDiff);
            Run("NormCatalog", TestNormCatalog);
            Run("NormSearch", TestNormSearch);
            Run("NormCalculation", TestNormCalculation);
            Run("NormConditionsAndValidation", TestNormConditionsAndValidation);
            Run("CostRuleCatalog", TestCostRuleCatalog);
            Run("CostRuleLegacyCatalog", TestCostRuleLegacyCatalog);
            Run("CostRuleCalculation", TestCostRuleCalculation);
            Run("CostRuleEngine", TestCostRuleEngine);
            Run("CostRuleLocaleAndValidation", TestCostRuleLocaleAndValidation);
            Run("PriceProfileRoundTrip", TestPriceProfileRoundTrip);
            Run("PriceProfileValidationAndCoverage", TestPriceProfileValidationAndCoverage);
            Run("PriceProfileMachineAdapter", TestPriceProfileMachineAdapter);
            Run("PriceProfileStore", TestPriceProfileStore);
            Run("PriceProfilePortfolioRoundTrip", TestPriceProfilePortfolioRoundTrip);
            Run("UnitRateCalculation", TestUnitRateCalculation);
            Run("UnitRateConditionsAndBindings", TestUnitRateConditionsAndBindings);
            Run("UnitRateValidation", TestUnitRateValidation);
            Run("EstimateAppendixCalculation", TestEstimateAppendixCalculation);
            Run("EstimateAppendixValidation", TestEstimateAppendixValidation);
            Run("EstimateWorkspaceRoundTrip", TestEstimateWorkspaceRoundTrip);
            Run("EstimateV2StateRoundTrip", TestEstimateV2StateRoundTrip);
            Run("EstimateV2Fingerprint", TestEstimateV2Fingerprint);
            Run("EstimateRateGrouping", TestEstimateRateGrouping);
            Run("CostSummaryCalculation", TestCostSummaryCalculation);
            Run("CostSummaryValidation", TestCostSummaryValidation);
            Run("ResultAuditRoundTrip", TestResultAuditRoundTrip);
            Run("ResultAuditValidation", TestResultAuditValidation);
            Run("ResultAuditLookup", TestResultAuditLookup);
            Run("WorkbookValidationReport", TestWorkbookValidationReport);
            Run("WorkbookValidationInherited", TestWorkbookValidationInherited);
            Run("VietnameseMoneyWords", TestVietnameseMoneyWords);
            Run("MachineRateCatalog", TestMachineRateCatalog);
            Run("MachineRateCalculation", TestMachineRateCalculation);
            Run("MachineRateLocaleAndValidation", TestMachineRateLocaleAndValidation);
            Run("RegulationPackageStoreInstall", TestRegulationPackageStoreInstall);
            Run("RegulationPackageStoreLoadBundle", TestRegulationPackageStoreLoadBundle);
            Run("RegulationPackageStoreRejectsInvalid", TestRegulationPackageStoreRejectsInvalid);
            Run("RegulationPackageStoreRollback", TestRegulationPackageStoreRollback);
            Run("RegulationPackageEffectiveDateResolver", TestRegulationPackageEffectiveDateResolver);
            Run("RegulationPackageTransitionResolver", TestRegulationPackageTransitionResolver);
            Run("RegulationPackagePinning", TestRegulationPackagePinning);
            Run("RegulationPackageDiffAndPlan", TestRegulationPackageDiffAndPlan);
            Run("OfflineUpdateValidPackage", TestOfflineUpdateValidPackage);
            Run("OfflineUpdateRejectsTamper", TestOfflineUpdateRejectsTamper);
            Run("OfflineUpdateRejectsWrongSignature", TestOfflineUpdateRejectsWrongSignature);
            Run("OfflineUpdateRejectsDowngrade", TestOfflineUpdateRejectsDowngrade);
            Run("OfflineUpdateDuplicateVersionPolicy", TestOfflineUpdateDuplicateVersionPolicy);
            Run("RegulationPackageActivationRollback", TestRegulationPackageActivationRollback);
            Run("OfflineUpdateInstallAndRestart", TestOfflineUpdateInstallAndRestart);
            Run("OfflineUpdateProviderOfflineOnly", TestOfflineUpdateProviderOfflineOnly);

            if (failed == 0)
            {
                Console.WriteLine("All tests passed.");
                return 0;
            }

            Console.Error.WriteLine(failed + " test(s) failed.");
            return 1;
        }

        private static void Run(string name, Action test)
        {
            try
            {
                test();
                Console.WriteLine("[PASS] " + name);
            }
            catch (Exception ex)
            {
                failed++;
                Console.Error.WriteLine("[FAIL] " + name + ": " + ex.Message);
            }
        }

        private static void TestCalculateVolume()
        {
            double volume = DaodatMath.CalculateVolume(2, 1.5, 1.2, 0.8, 3);
            AssertNear(5.6571, volume, 0.0001);
        }

        private static void TestCalculateHeightFromVolume()
        {
            double height = DaodatMath.CalculateHeightFromVolume(10, 2, 1.5, 1.2, 0.8);
            double volume = DaodatMath.CalculateVolume(2, 1.5, 1.2, 0.8, height);
            AssertNear(10, volume, 0.0001);
        }

        private static void TestExcelColumnAddress()
        {
            AssertEqual(1, ExcelColumnAddress.ToNumber("A"));
            AssertEqual(26, ExcelColumnAddress.ToNumber("Z"));
            AssertEqual(27, ExcelColumnAddress.ToNumber("AA"));
            AssertEqual("A", ExcelColumnAddress.ToLetters(1));
            AssertEqual("XFD", ExcelColumnAddress.ToLetters(16384));
            AssertEqual("AB", ExcelColumnAddress.Normalize(" ab "));
            AssertFalse(ExcelColumnAddress.IsValid("A1"));
            AssertThrows<ArgumentOutOfRangeException>(() => ExcelColumnAddress.ToLetters(0));
        }

        private static void TestDaodatFeasibility()
        {
            AssertTrue(DaodatFeasibility.Analyze(5, 50, 5, 15).CanGenerate);
            AssertEqual(DaodatFeasibilityStatus.BelowMinimum, DaodatFeasibility.Analyze(5, 10, 5, 15).Status);
            AssertEqual(DaodatFeasibilityStatus.AboveMaximum, DaodatFeasibility.Analyze(5, 100, 5, 15).Status);
            AssertEqual(DaodatFeasibilityStatus.InvalidInput, DaodatFeasibility.Analyze(0, 10, 5, 15).Status);
        }

        private static void TestProductKeyCodec()
        {
            string machineId = "A8F3-C21D-9B04-7A6E";
            DateTime expiry = new DateTime(2027, 5, 26);
            var parameters = new CngKeyCreationParameters
            {
                ExportPolicy = CngExportPolicies.AllowPlaintextExport
            };
            using (CngKey signingKey = CngKey.Create(CngAlgorithm.ECDsaP256, null, parameters))
            {
                byte[] privateKey = signingKey.Export(CngKeyBlobFormat.EccPrivateBlob);
                byte[] publicKey = signingKey.Export(CngKeyBlobFormat.EccPublicBlob);
                string key = ProductKeyCodec.Generate(machineId, expiry, privateKey);

                AssertTrue(ProductKeyCodec.TryValidate(key, machineId, publicKey, out DateTime actualExpiry));
                AssertEqual(expiry, actualExpiry);
                AssertFalse(ProductKeyCodec.TryValidate(key, "1111-2222-3333-4444", publicKey, out _));
                AssertFalse(ProductKeyCodec.TryValidate(
                    key.Substring(0, key.Length - 1) + "A",
                    machineId,
                    publicKey,
                    out _));
                AssertFalse(key.Contains("2027"));
                AssertFalse(key.Contains("0526"));
            }
        }

        private static void TestWorksheetRoleCatalog()
        {
            AssertEqual(7, WorksheetRoleCatalog.All.Count);
            foreach (WorksheetRole expected in WorksheetRoleCatalog.All)
            {
                string roleId = WorksheetRoleCatalog.ToId(expected);
                AssertTrue(WorksheetRoleCatalog.TryParse(roleId.ToLowerInvariant(), out WorksheetRole actual));
                AssertEqual(expected, actual);
            }

            AssertFalse(WorksheetRoleCatalog.TryParse("FutureRole", out _));
        }

        private static void TestWorksheetRoleValidation()
        {
            List<WorksheetRoleAssignment> valid = CreateValidRoleAssignments();
            AssertTrue(WorksheetRoleValidator.Validate(valid).IsValid);

            var duplicate = new List<WorksheetRoleAssignment>(valid)
            {
                new WorksheetRoleAssignment("Sheet8", "Bang gia khac", "ResourcePrices")
            };
            WorksheetRoleValidationResult duplicateResult = WorksheetRoleValidator.Validate(duplicate);
            AssertFalse(duplicateResult.IsValid);
            AssertHasIssue(duplicateResult, WorksheetRoleIssueCode.DuplicateRole, "ResourcePrices");

            var missing = new List<WorksheetRoleAssignment>(valid);
            missing.RemoveAt(missing.Count - 1);
            WorksheetRoleValidationResult missingResult = WorksheetRoleValidator.Validate(missing);
            AssertHasIssue(missingResult, WorksheetRoleIssueCode.MissingRole, "CostRuleView");

            var unknown = new List<WorksheetRoleAssignment>(valid)
            {
                new WorksheetRoleAssignment("Sheet8", "Sheet la", "FutureRole")
            };
            WorksheetRoleValidationResult unknownResult = WorksheetRoleValidator.Validate(unknown);
            AssertHasIssue(unknownResult, WorksheetRoleIssueCode.UnknownRole, "FutureRole");

            var duplicateSheet = new List<WorksheetRoleAssignment>(valid)
            {
                new WorksheetRoleAssignment("Sheet1", "VL-NC-M", "UnitRateLand")
            };
            WorksheetRoleValidationResult duplicateSheetResult = WorksheetRoleValidator.Validate(duplicateSheet);
            AssertHasIssue(duplicateSheetResult, WorksheetRoleIssueCode.DuplicateSheet, string.Empty);
        }

        private static void TestWorksheetRoleResolver()
        {
            List<WorksheetRoleAssignment> valid = CreateValidRoleAssignments();
            WorksheetRoleAssignment resolved = WorksheetRoleResolver.ResolveUnique(
                valid,
                WorksheetRole.ResourcePrices);
            AssertEqual("Sheet1", resolved.SheetKey);

            var missing = new List<WorksheetRoleAssignment>(valid);
            missing.RemoveAt(0);
            AssertThrows<InvalidOperationException>(() => WorksheetRoleResolver.ResolveUnique(
                missing,
                WorksheetRole.ResourcePrices));

            var duplicate = new List<WorksheetRoleAssignment>(valid)
            {
                new WorksheetRoleAssignment("Sheet8", "Bang gia khac", "ResourcePrices")
            };
            AssertThrows<InvalidOperationException>(() => WorksheetRoleResolver.ResolveUnique(
                duplicate,
                WorksheetRole.ResourcePrices));
        }

        private static void TestProjectProfileRoundTrip()
        {
            ProjectProfile expected = CreateTestProjectProfile();
            string first = ProjectProfileSerializer.Serialize(expected);
            string second = ProjectProfileSerializer.Serialize(expected.Clone());
            AssertEqual(first, second);
            AssertEqual(64, ProjectProfileSerializer.ComputeChecksum(first).Length);

            ProjectProfile actual = ProjectProfileSerializer.Deserialize(first);
            AssertProjectProfileEqual(expected, actual);
        }

        private static void TestProjectProfileMigration()
        {
            string legacy = string.Join("\n", new[]
            {
                "TTBMVN_PROJECT_PROFILE",
                "schema=0",
                "project=TEVHQUNZLURPQw==",
                "created=2026-08-01",
                "approved=2026-08-02",
                "price=2026-07-15",
                "package=UEFDS0FHRS1MRUdBQ1k=",
                "priceProfile=UFJJQ0UtTEVHQUNZ",
                "overrides="
            });

            ProjectProfile migrated = ProjectProfileSerializer.Deserialize(legacy);
            AssertEqual(ProjectProfile.CurrentSchemaVersion, migrated.SchemaVersion);
            AssertEqual("LEGACY-DOC", migrated.ProjectId);
            AssertEqual("PACKAGE-LEGACY", migrated.RegulationPackageId);
            AssertEqual(string.Empty, migrated.RegulationPackageVersion);
            AssertEqual(string.Empty, migrated.RegulationPackageChecksum);
            AssertEqual(new DateTime(2026, 8, 2), migrated.ApprovalDate.Value);

            string schema1 = string.Join("\n", new[]
            {
                "TTBMVN_PROJECT_PROFILE",
                "schema=1",
                "projectId=VEVTVC1TQ0hFTUEx",
                "preparedDate=2026-08-01",
                "approvalDate=",
                "priceDate=2026-07-15",
                "regulationPackageId=VEVTVC1QQUNLQUdF",
                "priceProfileId=VEVTVC1QUklDRQ==",
                "overrideSummary="
            });
            ProjectProfile migratedSchema1 = ProjectProfileSerializer.Deserialize(schema1);
            AssertEqual(ProjectProfile.CurrentSchemaVersion, migratedSchema1.SchemaVersion);
            AssertEqual("TEST-SCHEMA1", migratedSchema1.ProjectId);
            AssertEqual(string.Empty, migratedSchema1.RegulationPackageVersion);
            AssertEqual(string.Empty, migratedSchema1.RegulationPackageChecksum);
            AssertThrows<InvalidDataException>(() =>
                ProjectProfileSerializer.Deserialize(schema1 + "\nunknown=value"));
        }

        private static void TestProjectProfileValidation()
        {
            ProjectProfile invalid = CreateTestProjectProfile();
            invalid.ProjectId = " ";
            invalid.ApprovalDate = invalid.PreparedDate.Value.AddDays(-1);
            invalid.RegulationPackageChecksum = string.Empty;
            ProjectProfileValidationResult result = ProjectProfileValidator.Validate(invalid);
            AssertFalse(result.IsValid);
            AssertTrue(result.Errors.Count >= 2);
            AssertThrows<ArgumentException>(() => ProjectProfileSerializer.Serialize(invalid));
            AssertThrows<System.IO.InvalidDataException>(() =>
                ProjectProfileSerializer.Deserialize("TTBMVN_PROJECT_PROFILE\nschema=9"));
        }

        private static void TestProjectSetupPlanning()
        {
            IReadOnlyList<WorkbookSheetDescriptor> sheets = new[]
            {
                new WorkbookSheetDescriptor("S1", "VL-NC-M"),
                new WorkbookSheetDescriptor("S2", "DG Can"),
                new WorkbookSheetDescriptor("S3", "DG Nuoc"),
                new WorkbookSheetDescriptor("S4", "Gia DT TC"),
                new WorkbookSheetDescriptor("S5", "THKP-TC"),
                new WorkbookSheetDescriptor("S6", "Tracuu"),
                new WorkbookSheetDescriptor("S7", "ChiPhi")
            };
            IReadOnlyList<WorksheetRoleMappingEntry> mapping =
                WorksheetRoleMappingSuggester.Suggest(sheets);
            RegulationPackage package2021 = CreateResolverPackage(
                "BQP-RPBM-2021",
                "1.0.0",
                new DateTime(2021, 11, 5),
                new DateTime(2025, 10, 27),
                RegulationPackageStatus.Superseded,
                false);
            RegulationPackage package2025 = CreateResolverPackage(
                "BQP-RPBM-2025",
                "2.0.0",
                new DateTime(2025, 10, 28),
                null,
                RegulationPackageStatus.Published,
                true);

            var currentDraft = new ProjectSetupDraft(
                "PROJECT-401",
                new DateTime(2026, 8, 1),
                null,
                new DateTime(2026, 8, 3),
                new DateTime(2026, 7, 31),
                "PRICE-2026-07",
                string.Empty,
                mapping);
            ProjectSetupPlanningResult current = ProjectSetupPlanner.CreatePlan(
                currentDraft,
                sheets,
                new[] { package2021, package2025 },
                RegulationTransitionRuleCatalog.All);
            AssertTrue(current.IsValid);
            AssertEqual("BQP-RPBM-2025", current.Plan.Profile.RegulationPackageId);
            AssertEqual(package2025.PackageChecksum, current.Plan.Profile.RegulationPackageChecksum);

            var grandfatheredDraft = new ProjectSetupDraft(
                "PROJECT-401",
                new DateTime(2025, 10, 1),
                new DateTime(2025, 10, 20),
                new DateTime(2026, 8, 3),
                new DateTime(2025, 10, 20),
                "PRICE-2025-10",
                string.Empty,
                mapping);
            ProjectSetupPlanningResult grandfathered = ProjectSetupPlanner.CreatePlan(
                grandfatheredDraft,
                sheets,
                new[] { package2021, package2025 },
                RegulationTransitionRuleCatalog.All);
            AssertTrue(grandfathered.IsValid);
            AssertEqual("BQP-RPBM-2021", grandfathered.Plan.Profile.RegulationPackageId);
            AssertEqual(
                RegulationPackageResolutionCode.ApprovedBeforeTransition,
                grandfathered.PackageResolution.Code);

            var missingMapping = new List<WorksheetRoleMappingEntry>(mapping);
            missingMapping.RemoveAt(0);
            ProjectSetupPlanningResult invalidMapping = ProjectSetupPlanner.CreatePlan(
                new ProjectSetupDraft(
                    "PROJECT-401",
                    currentDraft.PreparedDate,
                    null,
                    currentDraft.EvaluationDate,
                    currentDraft.PriceDate,
                    currentDraft.PriceProfileId,
                    string.Empty,
                    missingMapping),
                sheets,
                new[] { package2021, package2025 },
                RegulationTransitionRuleCatalog.All);
            AssertTrue(!invalidMapping.IsValid);

            ProjectSetupPlanningResult invalidDates = ProjectSetupPlanner.CreatePlan(
                new ProjectSetupDraft(
                    "PROJECT-401",
                    new DateTime(2026, 8, 2),
                    new DateTime(2026, 8, 1),
                    new DateTime(2026, 8, 3),
                    new DateTime(2026, 8, 1),
                    "PRICE-2026-08",
                    string.Empty,
                    mapping),
                sheets,
                new[] { package2021, package2025 },
                RegulationTransitionRuleCatalog.All);
            AssertTrue(!invalidDates.IsValid);
        }

        private static void TestWorkbookSheetListState()
        {
            var initial = new[]
            {
                new WorkbookSheetDescriptor("Sheet1", "VL-NC-M"),
                new WorkbookSheetDescriptor("Sheet2", "DG Can")
            };
            string initialSignature = WorkbookSheetList.BuildSignature(initial);
            AssertEqual("DG Can", WorkbookSheetList.ResolveSelection(initial, "Sheet2", string.Empty).Name);

            var renamed = new[]
            {
                new WorkbookSheetDescriptor("Sheet1", "VL-NC-M"),
                new WorkbookSheetDescriptor("Sheet2", "Don gia can")
            };
            AssertFalse(string.Equals(
                initialSignature,
                WorkbookSheetList.BuildSignature(renamed),
                StringComparison.Ordinal));
            AssertEqual("Don gia can", WorkbookSheetList.ResolveSelection(
                renamed,
                "Sheet2",
                "DG Can").Name);

            var added = new[]
            {
                renamed[0],
                renamed[1],
                new WorkbookSheetDescriptor("Sheet3", "Sheet moi")
            };
            AssertFalse(string.Equals(
                WorkbookSheetList.BuildSignature(renamed),
                WorkbookSheetList.BuildSignature(added),
                StringComparison.Ordinal));

            var deleted = new[] { renamed[0] };
            AssertEqual("VL-NC-M", WorkbookSheetList.ResolveSelection(
                deleted,
                "Sheet2",
                "Don gia can").Name);
            AssertEqual(null, WorkbookSheetList.ResolveSelection(
                new WorkbookSheetDescriptor[0],
                "Sheet2",
                "Don gia can"));
        }

        private static void TestWorksheetRoleMapping()
        {
            var sheets = new[]
            {
                new WorkbookSheetDescriptor("S1", "VL-NC-M"),
                new WorkbookSheetDescriptor("S2", "DG Can"),
                new WorkbookSheetDescriptor("S3", "DG Nuoc"),
                new WorkbookSheetDescriptor("S4", "Gia DT TC"),
                new WorkbookSheetDescriptor("S5", "THKP-TC"),
                new WorkbookSheetDescriptor("S6", "Tracuu"),
                new WorkbookSheetDescriptor("S7", "ChiPhi")
            };
            IReadOnlyList<WorksheetRoleMappingEntry> suggested =
                WorksheetRoleMappingSuggester.Suggest(sheets);
            AssertTrue(WorksheetRoleMappingValidator.Validate(suggested, sheets).IsValid);

            IReadOnlyList<WorksheetRoleMappingEntry> renamedMapping =
                WorksheetRoleMappingBuilder.FromAssignments(
                    suggested.Select(entry => new WorksheetRoleAssignment(
                        entry.SheetKey,
                        "Ten da doi " + entry.SheetName,
                        WorksheetRoleCatalog.ToId(entry.Role))),
                    sheets.Select(sheet => new WorkbookSheetDescriptor(
                        sheet.Key,
                        "Ten da doi " + sheet.Name)));
            AssertTrue(WorksheetRoleMappingValidator.Validate(
                renamedMapping,
                sheets.Select(sheet => new WorkbookSheetDescriptor(
                    sheet.Key,
                    "Ten da doi " + sheet.Name))).IsValid);

            var duplicateSheet = new List<WorksheetRoleMappingEntry>(suggested);
            duplicateSheet[1] = new WorksheetRoleMappingEntry(
                WorksheetRole.UnitRateLand,
                "S1",
                "VL-NC-M");
            WorksheetRoleMappingValidationResult duplicateResult =
                WorksheetRoleMappingValidator.Validate(duplicateSheet, sheets);
            AssertFalse(duplicateResult.IsValid);
            AssertMappingIssue(duplicateResult, WorksheetRoleMappingIssueCode.SheetAssignedMultipleRoles);

            var missingSelection = new List<WorksheetRoleMappingEntry>(suggested);
            missingSelection[6] = new WorksheetRoleMappingEntry(
                WorksheetRole.CostRuleView,
                string.Empty,
                string.Empty);
            AssertMappingIssue(
                WorksheetRoleMappingValidator.Validate(missingSelection, sheets),
                WorksheetRoleMappingIssueCode.SheetNotSelected);

            var missingSheet = new List<WorkbookSheetDescriptor>(sheets);
            missingSheet.RemoveAt(6);
            AssertMappingIssue(
                WorksheetRoleMappingValidator.Validate(suggested, missingSheet),
                WorksheetRoleMappingIssueCode.SheetNotFound);
        }

        private static void TestRegulationPackageRoundTrip()
        {
            List<RegulationPackageModuleManifest> modules = CreateTestRegulationModules();
            modules.Reverse();
            RegulationPackage expected = RegulationPackage.Create(
                "BQP-RPBM-2021",
                "1.0.0",
                new DateTime(2021, 11, 5),
                new DateTime(2025, 10, 27),
                RegulationPackageStatus.Superseded,
                "Ap dung truoc goi hop nhat 2025.",
                CreateTestRegulationSources(),
                modules);
            modules.Clear();

            AssertEqual(6, expected.Modules.Count);
            AssertEqual(64, expected.PackageChecksum.Length);
            AssertTrue(RegulationPackageValidator.Validate(expected).IsValid);

            string first = RegulationPackageSerializer.Serialize(expected);
            RegulationPackage actual = RegulationPackageSerializer.Deserialize(first);
            string second = RegulationPackageSerializer.Serialize(actual);
            AssertEqual(first, second);
            AssertEqual(expected.PackageId, actual.PackageId);
            AssertEqual(expected.DataVersion, actual.DataVersion);
            AssertEqual(expected.EffectiveFrom, actual.EffectiveFrom);
            AssertEqual(expected.EffectiveTo, actual.EffectiveTo);
            AssertEqual(expected.Status, actual.Status);
            AssertEqual(expected.TransitionNote, actual.TransitionNote);
            AssertEqual(expected.PackageChecksum, actual.PackageChecksum);
            AssertEqual(3, actual.Sources.Count);
            AssertEqual(6, actual.Modules.Count);
            AssertEqual(expected.Sources[0].DocumentId, actual.Sources[0].DocumentId);
            AssertEqual(expected.Modules[0].ModuleId, actual.Modules[0].ModuleId);
        }

        private static void TestRegulationPackageValidation()
        {
            List<RegulationPackageSourceDocument> sources = CreateTestRegulationSources();
            List<RegulationPackageModuleManifest> modules = CreateTestRegulationModules();

            AssertThrows<ArgumentException>(() => RegulationPackage.Create(
                "bad id!",
                "1",
                new DateTime(2021, 11, 5, 1, 0, 0),
                new DateTime(2021, 1, 1),
                RegulationPackageStatus.Published,
                string.Empty,
                sources,
                modules));

            var missingModule = new List<RegulationPackageModuleManifest>(modules);
            missingModule.RemoveAt(missingModule.Count - 1);
            AssertThrows<ArgumentException>(() => RegulationPackage.Create(
                "BQP-RPBM-MISSING",
                "1.0.0",
                new DateTime(2021, 11, 5),
                null,
                RegulationPackageStatus.Published,
                string.Empty,
                sources,
                missingModule));

            var duplicateSources = new List<RegulationPackageSourceDocument>(sources)
            {
                sources[0]
            };
            AssertThrows<ArgumentException>(() => RegulationPackage.Create(
                "BQP-RPBM-DUPLICATE",
                "1.0.0",
                new DateTime(2021, 11, 5),
                null,
                RegulationPackageStatus.Published,
                string.Empty,
                duplicateSources,
                modules));

            var consolidatedSources = new List<RegulationPackageSourceDocument>(sources)
            {
                new RegulationPackageSourceDocument(
                    "VBHN95-2025-BQP",
                    "Van ban hop nhat quy trinh",
                    "Bo Quoc phong",
                    new DateTime(2025, 11, 25),
                    new DateTime(2025, 10, 28),
                    null,
                    "https://vbpl.vn/van-ban-hop-nhat/95-vbhn-bqp",
                    new string('D', 64))
            };
            RegulationPackage consolidated = RegulationPackage.Create(
                "BQP-RPBM-CONSOLIDATED",
                "2.0.0",
                new DateTime(2025, 10, 28),
                null,
                RegulationPackageStatus.Published,
                "Van ban hop nhat co the duoc ky sau ngay hieu luc cua noi dung hop nhat.",
                consolidatedSources,
                modules);
            AssertTrue(RegulationPackageValidator.Validate(consolidated).IsValid);

            RegulationPackage draft = RegulationPackage.Create(
                "BQP-RPBM-DRAFT",
                "0.1.0-preview",
                new DateTime(2027, 1, 1),
                null,
                RegulationPackageStatus.Draft,
                "Package du thao co the chua du module.",
                new RegulationPackageSourceDocument[0],
                new[] { modules[0] });
            AssertTrue(RegulationPackageValidator.Validate(draft).IsValid);
        }

        private static void TestRegulationPackageChecksum()
        {
            RegulationPackage package = CreateTestRegulationPackage();
            string payload = RegulationPackageSerializer.Serialize(package);
            string[] lines = payload.Split('\n');
            for (int index = 0; index < lines.Length; index++)
            {
                if (lines[index].StartsWith("transitionNote=", StringComparison.Ordinal))
                {
                    lines[index] = "transitionNote=VEFNUEVSRUQ=";
                    break;
                }
            }
            string tampered = string.Join("\n", lines);
            AssertThrows<System.IO.InvalidDataException>(() =>
                RegulationPackageSerializer.Deserialize(tampered));
            AssertThrows<System.IO.InvalidDataException>(() =>
                RegulationPackageSerializer.Deserialize(payload + "\nunknown=value"));
        }

        private static void TestRegulationDataModuleRoundTrip()
        {
            RegulationDataModule expected = CreateTestDataModule();
            string first = RegulationDataModuleSerializer.Serialize(expected);
            RegulationDataModule actual = RegulationDataModuleSerializer.Deserialize(first);
            string second = RegulationDataModuleSerializer.Serialize(actual);

            AssertEqual(first, second);
            AssertEqual(RegulationModuleKind.Norm, actual.Kind);
            AssertEqual("1.0.0", actual.DataVersion);
            AssertEqual(2, actual.Records.Count);
            AssertEqual("NORM-000.0100", actual.Records[0].Key);
            AssertEqual("TT123-2021-BQP", actual.Records[0].Source.DocumentId);
            AssertEqual(10, actual.Records[0].Source.PageFrom);
            AssertEqual(64, actual.ContentChecksum.Length);
        }

        private static void TestRegulationDataModuleValidation()
        {
            RegulationDataModule module = CreateTestDataModule();
            List<RegulationPackageModuleManifest> manifests = CreateTestRegulationModules();
            int normIndex = manifests.FindIndex(item => item.Kind == RegulationModuleKind.Norm);
            manifests[normIndex] = new RegulationPackageModuleManifest(
                RegulationModuleKind.Norm,
                "BQP-2021-NORM",
                RegulationDataModule.CurrentSchemaVersion,
                module.DataVersion,
                module.Records.Count,
                RegulationDataModuleSerializer.ComputeSerializedChecksum(module));
            RegulationPackage package = RegulationPackage.Create(
                "BQP-RPBM-DATA-TEST",
                "1.0.0",
                new DateTime(2021, 11, 5),
                null,
                RegulationPackageStatus.Published,
                "Data module validation test.",
                CreateTestRegulationSources(),
                manifests);
            AssertTrue(RegulationDataValidator.ValidateAgainstPackage(
                module,
                package,
                true).IsValid);

            var duplicate = new[] { module.Records[0], module.Records[0] };
            AssertThrows<ArgumentException>(() => RegulationDataModule.Create(
                RegulationModuleKind.Norm,
                "1.0.0",
                duplicate));
            AssertThrows<ArgumentException>(() => RegulationDataModule.Create(
                RegulationModuleKind.Norm,
                "1.0.0",
                new[]
                {
                    new RegulationDataRecord(
                        "NORM-UNVERIFIED",
                        "NormHeader",
                        "m2-10000",
                        "Unverified record",
                        string.Empty,
                        new RegulationSourceLocator("TT123-2021-BQP", 10, 10, "Phu luc I"),
                        RegulationDataVerification.Unverified)
                }));

            string payload = RegulationDataModuleSerializer.Serialize(module);
            AssertThrows<InvalidDataException>(() =>
                RegulationDataModuleSerializer.Deserialize(payload + "\nunknown=value"));

            RegulationPackage wrongSourcePackage = RegulationPackage.Create(
                "BQP-RPBM-DATA-WRONG-SOURCE",
                "1.0.0",
                new DateTime(2021, 11, 5),
                null,
                RegulationPackageStatus.Published,
                "Wrong source validation test.",
                new[]
                {
                    new RegulationPackageSourceDocument(
                        "OTHER-SOURCE",
                        "Other",
                        "Bo Quoc phong",
                        new DateTime(2021, 9, 20),
                        new DateTime(2021, 11, 5),
                        null,
                        "https://example.invalid/other",
                        new string('A', 64))
                },
                manifests);
            AssertFalse(RegulationDataValidator.ValidateAgainstPackage(
                module,
                wrongSourcePackage,
                true).IsValid);
        }

        private static RegulationDataModule CreateTestDataModule()
        {
            return RegulationDataModule.Create(
                RegulationModuleKind.Norm,
                "1.0.0",
                new[]
                {
                    new RegulationDataRecord(
                        "NORM-000.0200",
                        "NormHeader",
                        "m2-10000",
                        "Don mat bang khao sat",
                        "forestClasses=4",
                        new RegulationSourceLocator(
                            "TT123-2021-BQP",
                            10,
                            10,
                            "Phu luc I, Phan II, Chuong I, muc 2a"),
                        RegulationDataVerification.VerifiedAgainstOfficialSource),
                    new RegulationDataRecord(
                        "NORM-000.0100",
                        "NormHeader",
                        "commune",
                        "Dieu tra khu vuc o nhiem",
                        "geographyClasses=2",
                        new RegulationSourceLocator(
                            "TT123-2021-BQP",
                            10,
                            10,
                            "Phu luc I, Phan II, Chuong I, muc 1"),
                        RegulationDataVerification.VerifiedAgainstOfficialSource)
                });
        }

        private static void TestBqp2021PackageBundle()
        {
            string root = GetRepositoryRoot();
            string sourcePath = Path.Combine(
                root, "data", "regulations", "packages", "BQP-RPBM-2021", "source");
            string bundlePath = Path.Combine(
                root, "data", "regulations", "packages", "BQP-RPBM-2021", "bundle");
            RegulationPackageSourceDefinition source = RegulationPackageSourceReader.Read(sourcePath);
            RegulationPackageBundle checkedIn = RegulationPackageBundleReader.Read(bundlePath);

            AssertEqual("BQP-RPBM-2021", source.PackageId);
            AssertEqual("1.0.0", source.DataVersion);
            AssertEqual(new DateTime(2021, 11, 5), source.EffectiveFrom);
            AssertEqual(new DateTime(2025, 10, 27), source.EffectiveTo.Value);
            AssertEqual(RegulationPackageStatus.Superseded, source.Status);
            AssertEqual(3, source.Sources.Count);
            AssertEqual(6, source.Records.Count);
            AssertEqual(59, checkedIn.Modules[RegulationModuleKind.TechnicalProcess].Records.Count);
            AssertEqual(34, checkedIn.Modules[RegulationModuleKind.Norm].Records.Count);
            AssertEqual(34, checkedIn.Modules[RegulationModuleKind.CostRule].Records.Count);
            AssertEqual(33, checkedIn.Modules[RegulationModuleKind.MachineRate].Records.Count);
            AssertEqual(35, checkedIn.Modules[RegulationModuleKind.Geography].Records.Count);
            AssertEqual(3, checkedIn.Modules[RegulationModuleKind.Compliance].Records.Count);

            var maximumPages = new Dictionary<string, int>(StringComparer.Ordinal)
            {
                { "TT121-2021-BQP", 99 },
                { "TT122-2021-BQP", 28 },
                { "TT123-2021-BQP", 39 }
            };
            foreach (RegulationDataModule module in checkedIn.Modules.Values)
            {
                foreach (RegulationDataRecord record in module.Records)
                {
                    AssertEqual(
                        RegulationDataVerification.VerifiedAgainstOfficialSource,
                        record.Verification);
                    AssertTrue(maximumPages.ContainsKey(record.Source.DocumentId));
                    AssertTrue(record.Source.PageTo <= maximumPages[record.Source.DocumentId]);
                    AssertFalse(string.IsNullOrWhiteSpace(record.Data));
                }
            }

            RegulationDataModule processes = checkedIn.Modules[RegulationModuleKind.TechnicalProcess];
            AssertEqual(
                "Nghiệm thu bàn giao",
                FindRecord(processes, "PROC-TT121-DIEU-057").Title);

            RegulationDataModule norms = checkedIn.Modules[RegulationModuleKind.Norm];
            AssertEqual("m3-excavated", FindRecord(norms, "NORM-020.0800").Unit);
            AssertEqual("bomb", FindRecord(norms, "NORM-040.0600").Unit);
            AssertCodeRange(norms, "NORM-000.", 100, 400, 100);
            AssertCodeRange(norms, "NORM-010.", 100, 400, 100);
            AssertCodeRange(norms, "NORM-020.", 100, 1200, 100);
            AssertCodeRange(norms, "NORM-030.", 100, 800, 100);
            AssertCodeRange(norms, "NORM-040.", 100, 600, 100);

            RegulationDataModule machines = checkedIn.Modules[RegulationModuleKind.MachineRate];
            for (int index = 1; index <= 33; index++)
            {
                string key = "MACHINE-M010." + index.ToString("000", CultureInfo.InvariantCulture);
                RegulationDataRecord machine = FindRecord(machines, key);
                Dictionary<string, string> data = ParseRecordData(machine.Data);
                double annualShifts = ParseInvariantDouble(data["annualShifts"]);
                double depreciation = ParseInvariantDouble(data["depreciationPercent"]);
                double repair = ParseInvariantDouble(data["repairPercent"]);
                double other = ParseInvariantDouble(data["otherPercent"]);
                long price = long.Parse(data["referencePriceVnd"], CultureInfo.InvariantCulture);
                double vat = ParseInvariantDouble(data["recoverableVatPercent"]);
                AssertTrue(annualShifts > 0);
                AssertTrue(depreciation > 0 && repair >= 0 && other >= 0);
                AssertTrue(depreciation + repair + other < 100);
                AssertTrue(price > 0);
                AssertTrue(vat == 0 || vat == 10);
                AssertTrue(data.ContainsKey("fuel"));
                AssertTrue(data.ContainsKey("operators"));
            }
            AssertEqual(
                "566835000",
                ParseRecordData(FindRecord(machines, "MACHINE-M010.004").Data)["referencePriceVnd"]);
            AssertEqual(
                "890000",
                ParseRecordData(FindRecord(machines, "MACHINE-M010.025").Data)["referencePriceVnd"]);

            RegulationDataModule costs = checkedIn.Modules[RegulationModuleKind.CostRule];
            AssertEqual("40", ParseRecordData(FindRecord(costs, "COST-COMMON").Data)["ratePercent"]);
            Dictionary<string, string> k3 = ParseRecordData(FindRecord(costs, "COST-K3-LIMITS").Data);
            AssertEqual("2000000", k3["minimumVnd"]);
            AssertEqual("60000000", k3["maximumVnd"]);
            AssertEqual(
                "3.285,2.853,2.435,1.845,1.546,1.188,0.797,0.694",
                ParseRecordData(FindRecord(costs, "COST-K5-CIVIL").Data)["ratesPercent"]);

            RegulationDataModule geography = checkedIn.Modules[RegulationModuleKind.Geography];
            AssertEqual(
                "245",
                ParseRecordData(FindRecord(geography, "GEO-LAND-SIGNAL-ZONE-4").Data)["depth-0.3-or-0.5"]);
            AssertEqual(
                "1",
                ParseRecordData(FindRecord(geography, "GEO-WATER-SIGNAL-ZONE-4").Data)["depth-over-1-to-10"]);

            string temp = CreateTemporaryDirectory("bqp-2021-package");
            try
            {
                string rebuiltPath = Path.Combine(temp, "rebuilt");
                RegulationPackageBundle rebuilt = RegulationPackageBundleBuilder.Build(sourcePath, rebuiltPath);
                AssertEqual(checkedIn.Package.PackageChecksum, rebuilt.Package.PackageChecksum);
                AssertEqual(
                    File.ReadAllText(Path.Combine(bundlePath, RegulationPackageLayout.ManifestFileName)),
                    File.ReadAllText(Path.Combine(rebuiltPath, RegulationPackageLayout.ManifestFileName)));
                RegulationPackageBundle second = RegulationPackageBundleBuilder.Build(sourcePath, rebuiltPath);
                AssertEqual(rebuilt.Package.PackageChecksum, second.Package.PackageChecksum);

                var store = new RegulationPackageStore(Path.Combine(temp, "store"));
                RegulationPackageInstallResult installed = store.ImportFromDirectory(rebuiltPath);
                AssertEqual(RegulationPackageInstallStatus.Installed, installed.Status);
                AssertEqual(checkedIn.Package.PackageChecksum, installed.Package.PackageChecksum);
                AssertEqual(1, store.ListInstalled().Count);
            }
            finally
            {
                DeleteTemporaryDirectory(temp);
            }
        }

        private static void TestBqp2021PackageRejectsUnverified()
        {
            string root = GetRepositoryRoot();
            string sourcePath = Path.Combine(
                root, "data", "regulations", "packages", "BQP-RPBM-2021", "source");
            string temp = CreateTemporaryDirectory("bqp-2021-unverified");
            try
            {
                string copiedSource = Path.Combine(temp, "source");
                CopyDirectory(sourcePath, copiedSource);
                string technicalProcess = Path.Combine(copiedSource, "modules", "TechnicalProcess.tsv");
                string payload = File.ReadAllText(technicalProcess, Encoding.UTF8);
                payload = ReplaceFirst(
                    payload,
                    "VerifiedAgainstOfficialSource",
                    "Unverified");
                File.WriteAllText(technicalProcess, payload, new UTF8Encoding(false));
                AssertThrows<ArgumentException>(() => RegulationPackageBundleBuilder.Build(
                    copiedSource,
                    Path.Combine(temp, "bundle")));
            }
            finally
            {
                DeleteTemporaryDirectory(temp);
            }
        }

        private static void TestBqp2025PackageBundle()
        {
            string root = GetRepositoryRoot();
            string packageRoot = Path.Combine(
                root, "data", "regulations", "packages", "BQP-RPBM-2025");
            string sourcePath = Path.Combine(packageRoot, "source");
            string bundlePath = Path.Combine(packageRoot, "bundle");
            RegulationPackageSourceDefinition source = RegulationPackageSourceReader.Read(sourcePath);
            RegulationPackageBundle bundle = RegulationPackageBundleReader.Read(bundlePath);

            AssertEqual("BQP-RPBM-2025", source.PackageId);
            AssertEqual("2.0.1", source.DataVersion);
            AssertEqual(new DateTime(2025, 10, 28), source.EffectiveFrom);
            AssertFalse(source.EffectiveTo.HasValue);
            AssertEqual(RegulationPackageStatus.Published, source.Status);
            AssertEqual(6, source.Sources.Count);
            AssertEqual(60, bundle.Modules[RegulationModuleKind.TechnicalProcess].Records.Count);
            AssertEqual(42, bundle.Modules[RegulationModuleKind.Norm].Records.Count);
            AssertEqual(37, bundle.Modules[RegulationModuleKind.CostRule].Records.Count);
            AssertEqual(33, bundle.Modules[RegulationModuleKind.MachineRate].Records.Count);
            AssertEqual(69, bundle.Modules[RegulationModuleKind.Geography].Records.Count);
            AssertEqual(7, bundle.Modules[RegulationModuleKind.Compliance].Records.Count);

            var maximumPages = new Dictionary<string, int>(StringComparer.Ordinal)
            {
                { "TT101-2025-BQP", 69 },
                { "VBHN94-2025-BQP", 30 },
                { "VBHN95-2025-BQP", 99 },
                { "VBHN96-2025-BQP", 35 },
                { "VBHN97-2025-BQP", 52 },
                { "VBHN98-2025-BQP", 76 }
            };
            foreach (RegulationDataModule module in bundle.Modules.Values)
            {
                foreach (RegulationDataRecord record in module.Records)
                {
                    AssertEqual(
                        RegulationDataVerification.VerifiedAgainstOfficialSource,
                        record.Verification);
                    AssertTrue(maximumPages.ContainsKey(record.Source.DocumentId));
                    AssertTrue(record.Source.PageTo <= maximumPages[record.Source.DocumentId]);
                    AssertFalse(string.IsNullOrWhiteSpace(record.Data));
                    ParseRecordData(record.Data);
                }
            }

            RegulationDataModule machines = bundle.Modules[RegulationModuleKind.MachineRate];
            AssertEqual(
                "613644600",
                ParseRecordData(FindRecord(machines, "MACHINE-M010.004").Data)["referencePriceVnd"]);
            AssertEqual(
                "4115480000",
                ParseRecordData(FindRecord(machines, "MACHINE-M010.016").Data)["referencePriceVnd"]);
            AssertEqual(
                "none",
                ParseRecordData(FindRecord(machines, "MACHINE-M010.020").Data)["operators"]);
            AssertEqual(
                "60-lit-gasoline-E5-RON-92-II",
                ParseRecordData(FindRecord(machines, "MACHINE-M010.027").Data)["fuel"]);
            AssertEqual(
                "6-si-quan+20-thuy-thu",
                ParseRecordData(FindRecord(machines, "MACHINE-M010.011").Data)["operators"]);
            AssertEqual(
                "1-bac-5-10",
                ParseRecordData(FindRecord(machines, "MACHINE-M010.020").Data)["nonStateOperators"]);
            AssertEqual(
                "350000",
                ParseRecordData(FindRecord(machines, "MACHINE-M010.024").Data)["nonStateReferencePriceVnd"]);

            RegulationDataModule norms = bundle.Modules[RegulationModuleKind.Norm];
            Dictionary<string, string> norm0200500 = ParseRecordData(
                FindRecord(norms, "NORM-020.0500").Data);
            AssertEqual("depth-0.5-or-1,depth-3,depth-5,depth-10", norm0200500["variantCodes"]);
            AssertTrue(norm0200500["rates"].Contains("6.40,7.05,7.76,8.54"));
            AssertTrue(norm0200500["rates"].Contains("4.27,4.70,5.17,0"));
            AssertTrue(norm0200500["rates"].Contains("0,0,0,5.69"));
            AssertTrue(norm0200500["rates"].Contains("MAT-RED-FLAG-LARGE:each:1,1,1,1"));
            AssertEqual(
                "132.1",
                ParseRecordData(FindRecord(norms, "NORM-PROVISIONAL-NONSTATE-ZONE-4").Data)
                    ["underwaterTo12mRate"]);

            RegulationDataModule costs = bundle.Modules[RegulationModuleKind.CostRule];
            AssertEqual(
                "2.2,2.0,1.9,1.8,1.7",
                ParseRecordData(FindRecord(costs, "COST-K2-LINEAR").Data)["ratesPercent"]);
            AssertEqual(
                "6300000",
                ParseRecordData(FindRecord(costs, "COST-MIN-SURVEY-DESIGN-UNDER-4HA").Data)
                    ["minimumVnd"]);
            Dictionary<string, string> currentK5 = ParseRecordData(
                FindRecord(costs, "COST-K5-CIVIL").Data);
            AssertEqual(
                "10,20,50,100,200,500,1000,2000,5000,8000,10000",
                currentK5["thresholdBillion"]);
            AssertEqual("linear", currentK5["interpolation"]);
            AssertEqual("TT38-2026-BXD-PL8-table-2.24", currentK5["currentExternalBasis"]);
            FindRecord(costs, "COST-TEMPLATE-05");

            RegulationDataModule geography = bundle.Modules[RegulationModuleKind.Geography];
            RegulationDataRecord[] localities = geography.Records
                .Where(record => record.Key.StartsWith("GEO-LOCALITY-", StringComparison.Ordinal))
                .ToArray();
            AssertEqual(34, localities.Length);
            var provinces = new HashSet<string>(StringComparer.Ordinal);
            foreach (RegulationDataRecord locality in localities)
            {
                Dictionary<string, string> data = ParseRecordData(locality.Data);
                AssertTrue(provinces.Add(data["province"]));
                string[] zones = data["zones"].Split(',');
                AssertEqual(zones.Length, zones.Distinct(StringComparer.Ordinal).Count());
                AssertTrue(zones.Contains(data["fallbackZone"], StringComparer.Ordinal));
                foreach (string zone in zones)
                    AssertTrue(zone == "1" || zone == "2" || zone == "3" || zone == "4");
            }
            AssertEqual(
                "Quang-Ninh,Hai-Phong,Ha-Tinh,Quang-Tri",
                ParseRecordData(FindRecord(geography, "GEO-SEA-ZONE-3").Data)["provinces"]);

            RegulationDataModule compliance = bundle.Modules[RegulationModuleKind.Compliance];
            AssertEqual(
                "2025-10-28",
                ParseRecordData(FindRecord(compliance, "COMPLIANCE-TT101-TRANSITION").Data)
                    ["cutoverDate"]);
            FindRecord(compliance, "COMPLIANCE-QCVN01-2022-RESOLVED");

            string temp = CreateTemporaryDirectory("bqp-2025-package");
            try
            {
                RegulationPackageBundle rebuilt = RegulationPackageBundleBuilder.Build(
                    sourcePath,
                    Path.Combine(temp, "bundle"));
                AssertEqual(bundle.Package.PackageChecksum, rebuilt.Package.PackageChecksum);
            }
            finally
            {
                DeleteTemporaryDirectory(temp);
            }
        }

        private static void TestNormCatalog()
        {
            NormCatalog catalog = NormCatalog.Load(LoadBqp2025NormModule());
            AssertEqual(34, catalog.Definitions.Count);

            NormDefinition investigation = catalog.FindRequired("NORM-000.0100");
            AssertEqual("commune", investigation.WorkUnit);
            AssertEqual(2, investigation.Variants.Count);
            AssertEqual("plain-midland", investigation.Variants[0]);
            AssertEqual(3, investigation.Rates.Count);
            AssertEqual(18, investigation.Source.PageFrom);

            NormDefinition deepLand = catalog.FindRequired("NORM-020.0900");
            AssertEqual(4, deepLand.Variants.Count);
            AssertEqual(10, deepLand.Rates.Count);
            AssertEqual(1, deepLand.Adjustments.Count);

            NormDefinition underwater = catalog.FindRequired("NORM-030.0100");
            AssertEqual(3, underwater.Adjustments.Count);
            AssertEqual(1, underwater.Constraints.Count);
            AssertEqual("prohibit-condition", underwater.Constraints[0].Operation);
            AssertThrows<KeyNotFoundException>(() => catalog.FindRequired("NORM-MISSING"));
        }

        private static void TestNormCalculation()
        {
            NormCatalog catalog = NormCatalog.Load(LoadBqp2025NormModule());
            NormCalculationResult deepScan = NormCalculator.Calculate(
                catalog.FindRequired("NORM-020.0500"),
                "depth-5",
                2m);
            AssertEqual(15.52m, deepScan.FindRequired(
                NormResourceKind.Labor, "LAB-QNCN-7").Quantity);
            AssertEqual(10.34m, deepScan.FindRequired(
                NormResourceKind.Machine, "M010.002").Quantity);
            AssertEqual(0m, deepScan.FindRequired(
                NormResourceKind.Machine, "M010.003").Quantity);
            AssertEqual("none-invariant-decimal", deepScan.RoundingRule);
            AssertEqual(24, deepScan.Source.PageFrom);

            NormCalculationResult slope = NormCalculator.Calculate(
                catalog.FindRequired("NORM-010.0100"),
                "forest-2",
                1m,
                new[] { "slope-gt-25deg" });
            AssertEqual(137.5m, slope.FindRequired(
                NormResourceKind.Labor, "LAB-QNCN-8").Quantity);

            NormCalculationResult uxo = NormCalculator.Calculate(
                catalog.FindRequired("NORM-020.0300"),
                "soil-1",
                2m,
                new[] { "uxo-signal" });
            AssertEqual(0.176m, uxo.FindRequired(
                NormResourceKind.Labor, "LAB-QNCN-8").Quantity);

            NormCalculationResult water = NormCalculator.Calculate(
                catalog.FindRequired("NORM-020.0600"),
                "soil-4",
                3m,
                new[] { "water-excavation" });
            AssertEqual(0.036m, water.FindRequired(
                NormResourceKind.Machine, "M010.023").Quantity);
        }

        private static void TestNormSearch()
        {
            RegulationPackageBundle bundle2025 = LoadBqpPackageBundle("BQP-RPBM-2025");
            NormCatalog catalog2025 = NormCatalog.Load(
                bundle2025.Modules[RegulationModuleKind.Norm]);
            var index2025 = new NormSearchIndex(
                bundle2025.Package.PackageId,
                bundle2025.Package.DataVersion,
                bundle2025.Package.PackageChecksum,
                catalog2025);

            IReadOnlyList<NormSearchResult> exact = index2025.Search(
                new NormSearchQuery("NORM-020.0500"));
            AssertEqual(1, exact.Count);
            AssertEqual(NormSearchMatchKind.ExactKey, exact[0].MatchKind);

            IReadOnlyList<NormSearchResult> noAccent = index2025.Search(
                new NormSearchQuery("ra pha bom min may do bom"));
            AssertTrue(noAccent.Any(result => result.Key == "NORM-020.0500"));

            IReadOnlyList<NormSearchResult> fuzzy = index2025.Search(
                new NormSearchQuery("dao dat kiem traa tin hieu"));
            AssertTrue(fuzzy.Any(result => result.Key == "NORM-020.0300"));

            IReadOnlyList<NormSearchResult> filtered = index2025.Search(
                new NormSearchQuery(
                    string.Empty,
                    NormEnvironment.Land,
                    3m,
                    "depth-3",
                    NormResourceKind.Machine));
            AssertEqual(1, filtered.Count);
            AssertEqual("NORM-020.0500", filtered[0].Key);

            CultureInfo searchCulture = CultureInfo.CurrentCulture;
            try
            {
                foreach (string cultureName in new[] { "vi-VN", "de-DE" })
                {
                    CultureInfo.CurrentCulture = new CultureInfo(cultureName);
                    AssertEqual(1, index2025.Search(new NormSearchQuery(
                        "NORM-020.0500",
                        NormEnvironment.Land,
                        3m)).Count);
                }
            }
            finally
            {
                CultureInfo.CurrentCulture = searchCulture;
            }

            AssertEqual(0, index2025.Search(new NormSearchQuery("NORM-999.9999")).Count);
            AssertThrows<ArgumentOutOfRangeException>(() => index2025.Search(
                new NormSearchQuery(string.Empty, maximumResults: 0)));

            RegulationPackageBundle bundle2021 = LoadBqpPackageBundle("BQP-RPBM-2021");
            var index2021 = new NormSearchIndex(
                bundle2021.Package.PackageId,
                bundle2021.Package.DataVersion,
                bundle2021.Package.PackageChecksum,
                bundle2021.Modules[RegulationModuleKind.Norm]);
            IReadOnlyList<NormSearchResult> sameCode = NormSearchIndex.SearchMany(
                new[] { index2021, index2025 },
                new NormSearchQuery("NORM-020.0500"));
            AssertEqual(2, sameCode.Count);
            AssertTrue(sameCode.Select(result => result.PackageId).Distinct().Count() == 2);
            AssertTrue(sameCode.Any(result => !result.HasDetailedRates));
        }

        private static void TestNormConditionsAndValidation()
        {
            CultureInfo oldCulture = CultureInfo.CurrentCulture;
            try
            {
                foreach (string cultureName in new[] { "vi-VN", "de-DE" })
                {
                    CultureInfo.CurrentCulture = new CultureInfo(cultureName);
                    NormCatalog catalog = NormCatalog.Load(LoadBqp2025NormModule());
                    NormCalculationResult current = NormCalculator.Calculate(
                        catalog.FindRequired("NORM-030.0100"),
                        "water-0.5-12",
                        1m,
                        new[] { "current-gt-0.5-le-1" });
                    AssertEqual(29.7750m, current.FindRequired(
                        NormResourceKind.Labor, "LAB-QNCN-7").Quantity);
                    AssertEqual(11.5750m, current.FindRequired(
                        NormResourceKind.Machine, "M010.008").Quantity);
                    AssertEqual(0.032m, current.FindRequired(
                        NormResourceKind.Material, "MAT-ANCHOR-50KG").Quantity);
                }
            }
            finally
            {
                CultureInfo.CurrentCulture = oldCulture;
            }

            NormCatalog valid = NormCatalog.Load(LoadBqp2025NormModule());
            AssertThrows<InvalidOperationException>(() => NormCalculator.Calculate(
                valid.FindRequired("NORM-020.0500"),
                "depth-0.5-or-1",
                1m));
            NormCalculator.Calculate(
                valid.FindRequired("NORM-020.0500"),
                "depth-0.5-or-1",
                1m,
                new[] { "forestry-salt-independent-or-owner-request" });
            AssertThrows<InvalidOperationException>(() => NormCalculator.Calculate(
                valid.FindRequired("NORM-030.0100"),
                "water-0.5-12",
                1m,
                new[] { "current-gt-2" }));
            AssertThrows<ArgumentException>(() => NormCalculator.Calculate(
                valid.FindRequired("NORM-030.0100"),
                "water-0.5-12",
                1m,
                new[] { "current-gt-0-le-0.5", "current-gt-0.5-le-1" }));

            RegulationDataModule module = LoadBqp2025NormModule();
            var records = module.Records.ToList();
            RegulationDataRecord first = records.First(record => record.RecordType == "NormCatalog");
            int firstIndex = records.IndexOf(first);
            records[firstIndex] = new RegulationDataRecord(
                first.Key,
                first.RecordType,
                first.Unit,
                first.Title,
                first.Data.Replace("15,22.5", "15,22,5"),
                first.Source,
                first.Verification);
            RegulationDataModule invalid = RegulationDataModule.Create(
                RegulationModuleKind.Norm,
                module.DataVersion,
                records);
            AssertThrows<FormatException>(() => NormCatalog.Load(invalid));
        }

        private static RegulationDataModule LoadBqp2025NormModule()
        {
            return LoadBqpPackageBundle("BQP-RPBM-2025")
                .Modules[RegulationModuleKind.Norm];
        }

        private static RegulationPackageBundle LoadBqpPackageBundle(string packageId)
        {
            string root = GetRepositoryRoot();
            string bundlePath = Path.Combine(
                root,
                "data",
                "regulations",
                "packages",
                packageId,
                "bundle");
            return RegulationPackageBundleReader.Read(bundlePath);
        }

        private static void TestCostRuleCatalog()
        {
            CostRuleCatalog catalog = CostRuleCatalog.Load(LoadBqp2025CostRuleModule());
            AssertEqual(40m, catalog.CommonRatePercent);
            AssertEqual(8, catalog.TerrainDefinitions.Count);
            AssertEqual(2m, catalog.FindTerrain("plain-open").K1Percent);
            AssertEqual(1m, catalog.FindTerrain("plain-open").K4Percent);
            AssertEqual(4, catalog.GetK2(CostProjectKind.Linear).ThresholdsBillion.Count);
            AssertEqual(5, catalog.GetK2(CostProjectKind.Linear).RatesPercent.Count);
            AssertEqual(3, catalog.K3Brackets.Count);
            AssertEqual(2000000L, catalog.K3MinimumVnd);
            AssertEqual(60000000L, catalog.K3MaximumVnd);
            AssertEqual(11, catalog.GetK5(CostConstructionKind.Civil).ThresholdsBillion.Count);
            AssertEqual(0.478m, catalog.GetK5(CostConstructionKind.Civil).RatesPercent[10]);
            AssertEqual(
                "TT38-2026-BXD-PL8-table-2.24",
                catalog.GetK5(CostConstructionKind.Civil).CurrentExternalBasis);
            AssertEqual("nearest-vnd-away-from-zero", catalog.MoneyRoundingRule);
            AssertThrows<KeyNotFoundException>(() => catalog.FindTerrain("missing"));
        }

        private static void TestCostRuleLegacyCatalog()
        {
            CostRuleCatalog catalog = CostRuleCatalog.Load(LoadBqp2021CostRuleModule());
            AssertEqual(40m, catalog.CommonRatePercent);
            AssertEqual(8, catalog.GetK5(CostConstructionKind.Civil).ThresholdsBillion.Count);
            AssertEqual("linear", catalog.GetK5(CostConstructionKind.Civil).Interpolation);
            AssertEqual(string.Empty, catalog.GetK5(CostConstructionKind.Civil).CurrentExternalBasis);
            AssertEqual(0m, catalog.MinimumAreaHa);
            AssertEqual(0L, catalog.MinimumSurveyVnd);
            AssertEqual(0L, catalog.MinimumDisposalVnd);
            AssertEqual(2.3m, CostRuleCalculator.CalculateK2(
                catalog, CostProjectKind.Linear, 15000000000m).RatePercent);
            AssertEqual(2000000L, CostRuleCalculator.CalculateK1K4(
                catalog, "plain-open", 100000000m, 4m).K1Vnd);
            AssertEqual(5000000L, CostRuleCalculator.CalculateK6(
                catalog, 100000000m, 999m, 4m).AmountVnd);
            AssertThrows<InvalidOperationException>(() => CostRuleCalculator.CalculateK5(
                catalog, CostConstructionKind.Civil, 2000000000001m));
        }

        private static void TestCostRuleCalculation()
        {
            CostRuleCatalog catalog = CostRuleCatalog.Load(LoadBqp2025CostRuleModule());
            DirectCostResult direct = CostRuleCalculator.CalculateDirect(catalog, 100m, 200m, 300m);
            AssertEqual(600L, direct.DirectVnd);
            AssertEqual(80L, direct.CommonVnd);
            AssertEqual(680L, direct.ZVnd);

            TerrainCostResult terrain = CostRuleCalculator.CalculateK1K4(
                catalog, "plain-open", 1000000000m, 10m);
            AssertEqual(20000000L, terrain.K1Vnd);
            AssertEqual(10000000L, terrain.K4Vnd);
            AssertEqual(30000000L, terrain.TotalVnd);

            TerrainCostResult minimumTerrain = CostRuleCalculator.CalculateK1K4(
                catalog, "plain-open", 100000000m, 4m);
            AssertEqual(6300000L, minimumTerrain.K1Vnd);
            AssertEqual(1000000L, minimumTerrain.K4Vnd);

            CostRateAmountResult k2Boundary = CostRuleCalculator.CalculateK2(
                catalog, CostProjectKind.Linear, 15000000000m);
            AssertEqual(2.2m, k2Boundary.RatePercent);
            CostRateAmountResult k2Above = CostRuleCalculator.CalculateK2(
                catalog, CostProjectKind.Linear, 15000000001m);
            AssertEqual(2.0m, k2Above.RatePercent);

            AssertEqual(2000000L, CostRuleCalculator.CalculateK3(catalog, 100000000m).AmountVnd);
            AssertEqual(6000000L, CostRuleCalculator.CalculateK3(catalog, 2000000000m).AmountVnd);
            AssertEqual(60000000L, CostRuleCalculator.CalculateK3(catalog, 50000000000m).AmountVnd);

            CostRateAmountResult k5 = CostRuleCalculator.CalculateK5(
                catalog, CostConstructionKind.Civil, 15000000000m);
            AssertEqual(3.069m, k5.RatePercent);
            AssertEqual(460350000L, k5.AmountVnd);
            AssertEqual(0.478m, CostRuleCalculator.CalculateK5(
                catalog, CostConstructionKind.Civil, 10000000000000m).RatePercent);
            AssertThrows<InvalidOperationException>(() => CostRuleCalculator.CalculateK5(
                catalog, CostConstructionKind.Civil, 10000000000001m));

            AssertEqual(5000000L, CostRuleCalculator.CalculateK6(
                catalog, 100000000m, 999m, 10m).AmountVnd);
            AssertEqual(6300000L, CostRuleCalculator.CalculateK6(
                catalog, 100000000m, 999m, 4m).AmountVnd);
            AssertEqual(3000000L, CostRuleCalculator.CalculateK6(
                catalog, 100000000m, 1001m, 10m).AmountVnd);
            AssertThrows<InvalidOperationException>(() => CostRuleCalculator.CalculateK6(
                catalog, 100000000m, 1000m, 10m));

            EstimateSummaryResult summary = CostRuleCalculator.CalculateSummary(
                100000000m,
                new[]
                {
                    new KeyValuePair<string, long>("K1", 3000000L),
                    new KeyValuePair<string, long>("K3", 2000000L),
                    new KeyValuePair<string, long>("K4", 1000000L),
                    new KeyValuePair<string, long>("K5", 3000000L)
                },
                10m);
            AssertEqual(9000000L, summary.OtherCostTotalVnd);
            AssertEqual(109000000L, summary.BeforeTaxVnd);
            AssertEqual(106000000L, summary.TaxableBaseVnd);
            AssertEqual(10600000L, summary.VatVnd);
            AssertEqual(119600000L, summary.AfterTaxVnd);
        }

        private static void TestCostRuleEngine()
        {
            CostRuleCatalog catalog = CostRuleCatalog.Load(LoadBqp2025CostRuleModule());
            var request = new CostRuleCalculationRequest(
                300000000m,
                200000000m,
                100000000m,
                "plain-open",
                10m,
                CostProjectKind.Linear,
                CostConstructionKind.Civil,
                500m,
                10m,
                CostComponentSelection.All);
            CostRuleEngineResult result = CostRuleEngine.Calculate(catalog, request);

            AssertEqual(6, result.Components.Count);
            TerrainCostResult terrain = CostRuleCalculator.CalculateK1K4(
                catalog, request.Terrain, result.DirectCost.ZVnd, request.AreaHa);
            AssertEqual(terrain.K1Vnd, result.FindRequired("K1").CalculatedAmountVnd);
            AssertEqual(terrain.K4Vnd, result.FindRequired("K4").CalculatedAmountVnd);
            AssertEqual(CostRuleCalculator.CalculateK2(
                catalog, request.ProjectKind, result.DirectCost.DirectVnd).AmountVnd,
                result.FindRequired("K2").CalculatedAmountVnd);
            AssertEqual(CostRuleCalculator.CalculateK3(
                catalog, result.DirectCost.ZVnd).AmountVnd,
                result.FindRequired("K3").CalculatedAmountVnd);
            AssertEqual(CostRuleCalculator.CalculateK5(
                catalog, request.ConstructionKind, result.DirectCost.ZVnd).AmountVnd,
                result.FindRequired("K5").CalculatedAmountVnd);
            AssertEqual(CostRuleCalculator.CalculateK6(
                catalog, result.DirectCost.ZVnd, request.DisposalWeightKg, request.AreaHa).AmountVnd,
                result.FindRequired("K6").CalculatedAmountVnd);
            AssertTrue(result.FindRequired("K2").ExternalBasis.Length > 0);
            AssertTrue(result.FindRequired("K5").ExternalBasis.Length > 0);

            const long overriddenK5 = 12345678L;
            CostRuleEngineResult overridden = CostRuleEngine.Calculate(
                catalog,
                new CostRuleCalculationRequest(
                    request.MaterialVnd,
                    request.LaborVnd,
                    request.MachineVnd,
                    request.Terrain,
                    request.AreaHa,
                    request.ProjectKind,
                    request.ConstructionKind,
                    request.DisposalWeightKg,
                    request.VatRatePercent,
                    request.SelectedComponents,
                    new[] { new CostRuleOverride("k5", overriddenK5, "Du toan rieng duoc phe duyet") }));
            CostComponentCalculationResult originalK5 = result.FindRequired("K5");
            CostComponentCalculationResult appliedK5 = overridden.FindRequired("K5");
            AssertEqual(originalK5.CalculatedAmountVnd, appliedK5.CalculatedAmountVnd);
            AssertEqual(overriddenK5, appliedK5.AppliedAmountVnd);
            AssertTrue(appliedK5.IsOverridden);
            AssertEqual(
                result.Summary.BeforeTaxVnd - originalK5.AppliedAmountVnd + overriddenK5,
                overridden.Summary.BeforeTaxVnd);

            AssertThrows<ArgumentOutOfRangeException>(() => CostRuleEngine.Calculate(
                catalog,
                new CostRuleCalculationRequest(0m, 0m, 0m, "plain-open", 1m,
                    CostProjectKind.Linear, CostConstructionKind.Civil, 1m, 10m,
                    CostComponentSelection.None)));
            AssertThrows<ArgumentException>(() => CostRuleEngine.Calculate(
                catalog,
                new CostRuleCalculationRequest(0m, 1m, 0m, "plain-open", 1m,
                    CostProjectKind.Linear, CostConstructionKind.Civil, 1m, 10m,
                    CostComponentSelection.K1,
                    new[] { new CostRuleOverride("K1", 1L, string.Empty) })));
            AssertThrows<ArgumentException>(() => CostRuleEngine.Calculate(
                catalog,
                new CostRuleCalculationRequest(0m, 1m, 0m, "plain-open", 1m,
                    CostProjectKind.Linear, CostConstructionKind.Civil, 1m, 10m,
                    CostComponentSelection.K1,
                    new[] { new CostRuleOverride("K2", 1L, "Khong duoc chon") })));
            AssertThrows<ArgumentException>(() => CostRuleEngine.Calculate(
                catalog,
                new CostRuleCalculationRequest(0m, 1m, 0m, "plain-open", 1m,
                    CostProjectKind.Linear, CostConstructionKind.Civil, 1m, 10m,
                    CostComponentSelection.K1,
                    new[]
                    {
                        new CostRuleOverride("K1", 1L, "Lan 1"),
                        new CostRuleOverride("K1", 2L, "Lan 2")
                    })));
            AssertThrows<InvalidOperationException>(() => CostRuleEngine.Calculate(
                catalog,
                new CostRuleCalculationRequest(0m, 1m, 0m, "plain-open", 10m,
                    CostProjectKind.Linear, CostConstructionKind.Civil, 1000m, 10m,
                    CostComponentSelection.K6)));
        }

        private static void TestCostRuleLocaleAndValidation()
        {
            CultureInfo oldCulture = CultureInfo.CurrentCulture;
            try
            {
                foreach (string cultureName in new[] { "vi-VN", "de-DE" })
                {
                    CultureInfo.CurrentCulture = new CultureInfo(cultureName);
                    CostRuleCatalog catalog = CostRuleCatalog.Load(LoadBqp2025CostRuleModule());
                    AssertEqual(460350000L, CostRuleCalculator.CalculateK5(
                        catalog, CostConstructionKind.Civil, 15000000000m).AmountVnd);
                }
            }
            finally
            {
                CultureInfo.CurrentCulture = oldCulture;
            }

            RegulationDataModule module = LoadBqp2025CostRuleModule();
            var records = module.Records.ToList();
            RegulationDataRecord table = records.Single(record => record.Key == "COST-K5-CIVIL");
            int index = records.IndexOf(table);
            records[index] = new RegulationDataRecord(
                table.Key,
                table.RecordType,
                table.Unit,
                table.Title,
                table.Data.Replace("10,20,50", "10,50,20"),
                table.Source,
                table.Verification);
            RegulationDataModule invalid = RegulationDataModule.Create(
                RegulationModuleKind.CostRule,
                module.DataVersion,
                records);
            AssertThrows<FormatException>(() => CostRuleCatalog.Load(invalid));
            AssertThrows<ArgumentOutOfRangeException>(() => CostRuleCalculator.CalculateDirect(
                CostRuleCatalog.Load(module), -1m, 0m, 0m));
        }

        private static void TestPriceProfileRoundTrip()
        {
            PriceProfile expected = CreateSamplePriceProfile();
            string first = PriceProfileSerializer.Serialize(expected);
            string second = PriceProfileSerializer.Serialize(expected);
            AssertEqual(first, second);
            AssertEqual(64, expected.Checksum.Length);

            PriceProfile actual = PriceProfileSerializer.Deserialize(first);
            AssertEqual(expected.ProfileId, actual.ProfileId);
            AssertEqual(expected.DataVersion, actual.DataVersion);
            AssertEqual(expected.Location, actual.Location);
            AssertEqual(expected.ValuationDate, actual.ValuationDate);
            AssertEqual(expected.LaborAudience, actual.LaborAudience);
            AssertEqual(expected.Entries.Count, actual.Entries.Count);
            AssertEqual(expected.Overrides.Count, actual.Overrides.Count);
            AssertEqual(expected.Checksum, actual.Checksum);
            AssertEqual(517500m, actual.FindRequired("bac-8-10").AppliedUnitPriceVnd);
            AssertEqual(18029.126213592233009708737864m,
                actual.FindRequired("diesel").AppliedUnitPriceVnd);
            AssertTrue(actual.FindRequired("diesel").IsOverridden);

            string tampered = first.Replace("basePrice=18000", "basePrice=18001");
            AssertFalse(string.Equals(first, tampered, StringComparison.Ordinal));
            AssertThrows<InvalidDataException>(() => PriceProfileSerializer.Deserialize(tampered));

            CultureInfo oldCulture = CultureInfo.CurrentCulture;
            try
            {
                foreach (string cultureName in new[] { "vi-VN", "de-DE" })
                {
                    CultureInfo.CurrentCulture = new CultureInfo(cultureName);
                    AssertEqual(first, PriceProfileSerializer.Serialize(expected));
                    AssertEqual(expected.Checksum, PriceProfileSerializer.Deserialize(first).Checksum);
                }
            }
            finally
            {
                CultureInfo.CurrentCulture = oldCulture;
            }
        }

        private static void TestPriceProfileValidationAndCoverage()
        {
            PriceProfile profile = CreateSamplePriceProfile();
            PriceCoverageResult complete = PriceProfileCoverageValidator.Validate(
                profile,
                new[]
                {
                    new PriceRequirement("MAT-CONCRETE-STAKE", PriceResourceKind.Material, "each"),
                    new PriceRequirement("bac-8-10", PriceResourceKind.Labor, "worker-day")
                });
            AssertTrue(complete.IsComplete);

            PriceCoverageResult incomplete = PriceProfileCoverageValidator.Validate(
                profile,
                new[]
                {
                    new PriceRequirement("MAT-MISSING", PriceResourceKind.Material, "each"),
                    new PriceRequirement("MAT-CONCRETE-STAKE", PriceResourceKind.Material, "kg"),
                    new PriceRequirement("diesel", PriceResourceKind.Material, "lit"),
                    new PriceRequirement("diesel", PriceResourceKind.FuelEnergy, "lit")
                });
            AssertEqual(4, incomplete.Issues.Count);
            AssertTrue(incomplete.Issues.Any(item => item.IssueCode == PriceCoverageIssueCode.Missing));
            AssertTrue(incomplete.Issues.Any(item => item.IssueCode == PriceCoverageIssueCode.UnitMismatch));
            AssertTrue(incomplete.Issues.Any(item => item.IssueCode == PriceCoverageIssueCode.KindMismatch));
            AssertTrue(incomplete.Issues.Any(item => item.IssueCode == PriceCoverageIssueCode.InvalidRequirement));

            DateTime created = new DateTime(2026, 8, 3, 1, 2, 3, DateTimeKind.Utc);
            AssertThrows<ArgumentException>(() => PriceProfile.Create(
                "PRICE-DUPLICATE",
                "1.0.0",
                "Trung alias",
                "Ha Noi",
                new DateTime(2026, 7, 15),
                MachineRateAudience.NonStateSalary,
                created,
                new[]
                {
                    new PriceProfileEntry("A", PriceResourceKind.Material, "A", "each", 1m,
                        new DateTime(2026, 7, 1), PriceSourceKind.Manual, "Nguon A", new[] { "DUP" }),
                    new PriceProfileEntry("B", PriceResourceKind.Material, "B", "each", 1m,
                        new DateTime(2026, 7, 1), PriceSourceKind.Manual, "Nguon B", new[] { "dup" })
                }));
            AssertThrows<ArgumentException>(() => PriceProfile.Create(
                "PRICE-NEGATIVE",
                "1.0.0",
                "Gia am",
                "Ha Noi",
                new DateTime(2026, 7, 15),
                MachineRateAudience.NonStateSalary,
                created,
                new[]
                {
                    new PriceProfileEntry("A", PriceResourceKind.Material, "A", "each", -1m,
                        new DateTime(2026, 7, 1), PriceSourceKind.Manual, "Nguon A")
                }));
            AssertThrows<ArgumentException>(() => PriceProfile.Create(
                "PRICE-BLANK-REASON",
                "1.0.0",
                "Ly do trong",
                "Ha Noi",
                new DateTime(2026, 7, 15),
                MachineRateAudience.NonStateSalary,
                created,
                new[]
                {
                    new PriceProfileEntry("A", PriceResourceKind.Material, "A", "each", 1m,
                        new DateTime(2026, 7, 1), PriceSourceKind.Manual, "Nguon A")
                },
                new[]
                {
                    new PriceProfileOverride("A", 1m, 2m, string.Empty, "QD-1", "Tester", created)
                }));
        }

        private static void TestPriceProfileMachineAdapter()
        {
            PriceProfile profile = CreateSamplePriceProfile();
            MachineRateCatalog catalog = MachineRateCatalog.Load(
                LoadBqp2025MachineModule(),
                MachineRateAudience.NonStateSalary);
            MachineRateResult result = MachineRateCalculator.Calculate(
                catalog.FindRequiredByCode("M011.004"),
                profile.ToMachineRatePriceProfile());
            AssertEqual(315589L, result.DepreciationVnd);
            AssertEqual(105196L, result.RepairVnd);
            AssertEqual(538530L, result.FuelVnd);
            AssertEqual(517500L, result.OperatorLaborVnd);
            AssertEqual(109579L, result.OtherVnd);
            AssertEqual(1586394L, result.TotalVnd);
        }

        private static void TestPriceProfileStore()
        {
            string root = CreateTemporaryDirectory("price-profile-store");
            try
            {
                var store = new PriceProfileStore(root);
                PriceProfile profile = CreateSamplePriceProfile();
                PriceProfileInstallResult installed = store.Import(profile);
                AssertEqual(PriceProfileInstallStatus.Installed, installed.Status);
                AssertEqual(PriceProfileInstallStatus.AlreadyInstalled, store.Import(profile).Status);
                AssertEqual(1, store.ListInstalled().Count);
                AssertEqual(profile.Checksum, store.LoadRequired(
                    profile.ProfileId,
                    profile.DataVersion,
                    profile.Checksum).Checksum);
                AssertThrows<InvalidDataException>(() => store.LoadRequired(
                    profile.ProfileId,
                    profile.DataVersion,
                    new string('A', 64)));
                AssertThrows<ArgumentException>(() => store.LoadRequired(
                    "../outside",
                    profile.DataVersion,
                    profile.Checksum));

                DateTime created = new DateTime(2026, 8, 3, 1, 2, 3, DateTimeKind.Utc);
                PriceProfile conflict = PriceProfile.Create(
                    profile.ProfileId,
                    profile.DataVersion,
                    profile.DisplayName,
                    profile.Location,
                    profile.ValuationDate,
                    profile.LaborAudience,
                    created,
                    new[]
                    {
                        new PriceProfileEntry(
                            "MAT-CONCRETE-STAKE",
                            PriceResourceKind.Material,
                            "Coc be tong cot thep",
                            "each",
                            110000m,
                            new DateTime(2026, 7, 1),
                            PriceSourceKind.MarketQuote,
                            "Bao gia VL-02")
                    });
                AssertThrows<InvalidOperationException>(() => store.Import(conflict));

                string installedFile = Path.Combine(installed.Directory, "profile.ttbprice");
                string payload = File.ReadAllText(installedFile, Encoding.UTF8);
                File.WriteAllText(
                    installedFile,
                    payload.Replace("basePrice=18000", "basePrice=18001"),
                    new UTF8Encoding(false));
                AssertThrows<InvalidDataException>(() => store.LoadRequired(
                    profile.ProfileId,
                    profile.DataVersion,
                    profile.Checksum));
            }
            finally
            {
                DeleteTemporaryDirectory(root);
            }
        }

        private static PriceProfile CreateSamplePriceProfile()
        {
            DateTime created = new DateTime(2026, 8, 3, 1, 2, 3, DateTimeKind.Utc);
            return PriceProfile.Create(
                "PRICE-HA-NOI-2026-07",
                "1.0.0",
                "Gia Ha Noi thang 07/2026",
                "Ha Noi",
                new DateTime(2026, 7, 15),
                MachineRateAudience.NonStateSalary,
                created,
                new[]
                {
                    new PriceProfileEntry(
                        "MAT-CONCRETE-STAKE",
                        PriceResourceKind.Material,
                        "Coc be tong cot thep",
                        "each",
                        100000m,
                        new DateTime(2026, 7, 1),
                        PriceSourceKind.MarketQuote,
                        "Bao gia VL-01"),
                    new PriceProfileEntry(
                        "LAB-QNCN-8",
                        PriceResourceKind.Labor,
                        "Nhan cong bac 8/10",
                        "worker-day",
                        517500m,
                        new DateTime(2026, 7, 1),
                        PriceSourceKind.PublishedNotice,
                        "Bang gia NC-01",
                        new[] { "bac-8-10" }),
                    new PriceProfileEntry(
                        "diesel",
                        PriceResourceKind.FuelEnergy,
                        "Dau diesel",
                        "lit",
                        18000m,
                        new DateTime(2026, 7, 1),
                        PriceSourceKind.PublishedNotice,
                        "Thong bao gia NL-01")
                },
                new[]
                {
                    new PriceProfileOverride(
                        "diesel",
                        18000m,
                        18029.126213592233009708737864m,
                        "Gia giao den hien truong",
                        "Bao gia NL-02",
                        "Tester",
                        created)
                });
        }

        private static void TestPriceProfilePortfolioRoundTrip()
        {
            PriceProfile nonState = CreateSamplePriceProfile();
            PriceProfile state = PriceProfile.Create(
                nonState.ProfileId + "-HLNS",
                nonState.DataVersion,
                nonState.DisplayName + " HLNS",
                nonState.Location,
                nonState.ValuationDate,
                MachineRateAudience.StateBudgetSalary,
                DateTime.UtcNow,
                nonState.Entries,
                nonState.Overrides);
            PriceProfilePortfolio portfolio = PriceProfilePortfolio.Create(new[] { nonState, state });
            string payload = PriceProfilePortfolioSerializer.Serialize(portfolio);
            PriceProfilePortfolio restored = PriceProfilePortfolioSerializer.Deserialize(payload);

            AssertEqual(2, restored.Profiles.Count);
            AssertEqual(nonState.Checksum, restored.FindRequired(MachineRateAudience.NonStateSalary).Checksum);
            AssertEqual(state.Checksum, restored.FindRequired(MachineRateAudience.StateBudgetSalary).Checksum);
            AssertThrows<ArgumentException>(() => PriceProfilePortfolio.Create(new[] { nonState, nonState }));
        }

        private static void TestUnitRateCalculation()
        {
            NormCatalog catalog = NormCatalog.Load(LoadBqp2025NormModule());
            UnitRateCalculationResult result = UnitRateCalculator.Calculate(
                new UnitRateCalculationRequest(
                    catalog.FindRequired("NORM-020.0200"),
                    "density-1",
                    CreateUnitRatePriceProfile()));

            AssertEqual("NORM-020.0200", result.NormKey);
            AssertEqual("density-1", result.VariantCode);
            AssertEqual(7, result.Resources.Count);
            AssertEqual(851550m, result.MaterialAmountVnd);
            AssertEqual(8593200m, result.LaborAmountVnd);
            AssertEqual(8827863.72m, result.MachineAmountVnd);
            AssertEqual(18272613.72m, result.TotalAmountVnd);
            UnitRateResourceAmount other = result.Resources.Single(
                item => item.ResourceCode == "MAT-OTHER");
            AssertTrue(other.IsPercentage);
            AssertEqual(5m, other.Quantity);
            AssertEqual(40550m, other.AmountVnd);
            AssertTrue(other.PriceSourceReference.StartsWith("VBHN97-2025-BQP, page ", StringComparison.Ordinal));
            AssertEqual("none-invariant-decimal", result.RoundingRule);
        }

        private static void TestUnitRateConditionsAndBindings()
        {
            NormCatalog catalog = NormCatalog.Load(LoadBqp2025NormModule());
            NormDefinition definition = catalog.FindRequired("NORM-030.0400");
            PriceProfile prices = CreateUnitRatePriceProfile();
            var binding = new UnitRateResourceBinding(
                "M010.DIVING",
                "MACHINE-M010.029",
                "Thiet bi lan 0,5m-3m theo bien phap duoc duyet");

            UnitRateCalculationResult plain = UnitRateCalculator.Calculate(
                new UnitRateCalculationRequest(
                    definition,
                    "water-0.5-12",
                    prices,
                    null,
                    new[] { binding }));
            AssertEqual(113850m, plain.LaborAmountVnd);
            AssertEqual(441645.140m, plain.MachineAmountVnd);
            UnitRateResourceAmount diving = plain.Resources.Single(
                item => item.ResourceCode == "M010.DIVING");
            AssertEqual("MACHINE-M010.029", diving.PriceCode);
            AssertEqual(binding.Reason, diving.BindingReason);

            UnitRateCalculationResult factored = UnitRateCalculator.Calculate(
                new UnitRateCalculationRequest(
                    definition,
                    "water-0.5-12",
                    prices,
                    new[] { "current-gt-0-le-0.5" },
                    new[] { binding }));
            AssertEqual(125235.0m, factored.LaborAmountVnd);
            AssertEqual(485809.6540m, factored.MachineAmountVnd);

            CultureInfo oldCulture = CultureInfo.CurrentCulture;
            try
            {
                foreach (string cultureName in new[] { "vi-VN", "de-DE" })
                {
                    CultureInfo.CurrentCulture = new CultureInfo(cultureName);
                    UnitRateCalculationResult localized = UnitRateCalculator.Calculate(
                        new UnitRateCalculationRequest(
                            definition,
                            "water-0.5-12",
                            prices,
                            new[] { "current-gt-0-le-0.5" },
                            new[] { binding }));
                    AssertEqual(factored.TotalAmountVnd, localized.TotalAmountVnd);
                }
            }
            finally
            {
                CultureInfo.CurrentCulture = oldCulture;
            }
        }

        private static void TestUnitRateValidation()
        {
            NormCatalog catalog = NormCatalog.Load(LoadBqp2025NormModule());
            PriceProfile missingMachine = CreateUnitRatePriceProfile(false);
            var missingRequest = new UnitRateCalculationRequest(
                catalog.FindRequired("NORM-020.0200"),
                "density-1",
                missingMachine);
            UnitRateValidationResult missing = UnitRateCalculator.Validate(missingRequest);
            AssertFalse(missing.IsValid);
            AssertTrue(missing.Issues.Any(issue =>
                issue.IssueCode == UnitRateIssueCode.MissingPrice &&
                issue.ResourceCode == "M010.001"));
            AssertThrows<UnitRateValidationException>(() => UnitRateCalculator.Calculate(missingRequest));

            var logicalRequest = new UnitRateCalculationRequest(
                catalog.FindRequired("NORM-030.0400"),
                "water-0.5-12",
                CreateUnitRatePriceProfile());
            UnitRateValidationResult logical = UnitRateCalculator.Validate(logicalRequest);
            AssertTrue(logical.Issues.Any(issue =>
                issue.IssueCode == UnitRateIssueCode.BindingRequired &&
                issue.ResourceCode == "M010.DIVING"));

            var blankReason = new UnitRateCalculationRequest(
                catalog.FindRequired("NORM-030.0400"),
                "water-0.5-12",
                CreateUnitRatePriceProfile(),
                null,
                new[] { new UnitRateResourceBinding("M010.DIVING", "MACHINE-M010.029", "") });
            AssertTrue(UnitRateCalculator.Validate(blankReason).Issues.Any(
                issue => issue.IssueCode == UnitRateIssueCode.InvalidBinding));

            PriceProfile wrongUnit = CreateUnitRatePriceProfile(true, "hour");
            UnitRateValidationResult unitMismatch = UnitRateCalculator.Validate(
                new UnitRateCalculationRequest(
                    catalog.FindRequired("NORM-020.0200"),
                    "density-1",
                    wrongUnit));
            AssertTrue(unitMismatch.Issues.Any(issue =>
                issue.IssueCode == UnitRateIssueCode.UnitMismatch &&
                issue.ResourceCode == "M010.001"));

            UnitRateValidationResult zeroConsumption = UnitRateCalculator.Validate(
                new UnitRateCalculationRequest(
                    catalog.FindRequired("NORM-020.0500"),
                    "depth-3",
                    CreateUnitRatePriceProfile()));
            AssertTrue(zeroConsumption.IsValid);
            AssertFalse(zeroConsumption.Issues.Any(issue => issue.ResourceCode == "M010.003"));
        }

        private static void TestEstimateAppendixCalculation()
        {
            NormCatalog catalog = NormCatalog.Load(LoadBqp2025NormModule());
            PriceProfile prices = CreateUnitRatePriceProfile();
            UnitRateCalculationResult land = UnitRateCalculator.Calculate(
                new UnitRateCalculationRequest(
                    catalog.FindRequired("NORM-020.0200"),
                    "density-1",
                    prices));
            UnitRateCalculationResult water = UnitRateCalculator.Calculate(
                new UnitRateCalculationRequest(
                    catalog.FindRequired("NORM-030.0400"),
                    "water-0.5-12",
                    prices,
                    null,
                    new[]
                    {
                        new UnitRateResourceBinding(
                            "M010.DIVING",
                            "MACHINE-M010.029",
                            "Bien phap lan 0,5m-3m")
                    }));
            var request = new EstimateAppendixCalculationRequest(
                "BQP-RPBM-2025",
                "2.0.0",
                new string('A', 64),
                new[]
                {
                    new EstimateAppendixLineRequest(
                        "LAND-1", "LAND", "Do tim tren can", "ha",
                        3.81m, 1m, land, "THKL!F18", "Gia DT TC!F13:N13"),
                    new EstimateAppendixLineRequest(
                        "WATER-1", "WATER", "Lan xu ly", "signal",
                        133m, 2m, water, "Gia DT TC!D23", "Gia DT TC!F23:N23")
                });

            EstimateAppendixCalculationResult result = EstimateAppendixCalculator.Calculate(request);
            AssertEqual(2, result.Lines.Count);
            AssertEqual(2, result.Groups.Count);
            AssertEqual(3244405.5m, result.Lines[0].MaterialAmountVnd);
            AssertEqual(32740092m, result.Lines[0].LaborAmountVnd);
            AssertEqual(33634160.7732m, result.Lines[0].MachineAmountVnd);
            AssertEqual(15142050m, result.Lines[1].LaborAmountVnd);
            AssertEqual(58738803.620m, result.Lines[1].MachineAmountVnd);
            AssertEqual(851550m, result.Lines[0].AcceptedMaterialAmountVnd);
            AssertEqual(227700m, result.Lines[1].AcceptedLaborAmountVnd);
            AssertEqual(883290.280m, result.Lines[1].AcceptedMachineAmountVnd);
            AssertEqual(3244405.5m, result.MaterialAmountVnd);
            AssertEqual(47882142m, result.LaborAmountVnd);
            AssertEqual(92372964.3932m, result.MachineAmountVnd);
            AssertEqual(prices.Checksum, result.UnitRateProfileChecksum);
            AssertEqual("none-invariant-decimal", result.RoundingRule);
        }

        private static void TestEstimateAppendixValidation()
        {
            NormCatalog catalog = NormCatalog.Load(LoadBqp2025NormModule());
            PriceProfile firstProfile = CreateUnitRatePriceProfile();
            PriceProfile secondProfile = PriceProfile.Create(
                "PRICE-UNIT-RATE-OTHER",
                "1.0.0",
                "Bang gia khac",
                "Ha Noi",
                firstProfile.ValuationDate,
                firstProfile.LaborAudience,
                firstProfile.CreatedAtUtc,
                firstProfile.Entries);
            UnitRateCalculationResult firstRate = UnitRateCalculator.Calculate(
                new UnitRateCalculationRequest(
                    catalog.FindRequired("NORM-020.0200"),
                    "density-1",
                    firstProfile));
            UnitRateCalculationResult secondRate = UnitRateCalculator.Calculate(
                new UnitRateCalculationRequest(
                    catalog.FindRequired("NORM-020.0200"),
                    "density-1",
                    secondProfile));
            var invalid = new EstimateAppendixCalculationRequest(
                string.Empty,
                "2.0.0",
                new string('A', 64),
                new[]
                {
                    new EstimateAppendixLineRequest(
                        "DUP", "LAND", "Dong 1", "ha", -1m, 0m, firstRate, "D1", "F1"),
                    new EstimateAppendixLineRequest(
                        "DUP", "LAND", "Dong 2", "ha", 1m, 0m, secondRate, "D2", "F2"),
                    new EstimateAppendixLineRequest(
                        "NO-RATE", "WATER", "Dong 3", "signal", 1m, 0m, null, "D3", "F3"),
                    new EstimateAppendixLineRequest(
                        "ZERO", "WATER", "Dong 4", "signal", 0m, 0m, null, "D4", "F4")
                });
            EstimateAppendixValidationResult validation = EstimateAppendixCalculator.Validate(invalid);
            AssertFalse(validation.IsValid);
            AssertTrue(validation.Issues.Any(issue => issue.IssueCode == EstimateAppendixIssueCode.InvalidIdentity));
            AssertTrue(validation.Issues.Any(issue => issue.IssueCode == EstimateAppendixIssueCode.DuplicateLine));
            AssertTrue(validation.Issues.Any(issue => issue.IssueCode == EstimateAppendixIssueCode.InvalidQuantity));
            AssertTrue(validation.Issues.Any(issue => issue.IssueCode == EstimateAppendixIssueCode.MissingUnitRate));
            AssertTrue(validation.Issues.Any(issue => issue.IssueCode == EstimateAppendixIssueCode.MixedPriceProfile));
            AssertThrows<EstimateAppendixValidationException>(() =>
                EstimateAppendixCalculator.Calculate(invalid));

            var zeroOnly = new EstimateAppendixCalculationRequest(
                "BQP-RPBM-2025",
                "2.0.0",
                new string('B', 64),
                new[]
                {
                    new EstimateAppendixLineRequest(
                        "ZERO", "LAND", "Khong co khoi luong", "ha",
                        0m, 0m, null, "D1", "F1")
                });
            AssertTrue(EstimateAppendixCalculator.Validate(zeroOnly).IsValid);
            AssertEqual(0m, EstimateAppendixCalculator.Calculate(zeroOnly).TotalAmountVnd);
        }

        private static void TestEstimateWorkspaceRoundTrip()
        {
            EstimateWorkspace workspace = CreateEstimateWorkspace();
            string payload = EstimateWorkspaceSerializer.Serialize(workspace);
            EstimateWorkspace restored = EstimateWorkspaceSerializer.Deserialize(payload);

            AssertEqual("SheetEstimate", restored.Source.WorksheetCodeName);
            AssertEqual("$A$1:$K$7", restored.Source.SourceAddress);
            AssertEqual(6, restored.Rows.Count);
            AssertEqual("=B2*C2", restored.Rows[0].QuantityFormulaLocal);
            AssertEqual("NORM-020.0200", restored.Rows[0].NormKey);
            AssertEqual("density-1", restored.Rows[0].VariantCode);
            AssertTrue(restored.Rows[5].IsTextRow);
            AssertThrows<InvalidDataException>(() => EstimateWorkspaceSerializer.Deserialize(
                payload.Replace("description=", "description=X")));
        }

        private static void TestEstimateV2StateRoundTrip()
        {
            string id = "00112233445566778899aabbccddeeff";
            string fingerprint = EstimateV2Fingerprint.Compute(
                "TC.01",
                "Do tim tren can den do sau 0,3m",
                "ha",
                "WORKITEM");
            var item = new EstimateV2WorkItemState(
                id,
                "SheetEstimate",
                "NORM-020.0200",
                "density-1",
                "BQP-RPBM-2025",
                "2.0.1",
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
                "WORKITEM",
                fingerprint,
                false);
            var state = new EstimateV2State(
                new[] { item },
                new DateTime(2026, 10, 6, 2, 0, 0, DateTimeKind.Utc));

            string xml = EstimateV2StateSerializer.Serialize(state);
            EstimateV2State restored = EstimateV2StateSerializer.Deserialize(xml);

            AssertEqual(1, restored.WorkItems.Count);
            AssertEqual(id, restored.WorkItems[0].WorkItemId);
            AssertEqual("SheetEstimate", restored.WorkItems[0].SourceKey);
            AssertEqual("NORM-020.0200", restored.WorkItems[0].NormCode);
            AssertEqual("density-1", restored.WorkItems[0].VariantCode);
            AssertEqual("BQP-RPBM-2025", restored.WorkItems[0].PackageId);
            AssertEqual("2.0.1", restored.WorkItems[0].DataVersion);
            AssertEqual("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", restored.WorkItems[0].PackageChecksum);
            AssertEqual("WORKITEM", restored.WorkItems[0].Kind);
            AssertEqual(fingerprint, restored.WorkItems[0].Fingerprint);
            AssertFalse(restored.WorkItems[0].IsOrphaned);
            AssertTrue(restored.WorkItems[0].HasNormBinding);

            EstimateV2State updated = restored.Upsert(
                restored.WorkItems[0].WithoutBinding().WithOrphaned(true),
                DateTime.UtcNow);
            AssertFalse(updated.WorkItems[0].HasNormBinding);
            AssertTrue(updated.WorkItems[0].IsOrphaned);

            AssertThrows<ArgumentException>(() => new EstimateV2State(
                new[] { item, item },
                DateTime.UtcNow));
            AssertThrows<InvalidDataException>(() =>
                EstimateV2StateSerializer.Deserialize(
                    xml.Replace("schemaVersion=\"1\"", "schemaVersion=\"99\"")));
            AssertThrows<ArgumentException>(() => new EstimateV2WorkItemState(
                "not-a-guid",
                "SheetEstimate",
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                "WORKITEM",
                string.Empty,
                false));
        }

        private static void TestEstimateV2Fingerprint()
        {
            string first = EstimateV2Fingerprint.Compute(
                " tc.01 ",
                "Do   tim\ntren can",
                "ha",
                "workitem");
            string second = EstimateV2Fingerprint.Compute(
                "TC.01",
                "DO TIM TREN CAN",
                "HA",
                "WORKITEM");
            string third = EstimateV2Fingerprint.Compute(
                "TC.02",
                "DO TIM TREN CAN",
                "HA",
                "WORKITEM");

            AssertEqual(first, second);
            AssertFalse(string.Equals(first, third, StringComparison.Ordinal));
            AssertEqual(64, first.Length);
        }

        private static void TestEstimateRateGrouping()
        {
            EstimateWorkspace workspace = CreateEstimateWorkspace();
            PriceProfile nonState = CreateUnitRatePriceProfile();
            PriceProfile state = PriceProfile.Create(
                nonState.ProfileId + "-HLNS",
                nonState.DataVersion,
                nonState.DisplayName + " HLNS",
                nonState.Location,
                nonState.ValuationDate,
                MachineRateAudience.StateBudgetSalary,
                DateTime.UtcNow,
                nonState.Entries,
                nonState.Overrides);
            EstimateRatePlan plan = EstimateRatePlan.Build(
                workspace,
                "BQP-RPBM-2025",
                "2.0.1",
                new string('A', 64),
                new[] { nonState, state });

            AssertTrue(plan.IsValid);
            AssertEqual(4, plan.Groups.Count);
            AssertEqual(1, plan.TextRows.Count);
            EstimateRateGroup shared = plan.Groups.Single(group =>
                group.Identity.NormKey == "NORM-020.0200" &&
                group.Identity.VariantCode == "density-1" &&
                group.Identity.Environment == EstimateWorkEnvironment.Land &&
                group.Identity.LaborAudience == MachineRateAudience.NonStateSalary);
            AssertEqual(2, shared.Rows.Count);
            AssertTrue(shared.Identity.RateId.StartsWith("DG-", StringComparison.Ordinal));
            AssertEqual(shared.Identity.RateId, new EstimateRateIdentity(
                "BQP-RPBM-2025",
                "2.0.1",
                new string('A', 64),
                "NORM-020.0200",
                "density-1",
                EstimateWorkEnvironment.Land,
                MachineRateAudience.NonStateSalary,
                nonState,
                Array.Empty<string>()).RateId);

            EstimateRatePlan missingState = EstimateRatePlan.Build(
                workspace,
                "BQP-RPBM-2025",
                "2.0.1",
                new string('A', 64),
                new[] { nonState });
            AssertFalse(missingState.IsValid);
            AssertTrue(missingState.Errors.Any(error => error.StartsWith("r5:", StringComparison.Ordinal)));
        }

        private static EstimateWorkspace CreateEstimateWorkspace()
        {
            var source = new EstimateSourceBinding(
                "SheetEstimate",
                "Phu luc DT",
                "$A$1:$K$7",
                2,
                7,
                new EstimateColumnMap(1, 2, 3, 4, 0, 5, 6, 7, 8, 9, 10));
            return new EstimateWorkspace(
                source,
                new[]
                {
                    WorkspaceRow("r1", 2, "CT.01", "=B2*C2", 10m,
                        EstimateWorkEnvironment.Land, MachineRateAudience.NonStateSalary,
                        "NORM-020.0200", "density-1"),
                    WorkspaceRow("r2", 3, "CT.02", "20", 20m,
                        EstimateWorkEnvironment.Land, MachineRateAudience.NonStateSalary,
                        "NORM-020.0200", "density-1"),
                    WorkspaceRow("r3", 4, "CT.03", "30", 30m,
                        EstimateWorkEnvironment.Land, MachineRateAudience.NonStateSalary,
                        "NORM-020.0200", "density-2"),
                    WorkspaceRow("r4", 5, "CT.04", "40", 40m,
                        EstimateWorkEnvironment.Water, MachineRateAudience.NonStateSalary,
                        "NORM-020.0200", "density-1"),
                    WorkspaceRow("r5", 6, "CT.05", "50", 50m,
                        EstimateWorkEnvironment.Land, MachineRateAudience.StateBudgetSalary,
                        "NORM-020.0200", "density-1"),
                    WorkspaceRow("r6", 7, string.Empty, string.Empty, 0m,
                        EstimateWorkEnvironment.Land, MachineRateAudience.NonStateSalary,
                        string.Empty, string.Empty)
                },
                new DateTime(2026, 8, 4, 1, 2, 3, DateTimeKind.Utc));
        }

        private static EstimateWorkspaceRow WorkspaceRow(
            string id,
            int row,
            string code,
            string quantityFormula,
            decimal quantity,
            EstimateWorkEnvironment environment,
            MachineRateAudience audience,
            string normKey,
            string variant)
        {
            return new EstimateWorkspaceRow(
                id,
                "TTBMVN_PLDT_" + id,
                row,
                code,
                code.Length == 0 ? "Nhom van ban" : "Cong tac " + code,
                code.Length == 0 ? string.Empty : "ha",
                quantityFormula,
                quantity,
                string.Empty,
                0m,
                environment,
                audience,
                normKey,
                variant,
                Array.Empty<string>());
        }

        private static void TestCostSummaryCalculation()
        {
            CostRuleCatalog catalog = CostRuleCatalog.Load(LoadBqp2025CostRuleModule());
            var request = new CostSummaryCalculationRequest(
                CostSummaryTemplate.OtherFunding,
                12540585.6195888m,
                268414820.1m,
                558860046.122220m,
                "urban-residential",
                10.46m,
                CostProjectKind.Linear,
                CostConstructionKind.AgricultureAndEnvironment,
                1001m,
                5.5m,
                8m,
                CostComponentSelection.All);

            CostSummaryCalculationResult result = CostSummaryCalculator.Calculate(catalog, request);
            AssertEqual(12540586L, result.DirectCost.MaterialVnd);
            AssertEqual(268414820L, result.DirectCost.LaborVnd);
            AssertEqual(558860046L, result.DirectCost.MachineVnd);
            AssertEqual(839815452L, result.DirectCost.DirectVnd);
            AssertEqual(107365928L, result.DirectCost.CommonVnd);
            AssertEqual(52094976L, result.PreTaxIncomeVnd);
            AssertEqual(999276356L, result.ZVnd);
            AssertEqual(24981909L, result.FindRequired("K1").AppliedAmountVnd);
            AssertEqual(18475940L, result.FindRequired("K2").AppliedAmountVnd);
            AssertEqual(4996382L, result.FindRequired("K3").AppliedAmountVnd);
            AssertEqual(9992764L, result.FindRequired("K4").AppliedAmountVnd);
            AssertEqual(2.598m, result.FindRequired("K5").CalculatedRatePercent.Value);
            AssertEqual(25961200L, result.FindRequired("K5").AppliedAmountVnd);
            AssertEqual(29978291L, result.FindRequired("K6").AppliedAmountVnd);
            AssertEqual(114386486L, result.OtherCostTotalVnd);
            AssertEqual(1113662842L, result.BeforeTaxVnd);
            AssertEqual(1098673696L, result.TaxableBaseVnd);
            AssertEqual(87893896L, result.VatVnd);
            AssertEqual(1201556738L, result.AfterTaxVnd);
            AssertEqual(1201557000L, result.RoundedAfterTaxVnd);

            CostSummaryCalculationResult stateFunded = CostSummaryCalculator.Calculate(
                catalog,
                new CostSummaryCalculationRequest(
                    CostSummaryTemplate.IndependentStateFunded,
                    request.MaterialVnd,
                    request.LaborVnd,
                    request.MachineVnd,
                    request.Terrain,
                    request.AreaHa,
                    request.ProjectKind,
                    request.ConstructionKind,
                    request.DisposalWeightKg,
                    99m,
                    99m,
                    request.SelectedComponents));
            AssertEqual(0L, stateFunded.PreTaxIncomeVnd);
            AssertEqual(0L, stateFunded.VatVnd);
            AssertEqual(stateFunded.BeforeTaxVnd, stateFunded.AfterTaxVnd);
        }

        private static void TestCostSummaryValidation()
        {
            CostRuleCatalog catalog = CostRuleCatalog.Load(LoadBqp2025CostRuleModule());
            AssertThrows<ArgumentOutOfRangeException>(() => CostSummaryCalculator.Calculate(
                catalog,
                new CostSummaryCalculationRequest(
                    CostSummaryTemplate.OtherFunding,
                    1m, 1m, 1m, "plain-open", 10m,
                    CostProjectKind.Linear, CostConstructionKind.Civil,
                    999m, 5.5m, 8m, CostComponentSelection.None)));
            AssertThrows<ArgumentException>(() => CostSummaryCalculator.Calculate(
                catalog,
                new CostSummaryCalculationRequest(
                    CostSummaryTemplate.OtherFunding,
                    1m, 1m, 1m, "plain-open", 10m,
                    CostProjectKind.Linear, CostConstructionKind.Civil,
                    999m, 5.5m, 8m, CostComponentSelection.K1,
                    new[] { new CostRuleOverride("K1", 1L, string.Empty) })));
            AssertThrows<InvalidOperationException>(() => CostSummaryCalculator.Calculate(
                catalog,
                new CostSummaryCalculationRequest(
                    CostSummaryTemplate.OtherFunding,
                    1m, 1m, 1m, "plain-open", 10m,
                    CostProjectKind.Linear, CostConstructionKind.Civil,
                    1000m, 5.5m, 8m, CostComponentSelection.K6)));
            AssertEqual(2000L, CostSummaryCalculator.RoundToThousand(1500m));
            AssertEqual(1000L, CostSummaryCalculator.RoundToThousand(1499m));
        }

        private static void TestResultAuditRoundTrip()
        {
            ResultAuditEntry entry = CreateAuditEntry(
                "EST|12", "EstimateAppendix", ResultAuditKind.EstimateLine,
                12, 6, 12, 14, "NORM-020.0200", "density-1");
            var trail = new ResultAuditTrail(new[] { entry });
            string payload = ResultAuditTrailSerializer.Serialize(trail);
            string payloadAgain = ResultAuditTrailSerializer.Serialize(trail);
            AssertEqual(payload, payloadAgain);
            AssertEqual(64, ResultAuditTrailSerializer.ComputeChecksum(payload).Length);

            ResultAuditTrail restored = ResultAuditTrailSerializer.Deserialize(payload);
            AssertEqual(1, restored.Entries.Count);
            AssertEqual(entry.EntryId, restored.Entries[0].EntryId);
            AssertEqual(entry.NormKey, restored.Entries[0].NormKey);
            AssertEqual(3, restored.Entries[0].Sources.Count);
            AssertEqual("TT101-2025-BQP", restored.Entries[0].Sources[0].DocumentId);
        }

        private static void TestResultAuditValidation()
        {
            ResultAuditEntry valid = CreateAuditEntry(
                "EST|12", "EstimateAppendix", ResultAuditKind.EstimateLine,
                12, 6, 12, 14, "NORM-020.0200", "density-1");
            AssertTrue(ResultAuditTrailValidator.Validate(new ResultAuditTrail(new[] { valid })).IsValid);

            var invalid = new ResultAuditEntry(
                string.Empty, string.Empty, (ResultAuditKind)99, string.Empty, string.Empty,
                0, 2, 0, 1, string.Empty, string.Empty, string.Empty, "BAD",
                "PROFILE", string.Empty, string.Empty, string.Empty, string.Empty,
                new[] { new ResultAuditSource((ResultAuditSourceKind)99, string.Empty, -1, -2, string.Empty, string.Empty) });
            ResultAuditTrailValidationResult result = ResultAuditTrailValidator.Validate(
                new ResultAuditTrail(new[] { invalid }));
            AssertFalse(result.IsValid);
            AssertTrue(result.Errors.Count >= 5);
            AssertThrows<InvalidDataException>(() =>
                ResultAuditTrailSerializer.Deserialize("TTBMVN_RESULT_AUDIT\nschema=1\nentries=1\ninvalid"));
        }

        private static void TestResultAuditLookup()
        {
            ResultAuditEntry block = CreateAuditEntry(
                "EST|BLOCK", "EstimateAppendix", ResultAuditKind.EstimateBlock,
                11, 6, 27, 14, string.Empty, string.Empty);
            ResultAuditEntry line = CreateAuditEntry(
                "EST|12", "EstimateAppendix", ResultAuditKind.EstimateLine,
                12, 6, 12, 14, "NORM-020.0200", "density-1");
            var trail = new ResultAuditTrail(new[] { block, line });
            AssertEqual("EST|12", trail.FindMostSpecific("Sheet4", 12, 9).EntryId);
            AssertEqual("EST|BLOCK", trail.FindMostSpecific("Sheet4", 27, 9).EntryId);
            AssertEqual(null, trail.FindMostSpecific("Sheet4", 30, 9));

            ResultAuditEntry replacement = CreateAuditEntry(
                "EST|20", "EstimateAppendix", ResultAuditKind.EstimateLine,
                20, 6, 20, 14, "NORM-030.0100", "water-0.5-12");
            ResultAuditTrail replaced = trail.ReplaceScope("EstimateAppendix", new[] { replacement });
            AssertEqual(1, replaced.Entries.Count);
            AssertEqual("EST|20", replaced.Entries[0].EntryId);
        }

        private static ResultAuditEntry CreateAuditEntry(
            string id,
            string scope,
            ResultAuditKind kind,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            string normKey,
            string variant)
        {
            string packageChecksum = new string('A', 64);
            string profileChecksum = new string('B', 64);
            return new ResultAuditEntry(
                id, scope, kind, "Sheet4", "EstimateAppendix",
                firstRow, firstColumn, lastRow, lastColumn, "Ket qua " + id,
                "BQP-RPBM-2025", "2.0.1", packageChecksum,
                "PRICE-HANOI", "1.0.0", profileChecksum, normKey, variant,
                new[]
                {
                    new ResultAuditSource(
                        ResultAuditSourceKind.Regulation,
                        "TT101-2025-BQP", 32, 33, "Bang 02", string.Empty),
                    new ResultAuditSource(
                        ResultAuditSourceKind.Price,
                        string.Empty, 0, 0, string.Empty, "VL-NC-M!F81"),
                    new ResultAuditSource(
                        ResultAuditSourceKind.Quantity,
                        string.Empty, 0, 0, string.Empty, "Sheet4!D12")
                });
        }

        private static void TestWorkbookValidationReport()
        {
            var report = new WorkbookValidationReport(new[]
            {
                new WorkbookValidationIssue(
                    WorkbookValidationIssueCode.BrokenName,
                    WorkbookValidationSeverity.Warning,
                    string.Empty, string.Empty, string.Empty, "LegacyName",
                    "Name bi hong.", "Xoa name neu khong dung.", true),
                new WorkbookValidationIssue(
                    WorkbookValidationIssueCode.IncorrectTotal,
                    WorkbookValidationSeverity.Error,
                    "EstimateAppendix", "Sheet4", "K27", "Tong may",
                    "Tong khong khop.", "Tinh va ghi lai phu luc."),
                new WorkbookValidationIssue(
                    WorkbookValidationIssueCode.StaleFormula,
                    WorkbookValidationSeverity.Warning,
                    "EstimateAppendix", "Sheet4", "K22", "Cong thuc",
                    "Cong thuc cu.", "Ghi lai phu luc.")
            });
            AssertFalse(report.IsValid);
            AssertEqual(1, report.ErrorCount);
            AssertEqual(2, report.WarningCount);
            AssertEqual(WorkbookValidationIssueCode.IncorrectTotal, report.Issues[0].Code);
            AssertEqual(1, report.Find(WorkbookValidationIssueCode.StaleFormula).Count);
        }

        private static void TestWorkbookValidationInherited()
        {
            var inherited = new WorkbookValidationIssue(
                WorkbookValidationIssueCode.BrokenName,
                WorkbookValidationSeverity.Error,
                string.Empty, string.Empty, string.Empty, "LegacyName",
                "Name bi hong.", "Xem lai name.", true);
            var report = new WorkbookValidationReport(new[] { inherited });
            AssertTrue(report.IsValid);
            AssertEqual(0, report.ErrorCount);
            AssertEqual(1, report.WarningCount);
            AssertFalse(inherited.IsBlocking);
        }

        private static void TestVietnameseMoneyWords()
        {
            AssertEqual("Không đồng chẵn.", VietnameseMoneyWords.ToWords(0));
            AssertEqual(
                "Một tỷ, hai trăm linh một triệu, năm trăm năm mươi bảy nghìn đồng chẵn.",
                VietnameseMoneyWords.ToWords(1201557000L));
            AssertEqual("Một trăm linh năm đồng chẵn.", VietnameseMoneyWords.ToWords(105L));
            AssertThrows<ArgumentOutOfRangeException>(() => VietnameseMoneyWords.ToWords(-1L));
        }

        private static PriceProfile CreateUnitRatePriceProfile(
            bool includeLandDetector = true,
            string landDetectorUnit = "shift")
        {
            DateTime date = new DateTime(2026, 8, 3);
            var entries = new List<PriceProfileEntry>
            {
                new PriceProfileEntry("MAT-CONCRETE-STAKE", PriceResourceKind.Material,
                    "Coc be tong", "each", 100000m, date, PriceSourceKind.WorkbookImport, "VL-NC-M!F81"),
                new PriceProfileEntry("MAT-WOOD-STAKE", PriceResourceKind.Material,
                    "Coc go", "each", 8000m, date, PriceSourceKind.WorkbookImport, "VL-NC-M!F82"),
                new PriceProfileEntry("MAT-ROPE-10MM", PriceResourceKind.Material,
                    "Day thung", "m", 2000m, date, PriceSourceKind.WorkbookImport, "VL-NC-M!F83"),
                new PriceProfileEntry("MAT-RED-FLAG", PriceResourceKind.Material,
                    "Co do", "each", 2500m, date, PriceSourceKind.WorkbookImport, "VL-NC-M!F84"),
                new PriceProfileEntry("MAT-RED-FLAG-LARGE", PriceResourceKind.Material,
                    "Co do 0,4x0,6m", "each", 126000m, date,
                    PriceSourceKind.WorkbookImport, "VL-NC-M!F85"),
                new PriceProfileEntry("LAB-QNCN-7", PriceResourceKind.Labor,
                    "Nhan cong 7/10", "worker-day", 495000m, date,
                    PriceSourceKind.WorkbookImport, "VL-NC-M!F11"),
                new PriceProfileEntry("MACHINE-M010.008", PriceResourceKind.MachineShift,
                    "May do bom duoi nuoc", "shift", 1587824m, date,
                    PriceSourceKind.WorkbookImport, "VL-NC-M!F35", new[] { "M010.008" }),
                new PriceProfileEntry("MACHINE-M010.002", PriceResourceKind.MachineShift,
                    "May do bom tren can", "shift", 1070324m, date,
                    PriceSourceKind.WorkbookImport, "VL-NC-M!F28", new[] { "M010.002" }),
                new PriceProfileEntry("MACHINE-M010.027", PriceResourceKind.MachineShift,
                    "Thuyen Composit", "shift", 995354m, date,
                    PriceSourceKind.WorkbookImport, "VL-NC-M!F51", new[] { "M010.027" }),
                new PriceProfileEntry("MACHINE-M010.029", PriceResourceKind.MachineShift,
                    "Thiet bi lan 0,5m-3m", "shift", 1105582m, date,
                    PriceSourceKind.WorkbookImport, "VL-NC-M!F62", new[] { "M010.029" })
            };
            if (includeLandDetector)
            {
                entries.Add(new PriceProfileEntry(
                    "MACHINE-M010.001",
                    PriceResourceKind.MachineShift,
                    "May do min",
                    landDetectorUnit,
                    762996m,
                    date,
                    PriceSourceKind.WorkbookImport,
                    "VL-NC-M!F23",
                    new[] { "M010.001" }));
            }
            return PriceProfile.Create(
                "PRICE-UNIT-RATE-TEST",
                "1.0.0",
                "Bang gia test don gia",
                "Ha Noi",
                date,
                MachineRateAudience.NonStateSalary,
                DateTime.SpecifyKind(date, DateTimeKind.Utc),
                entries);
        }

        private static RegulationDataModule LoadBqp2025CostRuleModule()
        {
            string root = GetRepositoryRoot();
            string bundlePath = Path.Combine(
                root,
                "data",
                "regulations",
                "packages",
                "BQP-RPBM-2025",
                "bundle");
            return RegulationPackageBundleReader.Read(bundlePath)
                .Modules[RegulationModuleKind.CostRule];
        }

        private static RegulationDataModule LoadBqp2021CostRuleModule()
        {
            string root = GetRepositoryRoot();
            string bundlePath = Path.Combine(
                root,
                "data",
                "regulations",
                "packages",
                "BQP-RPBM-2021",
                "bundle");
            return RegulationPackageBundleReader.Read(bundlePath)
                .Modules[RegulationModuleKind.CostRule];
        }

        private static void TestMachineRateCatalog()
        {
            RegulationDataModule module = LoadBqp2025MachineModule();
            MachineRateCatalog state = MachineRateCatalog.Load(
                module,
                MachineRateAudience.StateBudgetSalary);
            MachineRateCatalog nonState = MachineRateCatalog.Load(
                module,
                MachineRateAudience.NonStateSalary);

            AssertEqual(33, state.Definitions.Count);
            AssertEqual(33, nonState.Definitions.Count);
            AssertEqual("M010.001", state.FindRequiredByKey("MACHINE-M010.001").Code);
            AssertEqual("M011.001", nonState.FindRequiredByKey("MACHINE-M010.001").Code);

            MachineRateDefinition ship = nonState.FindRequiredByCode("m011.011");
            AssertEqual(2, ship.Operators.Count);
            AssertEqual(6m, ship.Operators[0].Quantity);
            AssertEqual("si-quan", ship.Operators[0].LaborCode);
            AssertEqual(20m, ship.Operators[1].Quantity);
            AssertEqual("thuy-thu", ship.Operators[1].LaborCode);

            AssertEqual(0, state.FindRequiredByCode("M010.020").Operators.Count);
            AssertEqual(1, nonState.FindRequiredByCode("M011.020").Operators.Count);
            AssertEqual(
                "bac-5-10",
                nonState.FindRequiredByCode("M011.020").Operators[0].LaborCode);
            AssertEqual(1350000m, state.FindRequiredByCode("M010.024").ReferencePriceVnd);
            AssertEqual(350000m, nonState.FindRequiredByCode("M011.024").ReferencePriceVnd);
            AssertEqual(29, nonState.FindRequiredByCode("M011.033").Source.PageTo);
        }

        private static void TestMachineRateCalculation()
        {
            MachineRateCatalog catalog = MachineRateCatalog.Load(
                LoadBqp2025MachineModule(),
                MachineRateAudience.NonStateSalary);
            MachineRatePriceProfile prices = CreateMachineRatePriceProfile();

            MachineRateResult mineDetector = MachineRateCalculator.Calculate(
                catalog.FindRequiredByCode("M011.001"),
                prices);
            AssertMachineRate(
                mineDetector,
                135918,
                60408,
                24000,
                517500,
                25170,
                762996);
            AssertEqual(351879L, mineDetector.WaitingVnd);
            AssertEqual(95375L, mineDetector.GetHourlyRateVnd(8m));

            MachineRateResult excavator = MachineRateCalculator.Calculate(
                catalog.FindRequiredByCode("M011.004"),
                prices);
            AssertMachineRate(
                excavator,
                315589,
                105196,
                538530,
                517500,
                109579,
                1586394);

            MachineRateResult underwaterDetector = MachineRateCalculator.Calculate(
                catalog.FindRequiredByCode("M011.008"),
                prices);
            AssertMachineRate(
                underwaterDetector,
                317142,
                140952,
                36000,
                1035000,
                58730,
                1587824);

            MachineRateResult compositeWithoutFuel = MachineRateCalculator.Calculate(
                catalog.FindRequiredByCode("M011.027"),
                prices,
                new MachineRateCalculationOptions(fuelCostIncludedInMaterials: true));
            AssertMachineRate(
                compositeWithoutFuel,
                55727,
                24767,
                0,
                900000,
                14860,
                995354);

            MachineRateResult diving = MachineRateCalculator.Calculate(
                catalog.FindRequiredByCode("M011.029"),
                prices);
            AssertMachineRate(
                diving,
                41792,
                13931,
                0,
                1035000,
                14859,
                1105582);

            MachineRateResult corrosive = MachineRateCalculator.Calculate(
                catalog.FindRequiredByCode("M011.029"),
                prices,
                new MachineRateCalculationOptions(corrosiveEnvironment: true));
            AssertMachineRate(
                corrosive,
                43881,
                14627,
                0,
                1035000,
                14859,
                1108367);

            // Published Bảng 04 prints repair=44 for M011.024 although formula (3) gives 52.5.
            MachineRateResult radio = MachineRateCalculator.Calculate(
                catalog.FindRequiredByCode("M011.024"),
                prices);
            AssertMachineRate(radio, 175, 53, 0, 495000, 70, 495298);
        }

        private static void TestMachineRateLocaleAndValidation()
        {
            CultureInfo oldCulture = CultureInfo.CurrentCulture;
            CultureInfo oldUiCulture = CultureInfo.CurrentUICulture;
            try
            {
                foreach (string cultureName in new[] { "vi-VN", "de-DE" })
                {
                    CultureInfo.CurrentCulture = new CultureInfo(cultureName);
                    CultureInfo.CurrentUICulture = new CultureInfo(cultureName);
                    MachineRateCatalog catalog = MachineRateCatalog.Load(
                        LoadBqp2025MachineModule(),
                        MachineRateAudience.NonStateSalary);
                    MachineRateResult result = MachineRateCalculator.Calculate(
                        catalog.FindRequiredByCode("M011.001"),
                        CreateMachineRatePriceProfile());
                    AssertEqual(762996L, result.TotalVnd);
                }
            }
            finally
            {
                CultureInfo.CurrentCulture = oldCulture;
                CultureInfo.CurrentUICulture = oldUiCulture;
            }

            RegulationDataModule valid = LoadBqp2025MachineModule();
            RegulationDataRecord first = valid.Records[0];
            var invalidRecords = valid.Records.ToList();
            invalidRecords[0] = new RegulationDataRecord(
                first.Key,
                first.RecordType,
                first.Unit,
                first.Title,
                first.Data.Replace("repairPercent=12", "repairPercent=4,8"),
                first.Source,
                first.Verification);
            RegulationDataModule invalid = RegulationDataModule.Create(
                RegulationModuleKind.MachineRate,
                valid.DataVersion,
                invalidRecords);
            AssertThrows<FormatException>(() => MachineRateCatalog.Load(
                invalid,
                MachineRateAudience.StateBudgetSalary));

            MachineRateCatalog nonState = MachineRateCatalog.Load(
                valid,
                MachineRateAudience.NonStateSalary);
            AssertThrows<KeyNotFoundException>(() => MachineRateCalculator.Calculate(
                nonState.FindRequiredByCode("M011.004"),
                new MachineRatePriceProfile(
                    Array.Empty<KeyValuePair<string, decimal>>(),
                    Array.Empty<KeyValuePair<string, decimal>>())));
        }

        private static RegulationDataModule LoadBqp2025MachineModule()
        {
            string root = GetRepositoryRoot();
            string bundlePath = Path.Combine(
                root,
                "data",
                "regulations",
                "packages",
                "BQP-RPBM-2025",
                "bundle");
            return RegulationPackageBundleReader.Read(bundlePath)
                .Modules[RegulationModuleKind.MachineRate];
        }

        private static MachineRatePriceProfile CreateMachineRatePriceProfile()
        {
            return new MachineRatePriceProfile(
                new Dictionary<string, decimal>(StringComparer.Ordinal)
                {
                    { "pin-dai", 12000m },
                    { "pin-trung", 12000m },
                    { "pin-tieu", 10000m },
                    { "diesel", 18029.126213592233009708737864m },
                    { "gasoline-E5-RON-92-II", 20000m }
                },
                new Dictionary<string, decimal>(StringComparer.Ordinal)
                {
                    { "bac-5-10", 450000m },
                    { "bac-7-10", 495000m },
                    { "bac-8-10", 517500m },
                    { "si-quan", 569500m },
                    { "thuy-thu", 524500m }
                });
        }

        private static void AssertMachineRate(
            MachineRateResult actual,
            long depreciation,
            long repair,
            long fuel,
            long labor,
            long other,
            long total)
        {
            AssertEqual(depreciation, actual.DepreciationVnd);
            AssertEqual(repair, actual.RepairVnd);
            AssertEqual(fuel, actual.FuelVnd);
            AssertEqual(labor, actual.OperatorLaborVnd);
            AssertEqual(other, actual.OtherVnd);
            AssertEqual(total, actual.TotalVnd);
        }

        private static void TestBqp2021To2025RecordDiff()
        {
            string root = GetRepositoryRoot();
            RegulationPackageBundle oldBundle = RegulationPackageBundleReader.Read(Path.Combine(
                root, "data", "regulations", "packages", "BQP-RPBM-2021", "bundle"));
            RegulationPackageBundle newBundle = RegulationPackageBundleReader.Read(Path.Combine(
                root, "data", "regulations", "packages", "BQP-RPBM-2025", "bundle"));
            RegulationDataBundleDiffResult diff = RegulationDataBundleDiffer.Compare(oldBundle, newBundle);

            AssertTrue(diff.HasChanges);
            AssertEqual(248, diff.Changes.Count);
            AssertEqual(50, diff.AddedCount);
            AssertEqual(0, diff.RemovedCount);
            AssertEqual(198, diff.ChangedCount);
            AssertRecordFieldChanged(diff, RegulationModuleKind.Norm, "NORM-020.0500", "Data");
            AssertRecordFieldChanged(diff, RegulationModuleKind.CostRule, "COST-K2-LINEAR", "Data");
            AssertRecordFieldChanged(diff, RegulationModuleKind.MachineRate, "MACHINE-M010.004", "Data");
            AssertRecordChange(
                diff,
                RegulationModuleKind.Geography,
                "GEO-LOCALITY-HA-NOI",
                RegulationDataRecordChangeKind.Added);

            string auditPath = Path.Combine(
                root, "data", "regulations", "packages", "BQP-RPBM-2025", "audit", "change-reasons.tsv");
            string[] lines = File.ReadAllLines(auditPath, new UTF8Encoding(false, true));
            AssertEqual(249, lines.Length);
            var auditKeys = new HashSet<string>(StringComparer.Ordinal);
            for (int index = 1; index < lines.Length; index++)
            {
                string[] fields = lines[index].Split('\t');
                AssertEqual(9, fields.Length);
                AssertFalse(string.IsNullOrWhiteSpace(fields[4]));
                AssertFalse(string.IsNullOrWhiteSpace(fields[5]));
                AssertEqual("VerifiedAgainstOfficialSource", fields[8]);
                AssertTrue(auditKeys.Add(fields[0] + "\t" + fields[1]));
            }
            foreach (RegulationDataRecordChange change in diff.Changes)
                AssertTrue(auditKeys.Contains(change.ModuleKind + "\t" + change.Key));
        }

        private static void AssertRecordChange(
            RegulationDataBundleDiffResult diff,
            RegulationModuleKind module,
            string key,
            RegulationDataRecordChangeKind kind)
        {
            RegulationDataRecordChange change = diff.Changes.FirstOrDefault(item =>
                item.ModuleKind == module && string.Equals(item.Key, key, StringComparison.Ordinal));
            if (change == null)
                throw new InvalidOperationException("Missing record diff " + module + "/" + key + ".");
            AssertEqual(kind, change.Kind);
        }

        private static void AssertRecordFieldChanged(
            RegulationDataBundleDiffResult diff,
            RegulationModuleKind module,
            string key,
            string field)
        {
            RegulationDataRecordChange change = diff.Changes.FirstOrDefault(item =>
                item.ModuleKind == module && string.Equals(item.Key, key, StringComparison.Ordinal));
            if (change == null || !change.FieldChanges.Any(item =>
                string.Equals(item.Field, field, StringComparison.Ordinal)))
            {
                throw new InvalidOperationException(
                    "Missing record field diff " + module + "/" + key + "/" + field + ".");
            }
        }

        private static void AssertCodeRange(
            RegulationDataModule module,
            string prefix,
            int first,
            int last,
            int step)
        {
            for (int code = first; code <= last; code += step)
                FindRecord(module, prefix + code.ToString("0000", CultureInfo.InvariantCulture));
        }

        private static RegulationDataRecord FindRecord(RegulationDataModule module, string key)
        {
            RegulationDataRecord result = module.Records.FirstOrDefault(
                record => string.Equals(record.Key, key, StringComparison.Ordinal));
            if (result == null)
                throw new InvalidOperationException("Missing regulation record " + key + ".");
            return result;
        }

        private static Dictionary<string, string> ParseRecordData(string payload)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string field in payload.Split(';'))
            {
                int separator = field.IndexOf('=');
                if (separator <= 0 || separator == field.Length - 1)
                    throw new InvalidOperationException("Invalid record data field: " + field + ".");
                string key = field.Substring(0, separator);
                if (result.ContainsKey(key))
                    throw new InvalidOperationException("Duplicate record data field: " + key + ".");
                result.Add(key, field.Substring(separator + 1));
            }
            return result;
        }

        private static double ParseInvariantDouble(string value)
        {
            return double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture);
        }

        private static string GetRepositoryRoot()
        {
            return Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", ".."));
        }

        private static void CopyDirectory(string source, string destination)
        {
            Directory.CreateDirectory(destination);
            foreach (string directory in Directory.GetDirectories(source, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(destination + directory.Substring(source.Length));
            }
            foreach (string file in Directory.GetFiles(source, "*", SearchOption.AllDirectories))
            {
                File.Copy(file, destination + file.Substring(source.Length));
            }
        }

        private static string ReplaceFirst(string value, string oldValue, string newValue)
        {
            int index = value.IndexOf(oldValue, StringComparison.Ordinal);
            if (index < 0)
                throw new InvalidOperationException("Text to replace was not found.");
            return value.Substring(0, index) + newValue + value.Substring(index + oldValue.Length);
        }

        private static RegulationPackage CreateTestRegulationPackage()
        {
            return RegulationPackage.Create(
                "BQP-RPBM-2021",
                "1.0.0",
                new DateTime(2021, 11, 5),
                new DateTime(2025, 10, 27),
                RegulationPackageStatus.Superseded,
                "Ap dung truoc goi hop nhat 2025.",
                CreateTestRegulationSources(),
                CreateTestRegulationModules());
        }

        private static List<RegulationPackageSourceDocument> CreateTestRegulationSources()
        {
            return new List<RegulationPackageSourceDocument>
            {
                new RegulationPackageSourceDocument(
                    "TT123-2021-BQP",
                    "Dinh muc du toan va quan ly chi phi RPBM",
                    "Bo Quoc phong",
                    new DateTime(2021, 10, 18),
                    new DateTime(2021, 11, 5),
                    null,
                    "https://vbpl.vn/boquocphong/Pages/vbpq-toanvan.aspx?ItemID=149809",
                    new string('C', 64)),
                new RegulationPackageSourceDocument(
                    "TT121-2021-BQP",
                    "Quy trinh ky thuat RPBM",
                    "Bo Quoc phong",
                    new DateTime(2021, 10, 18),
                    new DateTime(2021, 11, 5),
                    null,
                    "https://vbpl.vn/TW/Pages/vbpq-toanvan.aspx?ItemID=149773",
                    new string('A', 64)),
                new RegulationPackageSourceDocument(
                    "TT122-2021-BQP",
                    "Don gia ca may va thiet bi RPBM",
                    "Bo Quoc phong",
                    new DateTime(2021, 10, 18),
                    new DateTime(2021, 11, 5),
                    null,
                    "https://vbpl.vn/TW/Pages/vbpq-toanvan.aspx?ItemID=149812",
                    new string('B', 64))
            };
        }

        private static List<RegulationPackageModuleManifest> CreateTestRegulationModules()
        {
            var result = new List<RegulationPackageModuleManifest>();
            foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
            {
                result.Add(new RegulationPackageModuleManifest(
                    kind,
                    "MODULE-" + kind.ToString().ToUpperInvariant(),
                    1,
                    "1.0.0",
                    (int)kind * 10,
                    new string(((int)kind).ToString()[0], 64)));
            }
            return result;
        }

        private static void TestRegulationPackageStoreInstall()
        {
            string testRoot = CreateTemporaryDirectory("store-install");
            try
            {
                string bundle = CreatePackageBundle(
                    testRoot,
                    "bundle-a",
                    "BQP-STORE-2021",
                    "1.0.0",
                    "A");
                var store = new RegulationPackageStore(Path.Combine(testRoot, "repository"));
                RegulationPackageInstallResult first = store.ImportFromDirectory(bundle);
                AssertEqual(RegulationPackageInstallStatus.Installed, first.Status);
                AssertTrue(File.Exists(Path.Combine(
                    first.InstallDirectory,
                    RegulationPackageLayout.InstallReceiptFileName)));

                RegulationPackageInstallResult second = store.ImportFromDirectory(bundle);
                AssertEqual(RegulationPackageInstallStatus.AlreadyInstalled, second.Status);
                AssertEqual(first.Package.PackageChecksum, second.Package.PackageChecksum);
                AssertEqual(1, store.ListInstalled().Count);

                RegulationPackage loaded = store.LoadRequired(
                    first.Package.PackageId,
                    first.Package.DataVersion,
                    first.Package.PackageChecksum.ToLowerInvariant());
                AssertEqual(first.Package.PackageChecksum, loaded.PackageChecksum);
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestRegulationPackageStoreLoadBundle()
        {
            string testRoot = CreateTemporaryDirectory("store-load-bundle");
            try
            {
                RegulationPackageBundle source = LoadBqpPackageBundle("BQP-RPBM-2025");
                var store = new RegulationPackageStore(Path.Combine(testRoot, "repository"));
                RegulationPackageInstallResult installed = store.ImportFromDirectory(source.Directory);
                RegulationPackageBundle loaded = store.LoadBundleRequired(
                    installed.Package.PackageId,
                    installed.Package.DataVersion,
                    installed.Package.PackageChecksum.ToLowerInvariant());
                AssertEqual(installed.Package.PackageChecksum, loaded.Package.PackageChecksum);
                AssertEqual(6, loaded.Modules.Count);
                AssertEqual(34, NormCatalog.Load(
                    loaded.Modules[RegulationModuleKind.Norm]).Definitions.Count);
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestRegulationPackageStoreRejectsInvalid()
        {
            string testRoot = CreateTemporaryDirectory("store-invalid");
            try
            {
                var store = new RegulationPackageStore(Path.Combine(testRoot, "repository"));

                string badChecksum = CreatePackageBundle(
                    testRoot,
                    "bad-checksum",
                    "BQP-BAD-CHECKSUM",
                    "1.0.0",
                    "B");
                File.AppendAllText(
                    Path.Combine(
                        badChecksum,
                        RegulationPackageLayout.ModulesDirectoryName,
                        RegulationPackageLayout.GetModuleFileName(RegulationModuleKind.Norm)),
                    "tampered",
                    new UTF8Encoding(false));
                AssertThrows<InvalidDataException>(() => store.ImportFromDirectory(badChecksum));

                string missingFile = CreatePackageBundle(
                    testRoot,
                    "missing-file",
                    "BQP-MISSING-FILE",
                    "1.0.0",
                    "C");
                File.Delete(Path.Combine(
                    missingFile,
                    RegulationPackageLayout.ModulesDirectoryName,
                    RegulationPackageLayout.GetModuleFileName(RegulationModuleKind.CostRule)));
                AssertThrows<FileNotFoundException>(() => store.ImportFromDirectory(missingFile));

                string oldSchema = CreatePackageBundle(
                    testRoot,
                    "old-schema",
                    "BQP-OLD-SCHEMA",
                    "1.0.0",
                    "D");
                string manifestPath = Path.Combine(oldSchema, RegulationPackageLayout.ManifestFileName);
                string manifest = File.ReadAllText(manifestPath, Encoding.UTF8)
                    .Replace("schema=1", "schema=0");
                File.WriteAllText(manifestPath, manifest, new UTF8Encoding(false));
                AssertThrows<InvalidDataException>(() => store.ImportFromDirectory(oldSchema));

                string firstBundle = CreatePackageBundle(
                    testRoot,
                    "version-first",
                    "BQP-VERSION-CONFLICT",
                    "1.0.0",
                    "E");
                RegulationPackageInstallResult installed = store.ImportFromDirectory(firstBundle);
                string conflictingBundle = CreatePackageBundle(
                    testRoot,
                    "version-second",
                    "BQP-VERSION-CONFLICT",
                    "1.0.0",
                    "F");
                AssertThrows<RegulationPackageVersionConflictException>(() =>
                    store.ImportFromDirectory(conflictingBundle));
                RegulationPackage retained = store.LoadRequired(
                    installed.Package.PackageId,
                    installed.Package.DataVersion,
                    installed.Package.PackageChecksum);
                AssertEqual(installed.Package.PackageChecksum, retained.PackageChecksum);
                AssertEqual(1, store.ListInstalled().Count);
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestRegulationPackageStoreRollback()
        {
            string testRoot = CreateTemporaryDirectory("store-rollback");
            try
            {
                string bundle = CreatePackageBundle(
                    testRoot,
                    "rollback-bundle",
                    "BQP-ROLLBACK",
                    "1.0.0",
                    "R");
                var store = new RegulationPackageStore(
                    Path.Combine(testRoot, "repository"),
                    phase =>
                    {
                        if (phase == RegulationPackageInstallPhase.BeforeCommit)
                            throw new IOException("Injected before commit.");
                    });
                AssertThrows<IOException>(() => store.ImportFromDirectory(bundle));
                AssertEqual(0, store.ListInstalled().Count);
                string staging = Path.Combine(store.RootDirectory, ".staging");
                AssertEqual(0, Directory.GetDirectories(staging).Length);
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestRegulationPackageEffectiveDateResolver()
        {
            RegulationPackage oldPackage = CreateResolverPackage(
                "BQP-RPBM-2021",
                "1.0.0",
                new DateTime(2021, 11, 5),
                new DateTime(2025, 10, 27),
                RegulationPackageStatus.Superseded,
                false);
            RegulationPackage newPackage = CreateResolverPackage(
                "BQP-RPBM-2025",
                "2.0.0",
                new DateTime(2025, 10, 28),
                null,
                RegulationPackageStatus.Published,
                true);
            RegulationPackage patchedPackage = CreateResolverPackage(
                "BQP-RPBM-2025",
                "2.0.1",
                new DateTime(2025, 10, 28),
                null,
                RegulationPackageStatus.Published,
                true);
            var packages = new[] { newPackage, patchedPackage, oldPackage };

            RegulationPackageResolutionResult before2021 = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    new DateTime(2021, 11, 4),
                    null,
                    new DateTime(2021, 11, 4)),
                packages,
                null);
            AssertEqual(RegulationPackageResolutionCode.NoApplicablePackage, before2021.Code);
            AssertFalse(before2021.IsSuccess);

            AssertResolvedPackage(
                "BQP-RPBM-2021",
                new DateTime(2021, 11, 5),
                packages);
            AssertResolvedPackage(
                "BQP-RPBM-2021",
                new DateTime(2025, 10, 27),
                packages);
            AssertResolvedPackage(
                "BQP-RPBM-2025",
                new DateTime(2025, 10, 28),
                packages);
            RegulationPackageResolutionResult patched = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    new DateTime(2025, 10, 28),
                    null,
                    new DateTime(2025, 10, 28)),
                packages,
                null);
            AssertEqual("2.0.1", patched.Package.DataVersion);

            RegulationPackage overlap = CreateResolverPackage(
                "BQP-RPBM-OVERLAP",
                "1.0.0",
                new DateTime(2025, 10, 20),
                null,
                RegulationPackageStatus.Published,
                true);
            RegulationPackageResolutionResult ambiguous = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    new DateTime(2025, 10, 22),
                    null,
                    new DateTime(2025, 10, 22)),
                new[] { oldPackage, overlap },
                null);
            AssertEqual(RegulationPackageResolutionCode.AmbiguousPackages, ambiguous.Code);
            AssertEqual(2, ambiguous.CandidatePackageIds.Count);

            RegulationPackage gapOld = CreateResolverPackage(
                "BQP-RPBM-GAP-OLD",
                "1.0.0",
                new DateTime(2021, 11, 5),
                new DateTime(2025, 10, 20),
                RegulationPackageStatus.Superseded,
                false);
            RegulationPackageResolutionResult gap = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    new DateTime(2025, 10, 25),
                    null,
                    new DateTime(2025, 10, 25)),
                new[] { gapOld, newPackage },
                null);
            AssertEqual(RegulationPackageResolutionCode.NoApplicablePackage, gap.Code);

            RegulationPackageResolutionResult invalid = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    new DateTime(2025, 10, 28, 1, 0, 0),
                    null,
                    new DateTime(2025, 10, 28)),
                packages,
                null);
            AssertEqual(RegulationPackageResolutionCode.InvalidRequest, invalid.Code);
            AssertTrue(invalid.Errors.Count > 0);

            RegulationPackageResolutionResult futureApproval = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    new DateTime(2025, 10, 28),
                    new DateTime(2025, 10, 30),
                    new DateTime(2025, 10, 29)),
                packages,
                null);
            AssertEqual(RegulationPackageResolutionCode.InvalidRequest, futureApproval.Code);
        }

        private static void TestRegulationPackageTransitionResolver()
        {
            RegulationPackage oldPackage = CreateResolverPackage(
                "BQP-RPBM-2021",
                "1.0.0",
                new DateTime(2021, 11, 5),
                new DateTime(2025, 10, 27),
                RegulationPackageStatus.Superseded,
                false);
            RegulationPackage newPackage = CreateResolverPackage(
                "BQP-RPBM-2025",
                "2.0.0",
                new DateTime(2025, 10, 28),
                null,
                RegulationPackageStatus.Published,
                true);
            RegulationTransitionRule rule = CreateTt101TransitionRule();

            RegulationPackageResolutionResult grandfathered = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    new DateTime(2025, 10, 20),
                    new DateTime(2025, 10, 25),
                    new DateTime(2026, 1, 10)),
                new[] { oldPackage, newPackage },
                new[] { rule });
            AssertTrue(grandfathered.IsSuccess);
            AssertEqual(RegulationPackageResolutionCode.ApprovedBeforeTransition, grandfathered.Code);
            AssertEqual("BQP-RPBM-2021", grandfathered.Package.PackageId);
            AssertEqual("BQP-TT101-2025-ARTICLE-6", grandfathered.RuleId);
            AssertCollectionContains(grandfathered.SourceDocumentIds, "TT101-2025-BQP");

            RegulationPackageResolutionResult approvedAfter = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    new DateTime(2025, 10, 20),
                    new DateTime(2025, 10, 29),
                    new DateTime(2025, 10, 29)),
                new[] { oldPackage, newPackage },
                new[] { rule });
            AssertTrue(approvedAfter.IsSuccess);
            AssertEqual(RegulationPackageResolutionCode.SelectedByApprovalDate, approvedAfter.Code);
            AssertEqual("BQP-RPBM-2025", approvedAfter.Package.PackageId);

            RegulationPackageResolutionResult pending = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    new DateTime(2025, 10, 20),
                    null,
                    new DateTime(2025, 10, 29)),
                new[] { oldPackage, newPackage },
                new[] { rule });
            AssertTrue(pending.IsSuccess);
            AssertEqual(RegulationPackageResolutionCode.SelectedByEvaluationDate, pending.Code);
            AssertEqual("BQP-RPBM-2025", pending.Package.PackageId);

            RegulationPackageResolutionResult unavailable = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    new DateTime(2025, 10, 20),
                    new DateTime(2025, 10, 25),
                    new DateTime(2026, 1, 10)),
                new[] { newPackage },
                new[] { rule });
            AssertEqual(
                RegulationPackageResolutionCode.RequiredTransitionPackageUnavailable,
                unavailable.Code);
            AssertFalse(unavailable.IsSuccess);

            RegulationPackage withdrawnOld = CreateResolverPackage(
                "BQP-RPBM-2021",
                "1.0.0",
                new DateTime(2021, 11, 5),
                new DateTime(2025, 10, 27),
                RegulationPackageStatus.Withdrawn,
                false);
            RegulationPackageResolutionResult withdrawn = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    new DateTime(2025, 10, 20),
                    new DateTime(2025, 10, 25),
                    new DateTime(2026, 1, 10)),
                new[] { withdrawnOld, newPackage },
                new[] { rule });
            AssertEqual(
                RegulationPackageResolutionCode.RequiredTransitionPackageUnavailable,
                withdrawn.Code);
        }

        private static void AssertResolvedPackage(
            string expectedPackageId,
            DateTime referenceDate,
            IEnumerable<RegulationPackage> packages)
        {
            RegulationPackageResolutionResult result = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(referenceDate, null, referenceDate),
                packages,
                null);
            AssertTrue(result.IsSuccess);
            AssertEqual(RegulationPackageResolutionCode.SelectedByEvaluationDate, result.Code);
            AssertEqual(expectedPackageId, result.Package.PackageId);
            AssertEqual(referenceDate, result.ReferenceDate.Value);
            AssertTrue(result.SourceDocumentIds.Count > 0);
            AssertFalse(string.IsNullOrWhiteSpace(result.Reason));
        }

        private static void TestRegulationPackagePinning()
        {
            string testRoot = CreateTemporaryDirectory("package-pin");
            try
            {
                var store = new RegulationPackageStore(Path.Combine(testRoot, "repository"));
                RegulationPackageInstallResult version1 = store.ImportFromDirectory(
                    CreatePackageBundle(
                        testRoot,
                        "pin-v1",
                        "BQP-PIN-TEST",
                        "1.0.0",
                        "P1"));

                ProjectProfile legacy = CreateTestProjectProfile();
                legacy.RegulationPackageVersion = string.Empty;
                legacy.RegulationPackageChecksum = string.Empty;
                RegulationPackagePinVerificationResult unpinned =
                    RegulationPackagePinService.Verify(legacy, store);
                AssertEqual(RegulationPackagePinStatus.UnpinnedLegacyProfile, unpinned.Status);

                ProjectProfile pinned = RegulationPackagePinService.Pin(legacy, version1.Package);
                AssertFalse(RegulationPackagePinService.IsPinned(legacy));
                AssertTrue(RegulationPackagePinService.IsPinned(pinned));
                RegulationPackagePinVerificationResult available =
                    RegulationPackagePinService.Verify(pinned, store);
                AssertTrue(available.IsAvailable);
                AssertEqual("1.0.0", available.Package.DataVersion);

                store.ImportFromDirectory(CreatePackageBundle(
                    testRoot,
                    "pin-v2",
                    "BQP-PIN-TEST",
                    "2.0.0",
                    "P2"));
                RegulationPackagePinVerificationResult withNewerInstalled =
                    RegulationPackagePinService.Verify(pinned, store);
                AssertTrue(withNewerInstalled.IsAvailable);
                AssertEqual("1.0.0", withNewerInstalled.Package.DataVersion);
                AssertEqual(version1.Package.PackageChecksum, withNewerInstalled.Package.PackageChecksum);

                var emptyStore = new RegulationPackageStore(Path.Combine(testRoot, "empty-repository"));
                RegulationPackagePinVerificationResult missing =
                    RegulationPackagePinService.Verify(pinned, emptyStore);
                AssertEqual(RegulationPackagePinStatus.Missing, missing.Status);

                File.AppendAllText(
                    Path.Combine(
                        version1.InstallDirectory,
                        RegulationPackageLayout.ModulesDirectoryName,
                        RegulationPackageLayout.GetModuleFileName(RegulationModuleKind.Norm)),
                    "corrupt",
                    new UTF8Encoding(false));
                RegulationPackagePinVerificationResult corrupt =
                    RegulationPackagePinService.Verify(pinned, store);
                AssertEqual(RegulationPackagePinStatus.Corrupt, corrupt.Status);

                ProjectProfile invalid = pinned.Clone();
                invalid.RegulationPackageChecksum = "BAD";
                RegulationPackagePinVerificationResult invalidResult =
                    RegulationPackagePinService.Verify(invalid, store);
                AssertEqual(RegulationPackagePinStatus.InvalidProfile, invalidResult.Status);
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestRegulationPackageDiffAndPlan()
        {
            RegulationPackage source = CreateResolverPackage(
                "BQP-RPBM-2021",
                "1.0.0",
                new DateTime(2021, 11, 5),
                new DateTime(2025, 10, 27),
                RegulationPackageStatus.Superseded,
                false);
            List<RegulationPackageModuleManifest> targetModules = CreateTestRegulationModules();
            int normIndex = targetModules.FindIndex(module => module.Kind == RegulationModuleKind.Norm);
            targetModules[normIndex] = new RegulationPackageModuleManifest(
                RegulationModuleKind.Norm,
                "MODULE-NORM",
                2,
                "2.0.0",
                999,
                new string('E', 64));
            List<RegulationPackageSourceDocument> targetSources = CreateTestRegulationSources();
            targetSources.Add(new RegulationPackageSourceDocument(
                "TT101-2025-BQP",
                "Sua doi cac thong tu ve RPBM",
                "Bo Quoc phong",
                new DateTime(2025, 9, 13),
                new DateTime(2025, 10, 28),
                null,
                "https://vbpl.vn/boquocphong/Pages/vbpq-toanvan.aspx?ItemID=182364",
                new string('D', 64)));
            RegulationPackage target = RegulationPackage.Create(
                "BQP-RPBM-2025",
                "2.0.0",
                new DateTime(2025, 10, 28),
                null,
                RegulationPackageStatus.Published,
                "Ap dung TT101/2025/TT-BQP.",
                targetSources,
                targetModules);

            RegulationPackageDiffResult diff = RegulationPackageDiffer.Compare(source, target);
            AssertTrue(diff.HasChanges);
            AssertTrue(diff.LegalBasisChangeCount > 0);
            AssertEqual(1, diff.CalculationDataChangeCount);
            AssertPackageChange(
                diff,
                RegulationPackageChangeKind.ModuleChanged,
                "Norm");
            AssertPackageChange(
                diff,
                RegulationPackageChangeKind.SourceAdded,
                "TT101-2025-BQP");

            RegulationPackageMigrationPlan first =
                RegulationPackageMigrationPlan.Create(source, target);
            RegulationPackageMigrationPlan second =
                RegulationPackageMigrationPlan.Create(source, target);
            AssertEqual(first.PlanId, second.PlanId);
            AssertEqual(source.PackageChecksum, first.Source.PackageChecksum);
            AssertEqual(target.PackageChecksum, first.Target.PackageChecksum);
            AssertTrue(first.Diff.HasChanges);

            RegulationPackageDiffResult noChange = RegulationPackageDiffer.Compare(source, source);
            AssertFalse(noChange.HasChanges);
            RegulationPackage draft = RegulationPackage.Create(
                "BQP-RPBM-DRAFT-MIGRATION",
                "0.1.0-preview",
                new DateTime(2027, 1, 1),
                null,
                RegulationPackageStatus.Draft,
                string.Empty,
                new RegulationPackageSourceDocument[0],
                new[] { CreateTestRegulationModules()[0] });
            AssertThrows<ArgumentException>(() =>
                RegulationPackageMigrationPlan.Create(source, draft));
        }

        private static void TestOfflineUpdateValidPackage()
        {
            string testRoot = CreateTemporaryDirectory("offline-update-valid");
            try
            {
                AssertEqual(
                    "TTBMVN-OFFLINE-2026-01",
                    OfflineUpdateTrustCatalog.Production.Single().KeyId);
                using (var signer = new RSACryptoServiceProvider(2048))
                {
                    var key = new OfflineUpdateTrustedKey(
                        "TTBMVN-TEST-UPDATE-KEY",
                        signer.ExportParameters(false));
                    string keyText = OfflineUpdatePublicKeySerializer.Serialize(key);
                    OfflineUpdateTrustedKey roundTrip = OfflineUpdatePublicKeySerializer.Deserialize(keyText);
                    AssertEqual(key.KeyId, roundTrip.KeyId);
                    AssertTrue(key.PublicKey.Modulus.SequenceEqual(roundTrip.PublicKey.Modulus));

                    string archive = CreateSignedOfflineUpdate(
                        testRoot,
                        LoadBqpPackageBundle("BQP-RPBM-2025").Directory,
                        "valid.ttbupdate",
                        signer,
                        key.KeyId,
                        "1.0.0.0");
                    OfflineUpdateVerificationResult verified = OfflineUpdatePackageVerifier.Verify(
                        archive,
                        new[] { key },
                        new Version(1, 0, 0, 0),
                        new RegulationPackage[0]);
                    AssertEqual(OfflineUpdateDisposition.Ready, verified.Disposition);
                    AssertEqual("BQP-RPBM-2025", verified.Package.PackageId);
                    AssertEqual(64, verified.ArchiveChecksum.Length);

                    string extracted = Path.Combine(testRoot, "extracted");
                    OfflineUpdatePackageVerifier.ExtractVerifiedBundle(verified, extracted);
                    RegulationPackageBundle bundle = RegulationPackageBundleReader.Read(extracted);
                    AssertEqual(verified.Package.PackageChecksum, bundle.Package.PackageChecksum);

                    OfflineUpdateVerificationResult duplicate = OfflineUpdatePackageVerifier.Verify(
                        archive,
                        new[] { key },
                        new Version(1, 0, 0, 0),
                        new[] { verified.Package });
                    AssertEqual(OfflineUpdateDisposition.AlreadyInstalled, duplicate.Disposition);
                }
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestOfflineUpdateRejectsTamper()
        {
            string testRoot = CreateTemporaryDirectory("offline-update-tamper");
            try
            {
                using (var signer = new RSACryptoServiceProvider(2048))
                {
                    var key = new OfflineUpdateTrustedKey(
                        "TTBMVN-TEST-TAMPER-KEY",
                        signer.ExportParameters(false));
                    string archive = CreateSignedOfflineUpdate(
                        testRoot,
                        LoadBqpPackageBundle("BQP-RPBM-2025").Directory,
                        "tamper.ttbupdate",
                        signer,
                        key.KeyId,
                        "1.0.0.0");
                    TamperArchiveEntry(archive, OfflineUpdateLayout.GetModulePath(RegulationModuleKind.Norm));
                    AssertOfflineUpdateFailure(
                        OfflineUpdateVerificationFailure.PayloadMismatch,
                        () => OfflineUpdatePackageVerifier.Verify(
                            archive,
                            new[] { key },
                            new Version(1, 0, 0, 0),
                            null));
                }
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestOfflineUpdateRejectsWrongSignature()
        {
            string testRoot = CreateTemporaryDirectory("offline-update-signature");
            try
            {
                using (var trustedSigner = new RSACryptoServiceProvider(2048))
                using (var wrongSigner = new RSACryptoServiceProvider(2048))
                {
                    var trustedKey = new OfflineUpdateTrustedKey(
                        "TTBMVN-TEST-SIGNATURE-KEY",
                        trustedSigner.ExportParameters(false));
                    string archive = CreateSignedOfflineUpdate(
                        testRoot,
                        LoadBqpPackageBundle("BQP-RPBM-2025").Directory,
                        "wrong-signature.ttbupdate",
                        wrongSigner,
                        trustedKey.KeyId,
                        "1.0.0.0");
                    AssertOfflineUpdateFailure(
                        OfflineUpdateVerificationFailure.InvalidSignature,
                        () => OfflineUpdatePackageVerifier.Verify(
                            archive,
                            new[] { trustedKey },
                            new Version(1, 0, 0, 0),
                            null));
                }
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestOfflineUpdateRejectsDowngrade()
        {
            string testRoot = CreateTemporaryDirectory("offline-update-downgrade");
            try
            {
                using (var signer = new RSACryptoServiceProvider(2048))
                {
                    var key = new OfflineUpdateTrustedKey(
                        "TTBMVN-TEST-DOWNGRADE-KEY",
                        signer.ExportParameters(false));
                    string archive = CreateSignedOfflineUpdate(
                        testRoot,
                        LoadBqpPackageBundle("BQP-RPBM-2021").Directory,
                        "downgrade.ttbupdate",
                        signer,
                        key.KeyId,
                        "1.0.0.0");
                    RegulationPackage newer = CreateResolverPackage(
                        "BQP-RPBM-2021",
                        "9.0.0",
                        new DateTime(2021, 11, 5),
                        null,
                        RegulationPackageStatus.Published,
                        false);
                    OfflineUpdateVerificationResult result = OfflineUpdatePackageVerifier.Verify(
                        archive,
                        new[] { key },
                        new Version(1, 0, 0, 0),
                        new[] { newer });
                    AssertEqual(OfflineUpdateDisposition.DowngradeRejected, result.Disposition);
                    AssertFalse(result.CanInstall);

                    OfflineUpdateVerificationResult appTooOld = OfflineUpdatePackageVerifier.Verify(
                        CreateSignedOfflineUpdate(
                            testRoot,
                            LoadBqpPackageBundle("BQP-RPBM-2021").Directory,
                            "requires-new-app.ttbupdate",
                            signer,
                            key.KeyId,
                            "2.0.0.0"),
                        new[] { key },
                        new Version(1, 0, 0, 0),
                        null);
                    AssertEqual(OfflineUpdateDisposition.AppUpgradeRequired, appTooOld.Disposition);
                }
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestOfflineUpdateDuplicateVersionPolicy()
        {
            string testRoot = CreateTemporaryDirectory("offline-update-conflict");
            try
            {
                using (var signer = new RSACryptoServiceProvider(2048))
                {
                    var key = new OfflineUpdateTrustedKey(
                        "TTBMVN-TEST-CONFLICT-KEY",
                        signer.ExportParameters(false));
                    string archive = CreateSignedOfflineUpdate(
                        testRoot,
                        LoadBqpPackageBundle("BQP-RPBM-2025").Directory,
                        "conflict.ttbupdate",
                        signer,
                        key.KeyId,
                        "1.0.0.0");
                    RegulationPackage conflicting = CreateResolverPackage(
                        "BQP-RPBM-2025",
                        "2.0.1",
                        new DateTime(2025, 10, 28),
                        null,
                        RegulationPackageStatus.Published,
                        true);
                    OfflineUpdateVerificationResult result = OfflineUpdatePackageVerifier.Verify(
                        archive,
                        new[] { key },
                        new Version(1, 0, 0, 0),
                        new[] { conflicting });
                    AssertEqual(OfflineUpdateDisposition.VersionConflict, result.Disposition);
                    AssertFalse(result.CanInstall);
                }
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestRegulationPackageActivationRollback()
        {
            string testRoot = CreateTemporaryDirectory("package-activation");
            try
            {
                string storeRoot = Path.Combine(testRoot, "store");
                var packageStore = new RegulationPackageStore(storeRoot);
                RegulationPackage version1 = packageStore.ImportFromDirectory(
                    CreatePackageBundle(
                        testRoot, "activation-v1", "BQP-ACTIVATION", "1.0.0", "A1"))
                    .Package;
                RegulationPackage version2 = packageStore.ImportFromDirectory(
                    CreatePackageBundle(
                        testRoot, "activation-v2", "BQP-ACTIVATION", "2.0.0", "A2"))
                    .Package;
                var activation = new RegulationPackageActivationStore(storeRoot);
                IReadOnlyList<RegulationPackage> installed = packageStore.ListInstalled();
                AssertEqual("2.0.0", activation.SelectPreferred(installed).Single().DataVersion);

                activation.SetPreferred(version1, installed);
                AssertTrue(activation.IsPreferred(version1, installed));
                AssertFalse(activation.IsPreferred(version2, installed));
                AssertEqual("1.0.0", new RegulationPackageActivationStore(storeRoot)
                    .SelectPreferred(packageStore.ListInstalled()).Single().DataVersion);

                activation.SetPreferred(version2, installed);
                AssertEqual("2.0.0", activation.SelectPreferred(installed).Single().DataVersion);
                AssertTrue(File.Exists(Path.Combine(
                    storeRoot,
                    RegulationPackageActivationStore.StateFileName)));
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestOfflineUpdateInstallAndRestart()
        {
            string testRoot = CreateTemporaryDirectory("offline-update-install");
            try
            {
                using (var signer = new RSACryptoServiceProvider(2048))
                {
                    var key = new OfflineUpdateTrustedKey(
                        "TTBMVN-TEST-INSTALL-KEY",
                        signer.ExportParameters(false));
                    string archive = CreateSignedOfflineUpdate(
                        testRoot,
                        LoadBqpPackageBundle("BQP-RPBM-2025").Directory,
                        "install.ttbupdate",
                        signer,
                        key.KeyId,
                        "1.0.0.0");
                    string storeRoot = Path.Combine(testRoot, "store");
                    var center = new OfflineUpdateCenterService(
                        storeRoot,
                        new[] { key },
                        new Version(1, 0, 0, 0));
                    OfflineUpdateInspection inspection = center.Inspect(
                        archive,
                        LoadBqpPackageBundle("BQP-RPBM-2021").Package);
                    AssertEqual(OfflineUpdateDisposition.Ready, inspection.Verification.Disposition);
                    AssertTrue(inspection.Diff != null && inspection.Diff.HasChanges);

                    OfflineUpdateInstallResult installed = center.InstallAndActivate(inspection);
                    AssertEqual(OfflineUpdateInstallStatus.InstalledAndActivated, installed.Status);
                    AssertTrue(Directory.Exists(installed.InstallDirectory));
                    AssertTrue(center.IsPreferred(installed.Package));

                    var restarted = new OfflineUpdateCenterService(
                        storeRoot,
                        new[] { key },
                        new Version(1, 0, 0, 0));
                    AssertEqual(1, restarted.ListInstalled().Count);
                    AssertEqual(installed.Package.PackageChecksum, restarted.ListPreferred().Single().PackageChecksum);
                    OfflineUpdateInspection duplicate = restarted.Inspect(archive, installed.Package);
                    AssertEqual(OfflineUpdateDisposition.AlreadyInstalled, duplicate.Verification.Disposition);
                    AssertEqual(
                        OfflineUpdateInstallStatus.AlreadyInstalledAndActivated,
                        restarted.InstallAndActivate(duplicate).Status);
                }
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private static void TestOfflineUpdateProviderOfflineOnly()
        {
            string testRoot = CreateTemporaryDirectory("offline-update-provider");
            try
            {
                using (var signer = new RSACryptoServiceProvider(2048))
                {
                    var key = new OfflineUpdateTrustedKey(
                        "TTBMVN-TEST-PROVIDER-KEY",
                        signer.ExportParameters(false));
                    var center = new OfflineUpdateCenterService(
                        Path.Combine(testRoot, "store"),
                        new[] { key },
                        new Version(1, 0, 0, 0));
                    RegulationUpdateProviderResult disabled = center.CheckProvider(
                        new DisabledOnlineRegulationUpdateProvider());
                    AssertEqual(RegulationUpdateProviderStatus.Disabled, disabled.Status);
                    RegulationUpdateProviderResult networkLost = center.CheckProvider(
                        new ThrowingOnlineUpdateProvider());
                    AssertEqual(RegulationUpdateProviderStatus.Unavailable, networkLost.Status);

                    string archive = CreateSignedOfflineUpdate(
                        testRoot,
                        LoadBqpPackageBundle("BQP-RPBM-2025").Directory,
                        "offline-after-network-loss.ttbupdate",
                        signer,
                        key.KeyId,
                        "1.0.0.0");
                    AssertEqual(
                        OfflineUpdateDisposition.Ready,
                        center.Inspect(archive, null).Verification.Disposition);
                }
            }
            finally
            {
                DeleteTemporaryDirectory(testRoot);
            }
        }

        private sealed class ThrowingOnlineUpdateProvider : IRegulationUpdateProvider
        {
            public string ProviderId => "test-network";
            public bool RequiresNetwork => true;

            public RegulationUpdateProviderResult CheckForUpdates()
            {
                throw new IOException("Network unavailable.");
            }
        }

        private static string CreateSignedOfflineUpdate(
            string testRoot,
            string bundleDirectory,
            string fileName,
            RSACryptoServiceProvider signer,
            string manifestKeyId,
            string minimumAppVersion)
        {
            RegulationPackageBundle bundle = RegulationPackageBundleReader.Read(bundleDirectory);
            var sourceFiles = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                {
                    OfflineUpdateLayout.GetPackageManifestPath(),
                    Path.Combine(bundle.Directory, RegulationPackageLayout.ManifestFileName)
                }
            };
            foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
            {
                sourceFiles.Add(
                    OfflineUpdateLayout.GetModulePath(kind),
                    Path.Combine(
                        bundle.Directory,
                        RegulationPackageLayout.ModulesDirectoryName,
                        RegulationPackageLayout.GetModuleFileName(kind)));
            }
            var files = sourceFiles.Select(item => new OfflineUpdateFileManifest(
                item.Key,
                new FileInfo(item.Value).Length,
                ComputeTestFileChecksum(item.Value)));
            OfflineUpdateManifest manifest = OfflineUpdateManifest.Create(
                bundle.Package.PackageId + "-" + bundle.Package.DataVersion,
                bundle.Package.DataVersion,
                minimumAppVersion,
                bundle.Package.PackageId,
                bundle.Package.DataVersion,
                bundle.Package.PackageChecksum,
                manifestKeyId,
                files);
            byte[] manifestBytes = new UTF8Encoding(false).GetBytes(
                OfflineUpdateManifestSerializer.Serialize(manifest));
            byte[] signature = signer.SignData(
                manifestBytes,
                CryptoConfig.MapNameToOID("SHA256"));
            string archive = Path.Combine(testRoot, fileName);
            using (FileStream stream = new FileStream(archive, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, false, Encoding.UTF8))
            {
                WriteTestArchiveEntry(zip, OfflineUpdateLayout.ManifestFileName, manifestBytes);
                WriteTestArchiveEntry(zip, OfflineUpdateLayout.SignatureFileName, signature);
                foreach (KeyValuePair<string, string> item in sourceFiles.OrderBy(item => item.Key, StringComparer.Ordinal))
                    WriteTestArchiveEntry(zip, item.Key, File.ReadAllBytes(item.Value));
            }
            return archive;
        }

        private static void TamperArchiveEntry(string archivePath, string entryPath)
        {
            using (FileStream stream = new FileStream(archivePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, false, Encoding.UTF8))
            {
                ZipArchiveEntry entry = zip.GetEntry(entryPath);
                byte[] payload;
                using (Stream input = entry.Open())
                using (var memory = new MemoryStream())
                {
                    input.CopyTo(memory);
                    payload = memory.ToArray();
                }
                payload[0] ^= 0x01;
                entry.Delete();
                WriteTestArchiveEntry(zip, entryPath, payload);
            }
        }

        private static void WriteTestArchiveEntry(ZipArchive zip, string path, byte[] payload)
        {
            ZipArchiveEntry entry = zip.CreateEntry(path, CompressionLevel.Optimal);
            entry.LastWriteTime = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
            using (Stream output = entry.Open())
                output.Write(payload, 0, payload.Length);
        }

        private static string ComputeTestFileChecksum(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (SHA256 sha256 = SHA256.Create())
                return BitConverter.ToString(sha256.ComputeHash(stream)).Replace("-", string.Empty);
        }

        private static void AssertOfflineUpdateFailure(
            OfflineUpdateVerificationFailure expected,
            Action action)
        {
            try
            {
                action();
            }
            catch (OfflineUpdateVerificationException ex)
            {
                AssertEqual(expected, ex.Failure);
                return;
            }
            throw new InvalidOperationException(
                "Expected offline update failure " + expected + ".");
        }

        private static void AssertPackageChange(
            RegulationPackageDiffResult diff,
            RegulationPackageChangeKind kind,
            string key)
        {
            foreach (RegulationPackageChange change in diff.Changes)
            {
                if (change.Kind == kind && string.Equals(change.Key, key, StringComparison.Ordinal))
                    return;
            }
            throw new InvalidOperationException("Expected package change " + kind + " for " + key + ".");
        }

        private static RegulationTransitionRule CreateTt101TransitionRule()
        {
            return new RegulationTransitionRule(
                "BQP-TT101-2025-ARTICLE-6",
                new DateTime(2025, 10, 28),
                "BQP-RPBM-2021",
                "BQP-RPBM-2025",
                "TT101-2025-BQP",
                "Thong tu 101/2025/TT-BQP, Dieu 6");
        }

        private static RegulationPackage CreateResolverPackage(
            string packageId,
            string dataVersion,
            DateTime effectiveFrom,
            DateTime? effectiveTo,
            RegulationPackageStatus status,
            bool includeTt101)
        {
            List<RegulationPackageSourceDocument> sources = CreateTestRegulationSources();
            if (includeTt101)
            {
                sources.Add(new RegulationPackageSourceDocument(
                    "TT101-2025-BQP",
                    "Sua doi cac thong tu ve RPBM",
                    "Bo Quoc phong",
                    new DateTime(2025, 9, 13),
                    new DateTime(2025, 10, 28),
                    null,
                    "https://vbpl.vn/boquocphong/Pages/vbpq-toanvan.aspx?ItemID=182364",
                    new string('D', 64)));
            }
            return RegulationPackage.Create(
                packageId,
                dataVersion,
                effectiveFrom,
                effectiveTo,
                status,
                "Resolver test package.",
                sources,
                CreateTestRegulationModules());
        }

        private static void AssertCollectionContains(
            IEnumerable<string> values,
            string expected)
        {
            foreach (string value in values)
            {
                if (string.Equals(value, expected, StringComparison.Ordinal))
                    return;
            }
            throw new InvalidOperationException("Expected collection value " + expected + ".");
        }

        private static string CreatePackageBundle(
            string testRoot,
            string directoryName,
            string packageId,
            string dataVersion,
            string contentMarker)
        {
            string bundle = Path.Combine(testRoot, directoryName);
            string modulesDirectory = Path.Combine(bundle, RegulationPackageLayout.ModulesDirectoryName);
            Directory.CreateDirectory(modulesDirectory);
            var modules = new List<RegulationPackageModuleManifest>();
            foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
            {
                string content = "TTBMVN_TEST_MODULE|" + kind + "|" + contentMarker;
                string filePath = Path.Combine(
                    modulesDirectory,
                    RegulationPackageLayout.GetModuleFileName(kind));
                File.WriteAllText(filePath, content, new UTF8Encoding(false));
                modules.Add(new RegulationPackageModuleManifest(
                    kind,
                    "MODULE-" + kind.ToString().ToUpperInvariant(),
                    1,
                    dataVersion,
                    content.Length,
                    RegulationPackageSerializer.ComputeSha256(content)));
            }

            RegulationPackage package = RegulationPackage.Create(
                packageId,
                dataVersion,
                new DateTime(2021, 11, 5),
                null,
                RegulationPackageStatus.Published,
                "Test package store.",
                CreateTestRegulationSources(),
                modules);
            File.WriteAllText(
                Path.Combine(bundle, RegulationPackageLayout.ManifestFileName),
                RegulationPackageSerializer.Serialize(package),
                new UTF8Encoding(false));
            return bundle;
        }

        private static string CreateTemporaryDirectory(string name)
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "TTBMVN.Tests",
                name + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return path;
        }

        private static void DeleteTemporaryDirectory(string path)
        {
            if (!string.IsNullOrWhiteSpace(path) &&
                path.StartsWith(
                    Path.Combine(Path.GetTempPath(), "TTBMVN.Tests") + Path.DirectorySeparatorChar,
                    StringComparison.OrdinalIgnoreCase) &&
                Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
        }

        private static void AssertMappingIssue(
            WorksheetRoleMappingValidationResult result,
            WorksheetRoleMappingIssueCode code)
        {
            foreach (WorksheetRoleMappingIssue issue in result.Issues)
            {
                if (issue.Code == code)
                    return;
            }
            throw new InvalidOperationException("Expected mapping issue " + code + ".");
        }

        private static ProjectProfile CreateTestProjectProfile()
        {
            return new ProjectProfile
            {
                SchemaVersion = ProjectProfile.CurrentSchemaVersion,
                ProjectId = "TEST-DT-103",
                PreparedDate = new DateTime(2026, 8, 1),
                ApprovalDate = new DateTime(2026, 8, 2),
                PriceDate = new DateTime(2026, 7, 15),
                RegulationPackageId = "TEST-PACKAGE",
                RegulationPackageVersion = "1.2.3",
                RegulationPackageChecksum = new string('A', 64),
                PriceProfileId = "TEST-PRICE",
                OverrideSummary = "No business override."
            };
        }

        private static void AssertProjectProfileEqual(ProjectProfile expected, ProjectProfile actual)
        {
            AssertEqual(expected.SchemaVersion, actual.SchemaVersion);
            AssertEqual(expected.ProjectId, actual.ProjectId);
            AssertEqual(expected.PreparedDate, actual.PreparedDate);
            AssertEqual(expected.ApprovalDate, actual.ApprovalDate);
            AssertEqual(expected.PriceDate, actual.PriceDate);
            AssertEqual(expected.RegulationPackageId, actual.RegulationPackageId);
            AssertEqual(expected.RegulationPackageVersion, actual.RegulationPackageVersion);
            AssertEqual(expected.RegulationPackageChecksum, actual.RegulationPackageChecksum);
            AssertEqual(expected.PriceProfileId, actual.PriceProfileId);
            AssertEqual(expected.OverrideSummary, actual.OverrideSummary);
        }

        private static List<WorksheetRoleAssignment> CreateValidRoleAssignments()
        {
            return new List<WorksheetRoleAssignment>
            {
                new WorksheetRoleAssignment("Sheet1", "VL-NC-M", "ResourcePrices"),
                new WorksheetRoleAssignment("Sheet2", "DG Can", "UnitRateLand"),
                new WorksheetRoleAssignment("Sheet3", "DG Nuoc", "UnitRateWater"),
                new WorksheetRoleAssignment("Sheet4", "Gia DT TC", "EstimateAppendix"),
                new WorksheetRoleAssignment("Sheet5", "THKP-TC", "CostSummary"),
                new WorksheetRoleAssignment("Sheet6", "Tracuu", "NormLookupView"),
                new WorksheetRoleAssignment("Sheet7", "ChiPhi", "CostRuleView")
            };
        }

        private static void AssertHasIssue(
            WorksheetRoleValidationResult result,
            WorksheetRoleIssueCode code,
            string roleId)
        {
            foreach (WorksheetRoleValidationIssue issue in result.Issues)
            {
                if (issue.Code == code && string.Equals(issue.RoleId, roleId, StringComparison.Ordinal))
                    return;
            }

            throw new InvalidOperationException("Expected issue " + code + " for role " + roleId + ".");
        }

        private static void AssertThrows<TException>(Action action) where TException : Exception
        {
            try
            {
                action();
            }
            catch (TException)
            {
                return;
            }

            throw new InvalidOperationException("Expected exception " + typeof(TException).Name + ".");
        }

        private static void AssertNear(double expected, double actual, double tolerance)
        {
            if (Math.Abs(expected - actual) > tolerance)
                throw new InvalidOperationException($"Expected {expected}, got {actual}.");
        }

        private static void AssertEqual<T>(T expected, T actual)
        {
            if (!object.Equals(expected, actual))
                throw new InvalidOperationException($"Expected {expected}, got {actual}.");
        }

        private static void AssertTrue(bool value)
        {
            if (!value)
                throw new InvalidOperationException("Expected true.");
        }

        private static void AssertFalse(bool value)
        {
            if (value)
                throw new InvalidOperationException("Expected false.");
        }
    }
}
