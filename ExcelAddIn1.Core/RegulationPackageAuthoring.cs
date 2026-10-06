using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelAddIn1.Core
{
    public static class RegulationPackageAuthoring
    {
        public static void WriteSource(string directory, string id, string version, DateTime from, DateTime? to,
            string note, IEnumerable<RegulationPackageSourceDocument> sources,
            IDictionary<RegulationModuleKind, IReadOnlyList<RegulationDataRecord>> records)
        {
            string root = Path.GetFullPath(directory);
            if (Directory.Exists(root) || File.Exists(root)) throw new IOException("Thu muc nhap da ton tai.");
            Directory.CreateDirectory(Path.GetDirectoryName(root));
            string stage = root + "." + Guid.NewGuid().ToString("N");
            try
            {
                Directory.CreateDirectory(Path.Combine(stage, "modules"));
                File.WriteAllLines(Path.Combine(stage, "package.properties"), new[] {
                    "packageId=" + Field(id), "dataVersion=" + Field(version), "effectiveFrom=" + Date(from),
                    "effectiveTo=" + Date(to), "status=Published", "transitionNote=" + Field(note) }, new UTF8Encoding(false));
                var citations = new List<string> { "documentId\ttitle\tpublisher\tissuedDate\teffectiveFrom\teffectiveTo\tofficialUri\tcontentChecksum" };
                citations.AddRange(sources.Select(s => Row(s.DocumentId, s.Title, s.Publisher, Date(s.IssuedDate),
                    Date(s.EffectiveFrom), Date(s.EffectiveTo), s.OfficialUri, s.ContentChecksum)));
                File.WriteAllLines(Path.Combine(stage, "sources.tsv"), citations, new UTF8Encoding(false));
                foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
                {
                    var lines = new List<string> { "key\trecordType\tunit\ttitle\tdata\tsourceDocumentId\tpageFrom\tpageTo\tsection\tverification" };
                    lines.AddRange(records[kind].Select(r => Row(r.Key, r.RecordType, r.Unit, r.Title, r.Data,
                        r.Source?.DocumentId, r.Source?.PageFrom.ToString(CultureInfo.InvariantCulture),
                        r.Source?.PageTo.ToString(CultureInfo.InvariantCulture), r.Source?.Section, r.Verification.ToString())));
                    File.WriteAllLines(Path.Combine(stage, "modules", kind + ".tsv"), lines, new UTF8Encoding(false));
                }
                Directory.Move(stage, root);
            }
            finally { if (Directory.Exists(stage)) Directory.Delete(stage, true); }
        }

        public static RegulationPackageBundle BuildChecked(string source, string output)
        {
            var bundle = RegulationPackageBundleBuilder.Build(source, output);
            NormCatalog.Load(bundle.Modules[RegulationModuleKind.Norm]);
            CostRuleCatalog.Load(bundle.Modules[RegulationModuleKind.CostRule]);
            foreach (MachineRateAudience audience in Enum.GetValues(typeof(MachineRateAudience)))
                MachineRateCatalog.Load(bundle.Modules[RegulationModuleKind.MachineRate], audience);
            return bundle;
        }

        private static string Date(DateTime? value) => value?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "";
        private static string Row(params string[] values) => string.Join("\t", values.Select(Field));
        private static string Field(string value)
        {
            value = value ?? "";
            if (value.IndexOfAny(new[] { '\t', '\r', '\n' }) >= 0)
                throw new FormatException("Gia tri khong duoc chua tab/xuong dong.");
            return value;
        }
    }
}
