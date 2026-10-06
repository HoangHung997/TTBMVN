using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum CostProjectKind
    {
        Linear = 1,
        Other = 2
    }

    public enum CostConstructionKind
    {
        Civil = 1,
        Industrial = 2,
        Transport = 3,
        AgricultureAndEnvironment = 4,
        TechnicalInfrastructure = 5
    }

    public sealed class CostTerrainDefinition
    {
        internal CostTerrainDefinition(
            string terrain,
            decimal k1Percent,
            decimal k4Percent,
            RegulationSourceLocator source)
        {
            Terrain = terrain;
            K1Percent = k1Percent;
            K4Percent = k4Percent;
            Source = source;
        }

        public string Terrain { get; }
        public decimal K1Percent { get; }
        public decimal K4Percent { get; }
        public decimal TotalPercent => K1Percent + K4Percent;
        public RegulationSourceLocator Source { get; }
    }

    public sealed class CostRateTableDefinition
    {
        internal CostRateTableDefinition(
            string key,
            IEnumerable<decimal> thresholdsBillion,
            IEnumerable<decimal> ratesPercent,
            string interpolation,
            string aboveMaximum,
            string currentExternalBasis,
            RegulationSourceLocator source)
        {
            Key = key;
            ThresholdsBillion = new ReadOnlyCollection<decimal>(thresholdsBillion.ToList());
            RatesPercent = new ReadOnlyCollection<decimal>(ratesPercent.ToList());
            Interpolation = interpolation;
            AboveMaximum = aboveMaximum;
            CurrentExternalBasis = currentExternalBasis;
            Source = source;
        }

        public string Key { get; }
        public IReadOnlyList<decimal> ThresholdsBillion { get; }
        public IReadOnlyList<decimal> RatesPercent { get; }
        public string Interpolation { get; }
        public string AboveMaximum { get; }
        public string CurrentExternalBasis { get; }
        public RegulationSourceLocator Source { get; }
    }

    public sealed class CostBracketDefinition
    {
        internal CostBracketDefinition(
            decimal minimumBillion,
            decimal? maximumBillionExclusive,
            decimal ratePercent,
            RegulationSourceLocator source)
        {
            MinimumBillion = minimumBillion;
            MaximumBillionExclusive = maximumBillionExclusive;
            RatePercent = ratePercent;
            Source = source;
        }

        public decimal MinimumBillion { get; }
        public decimal? MaximumBillionExclusive { get; }
        public decimal RatePercent { get; }
        public RegulationSourceLocator Source { get; }
    }

    public sealed class CostRuleCatalog
    {
        private readonly IReadOnlyDictionary<string, CostTerrainDefinition> terrains;
        private readonly IReadOnlyDictionary<CostProjectKind, CostRateTableDefinition> k2Tables;
        private readonly IReadOnlyDictionary<CostConstructionKind, CostRateTableDefinition> k5Tables;

        private CostRuleCatalog(
            decimal commonRatePercent,
            IEnumerable<CostTerrainDefinition> terrainDefinitions,
            IDictionary<CostProjectKind, CostRateTableDefinition> k2,
            IEnumerable<CostBracketDefinition> k3Brackets,
            long k3MinimumVnd,
            long k3MaximumVnd,
            IDictionary<CostConstructionKind, CostRateTableDefinition> k5,
            decimal k6Under1000Percent,
            decimal k6Over1000Percent,
            decimal minimumAreaHa,
            long minimumSurveyVnd,
            long minimumDisposalVnd,
            RegulationSourceLocator commonSource,
            RegulationSourceLocator k6Source)
        {
            CommonRatePercent = commonRatePercent;
            TerrainDefinitions = new ReadOnlyCollection<CostTerrainDefinition>(
                terrainDefinitions.OrderBy(item => item.Terrain, StringComparer.Ordinal).ToList());
            terrains = new ReadOnlyDictionary<string, CostTerrainDefinition>(
                TerrainDefinitions.ToDictionary(item => item.Terrain, StringComparer.Ordinal));
            k2Tables = new ReadOnlyDictionary<CostProjectKind, CostRateTableDefinition>(
                new Dictionary<CostProjectKind, CostRateTableDefinition>(k2));
            K3Brackets = new ReadOnlyCollection<CostBracketDefinition>(
                k3Brackets.OrderBy(item => item.MinimumBillion).ToList());
            K3MinimumVnd = k3MinimumVnd;
            K3MaximumVnd = k3MaximumVnd;
            k5Tables = new ReadOnlyDictionary<CostConstructionKind, CostRateTableDefinition>(
                new Dictionary<CostConstructionKind, CostRateTableDefinition>(k5));
            K6Under1000Percent = k6Under1000Percent;
            K6Over1000Percent = k6Over1000Percent;
            MinimumAreaHa = minimumAreaHa;
            MinimumSurveyVnd = minimumSurveyVnd;
            MinimumDisposalVnd = minimumDisposalVnd;
            CommonSource = commonSource;
            K6Source = k6Source;
        }

        public decimal CommonRatePercent { get; }
        public IReadOnlyList<CostTerrainDefinition> TerrainDefinitions { get; }
        public IReadOnlyList<CostBracketDefinition> K3Brackets { get; }
        public long K3MinimumVnd { get; }
        public long K3MaximumVnd { get; }
        public decimal K6Under1000Percent { get; }
        public decimal K6Over1000Percent { get; }
        public decimal MinimumAreaHa { get; }
        public long MinimumSurveyVnd { get; }
        public long MinimumDisposalVnd { get; }
        public RegulationSourceLocator CommonSource { get; }
        public RegulationSourceLocator K6Source { get; }
        public string MoneyRoundingRule => "nearest-vnd-away-from-zero";

        public static CostRuleCatalog Load(RegulationDataModule module)
        {
            if (module == null)
                throw new ArgumentNullException(nameof(module));
            if (module.Kind != RegulationModuleKind.CostRule)
                throw new ArgumentException("Module khong phai CostRule.", nameof(module));
            RegulationDataValidationResult validation = RegulationDataValidator.Validate(module);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(module));

            var records = module.Records.ToDictionary(record => record.Key, StringComparer.Ordinal);
            foreach (RegulationDataRecord record in module.Records)
                CostRuleRecordParser.Validate(record);
            bool legacyTt123 = module.Records.All(record =>
                record.Source != null &&
                string.Equals(record.Source.DocumentId, "TT123-2021-BQP", StringComparison.Ordinal));

            RegulationDataRecord common = Required(records, "COST-COMMON");
            Dictionary<string, string> commonData = CostRuleRecordParser.ParseFields(common);
            decimal commonRate = Decimal(common, commonData, "ratePercent");

            CostTerrainDefinition[] terrainDefinitions = module.Records
                .Where(record => record.Key.StartsWith("COST-K1K4-TERRAIN-", StringComparison.Ordinal))
                .Select(ParseTerrain)
                .ToArray();
            if (terrainDefinitions.Length != 8)
                throw new FormatException("CostRule phai co 8 bang dia hinh K1/K4.");

            var k2 = new Dictionary<CostProjectKind, CostRateTableDefinition>
            {
                { CostProjectKind.Linear, ParseTable(Required(records, "COST-K2-LINEAR"), true, legacyTt123) },
                { CostProjectKind.Other, ParseTable(Required(records, "COST-K2-OTHER"), true, legacyTt123) }
            };
            CostBracketDefinition[] k3 = new[]
            {
                ParseBracket(Required(records, "COST-K3-UNDER-1")),
                ParseBracket(Required(records, "COST-K3-1-TO-5")),
                ParseBracket(Required(records, "COST-K3-FROM-5"))
            };
            Dictionary<string, string> k3Limits = CostRuleRecordParser.ParseFields(
                Required(records, "COST-K3-LIMITS"));

            var k5 = new Dictionary<CostConstructionKind, CostRateTableDefinition>
            {
                { CostConstructionKind.Civil, ParseTable(Required(records, "COST-K5-CIVIL"), false, legacyTt123) },
                { CostConstructionKind.Industrial, ParseTable(Required(records, "COST-K5-INDUSTRIAL"), false, legacyTt123) },
                { CostConstructionKind.Transport, ParseTable(Required(records, "COST-K5-TRANSPORT"), false, legacyTt123) },
                { CostConstructionKind.AgricultureAndEnvironment, ParseTable(Required(records, "COST-K5-AGRICULTURE"), false, legacyTt123) },
                { CostConstructionKind.TechnicalInfrastructure, ParseTable(Required(records, "COST-K5-INFRASTRUCTURE"), false, legacyTt123) }
            };

            RegulationDataRecord k6Under = Required(records, "COST-K6-UNDER-1000");
            RegulationDataRecord k6Over = Required(records, "COST-K6-OVER-1000");
            Dictionary<string, string> k6UnderData = CostRuleRecordParser.ParseFields(k6Under);
            Dictionary<string, string> k6OverData = CostRuleRecordParser.ParseFields(k6Over);
            decimal surveyArea = 0m;
            decimal disposalArea = 0m;
            long surveyMinimumVnd = 0L;
            long disposalMinimumVnd = 0L;
            RegulationDataRecord surveyMinimumRecord;
            RegulationDataRecord disposalMinimumRecord;
            bool hasSurveyMinimum = records.TryGetValue(
                "COST-MIN-SURVEY-DESIGN-UNDER-4HA", out surveyMinimumRecord);
            bool hasDisposalMinimum = records.TryGetValue(
                "COST-MIN-DISPOSAL-UNDER-4HA", out disposalMinimumRecord);
            if (hasSurveyMinimum != hasDisposalMinimum)
                throw new FormatException("CostRule phai co du hai muc toi thieu K1 va K6.");
            if (hasSurveyMinimum)
            {
                Dictionary<string, string> surveyMinimum = CostRuleRecordParser.ParseFields(surveyMinimumRecord);
                Dictionary<string, string> disposalMinimum = CostRuleRecordParser.ParseFields(disposalMinimumRecord);
                surveyArea = Decimal(surveyMinimumRecord, surveyMinimum, "maxAreaHa");
                disposalArea = Decimal(disposalMinimumRecord, disposalMinimum, "maxAreaHa");
                surveyMinimumVnd = Long(surveyMinimumRecord, surveyMinimum, "minimumVnd");
                disposalMinimumVnd = Long(disposalMinimumRecord, disposalMinimum, "minimumVnd");
            }
            else if (!legacyTt123)
            {
                throw new FormatException("CostRule thieu muc toi thieu K1 va K6.");
            }
            if (surveyArea != disposalArea)
                throw new FormatException("Nguong dien tich toi thieu K1 va K6 khong khop.");

            return new CostRuleCatalog(
                commonRate,
                terrainDefinitions,
                k2,
                k3,
                Long(common, k3Limits, "minimumVnd"),
                Long(common, k3Limits, "maximumVnd"),
                k5,
                Decimal(k6Under, k6UnderData, "ratePercent"),
                Decimal(k6Over, k6OverData, "ratePercent"),
                surveyArea,
                surveyMinimumVnd,
                disposalMinimumVnd,
                common.Source,
                k6Under.Source);
        }

        public CostTerrainDefinition FindTerrain(string terrain)
        {
            if (!terrains.TryGetValue((terrain ?? string.Empty).Trim(), out CostTerrainDefinition value))
                throw new KeyNotFoundException("Khong co loai dia hinh: " + terrain + ".");
            return value;
        }

        public CostRateTableDefinition GetK2(CostProjectKind kind)
        {
            if (!k2Tables.TryGetValue(kind, out CostRateTableDefinition value))
                throw new ArgumentOutOfRangeException(nameof(kind));
            return value;
        }

        public CostRateTableDefinition GetK5(CostConstructionKind kind)
        {
            if (!k5Tables.TryGetValue(kind, out CostRateTableDefinition value))
                throw new ArgumentOutOfRangeException(nameof(kind));
            return value;
        }

        private static CostTerrainDefinition ParseTerrain(RegulationDataRecord record)
        {
            Dictionary<string, string> data = CostRuleRecordParser.ParseFields(record);
            decimal k1 = Decimal(record, data, "k1Percent");
            decimal k4 = Decimal(record, data, "k4Percent");
            if (Decimal(record, data, "totalPercent") != k1 + k4)
                throw new FormatException(record.Key + ": totalPercent khong bang K1 + K4.");
            return new CostTerrainDefinition(Required(data, record, "terrain"), k1, k4, record.Source);
        }

        private static CostRateTableDefinition ParseTable(
            RegulationDataRecord record,
            bool step,
            bool legacyTt123)
        {
            Dictionary<string, string> data = CostRuleRecordParser.ParseFields(record);
            decimal[] thresholds = DecimalList(record, Required(data, record, "thresholdBillion"));
            decimal[] rates = DecimalList(record, Required(data, record, "ratesPercent"));
            if (thresholds.Length == 0 || thresholds.Any(value => value <= 0m))
                throw new FormatException(record.Key + ": threshold phai duong.");
            for (int index = 1; index < thresholds.Length; index++)
            {
                if (thresholds[index] <= thresholds[index - 1])
                    throw new FormatException(record.Key + ": threshold phai tang nghiem ngat.");
            }
            if (rates.Length != thresholds.Length + (step ? 1 : 0))
                throw new FormatException(record.Key + ": so rate khong khop threshold.");
            string interpolation;
            if (!data.TryGetValue("interpolation", out interpolation) ||
                string.IsNullOrWhiteSpace(interpolation))
            {
                if (!legacyTt123)
                    throw new FormatException(record.Key + ": thieu field interpolation.");
                interpolation = step ? "step" : "linear";
            }
            if (interpolation != (step ? "step" : "linear"))
                throw new FormatException(record.Key + ": interpolation khong hop le.");
            data.TryGetValue("aboveMaximum", out string aboveMaximum);
            string currentExternalBasis;
            if (!data.TryGetValue("currentExternalBasis", out currentExternalBasis) ||
                string.IsNullOrWhiteSpace(currentExternalBasis))
            {
                if (!legacyTt123)
                    throw new FormatException(record.Key + ": thieu field currentExternalBasis.");
                currentExternalBasis = string.Empty;
            }
            return new CostRateTableDefinition(
                record.Key,
                thresholds,
                rates,
                interpolation,
                aboveMaximum ?? string.Empty,
                currentExternalBasis,
                record.Source);
        }

        private static CostBracketDefinition ParseBracket(RegulationDataRecord record)
        {
            Dictionary<string, string> data = CostRuleRecordParser.ParseFields(record);
            decimal minimum = Decimal(record, data, "minBillion");
            string maximumText = Required(data, record, "maxBillionExclusive");
            decimal? maximum = maximumText == "none"
                ? (decimal?)null
                : ParseDecimal(record, maximumText, "maxBillionExclusive");
            if (maximum.HasValue && maximum.Value <= minimum)
                throw new FormatException(record.Key + ": khoang K3 khong hop le.");
            return new CostBracketDefinition(
                minimum,
                maximum,
                Decimal(record, data, "ratePercent"),
                record.Source);
        }

        private static RegulationDataRecord Required(
            IDictionary<string, RegulationDataRecord> records,
            string key)
        {
            if (!records.TryGetValue(key, out RegulationDataRecord value))
                throw new FormatException("CostRule thieu record " + key + ".");
            return value;
        }

        private static string Required(
            IDictionary<string, string> fields,
            RegulationDataRecord record,
            string field)
        {
            if (!fields.TryGetValue(field, out string value) || string.IsNullOrWhiteSpace(value))
                throw new FormatException(record.Key + ": thieu field " + field + ".");
            return value;
        }

        private static decimal Decimal(
            RegulationDataRecord record,
            IDictionary<string, string> fields,
            string field)
        {
            return ParseDecimal(record, Required(fields, record, field), field);
        }

        private static long Long(
            RegulationDataRecord record,
            IDictionary<string, string> fields,
            string field)
        {
            string value = Required(fields, record, field);
            if (!long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long parsed) || parsed < 0)
                throw new FormatException(record.Key + ": " + field + " khong hop le: " + value + ".");
            return parsed;
        }

        private static decimal ParseDecimal(RegulationDataRecord record, string value, string field)
        {
            if (value.Contains(",") ||
                !decimal.TryParse(value, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal parsed) ||
                parsed < 0m)
            {
                throw new FormatException(record.Key + ": " + field + " khong hop le: " + value + ".");
            }
            return parsed;
        }

        private static decimal[] DecimalList(RegulationDataRecord record, string value)
        {
            return value.Split(',')
                .Select(item => ParseDecimal(record, item.Trim(), "decimal list"))
                .ToArray();
        }
    }

    internal static class CostRuleRecordParser
    {
        private static readonly IDictionary<string, ISet<string>> AllowedByType =
            new Dictionary<string, ISet<string>>(StringComparer.Ordinal)
            {
                { "CostComponent", Set("code", "basis", "grades") },
                { "CostFormula", Set("code", "formula", "ratePercent", "base", "rounding") },
                { "CostRate", Set("terrain", "k1Percent", "k4Percent", "totalPercent", "base", "rounding", "code", "minBillion", "maxBillionExclusive", "ratePercent", "maxWeightKgExclusive", "minWeightKgExclusive") },
                { "CostRateTable", Set("code", "base", "thresholdBillion", "ratesPercent", "boundary", "interpolation", "aboveMaximum", "rounding", "currentExternalBasis") },
                { "CostLimit", Set("code", "minimumVnd", "maximumVnd") },
                { "ExternalRule", Set("codes", "basis", "code", "approval") },
                { "CostMinimum", Set("maxAreaHa", "teamMembers", "allowancePerPersonDayVnd", "minimumVnd", "rounding") },
                { "EstimateTemplate", Set("template", "formulaT", "formulaC", "formulaZ", "formulaH", "formulaQ", "formulaVAT", "parts", "attachments") }
            };

        internal static void Validate(RegulationDataRecord record)
        {
            if (!AllowedByType.TryGetValue(record.RecordType, out ISet<string> allowed))
                throw new FormatException("CostRule record type khong duoc ho tro: " + record.RecordType + ".");
            foreach (string field in ParseFields(record).Keys)
            {
                if (!allowed.Contains(field))
                    throw new FormatException(record.Key + ": field khong duoc ho tro: " + field + ".");
            }
        }

        internal static Dictionary<string, string> ParseFields(RegulationDataRecord record)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string item in record.Data.Split(';'))
            {
                int separator = item.IndexOf('=');
                if (separator <= 0 || separator == item.Length - 1)
                    throw new FormatException(record.Key + ": data field khong hop le: " + item + ".");
                string key = item.Substring(0, separator).Trim();
                string value = item.Substring(separator + 1).Trim();
                if (result.ContainsKey(key))
                    throw new FormatException(record.Key + ": data field trung: " + key + ".");
                result.Add(key, value);
            }
            return result;
        }

        private static ISet<string> Set(params string[] values)
        {
            return new HashSet<string>(values, StringComparer.Ordinal);
        }
    }
}
