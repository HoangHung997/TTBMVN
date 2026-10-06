using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelAddIn1.Core
{
    public enum NormEnvironment
    {
        Survey = 1,
        Land = 2,
        InlandWater = 3,
        Sea = 4
    }

    public sealed class NormSearchQuery
    {
        public NormSearchQuery(
            string text,
            NormEnvironment? environment = null,
            decimal? depthMeters = null,
            string variantCode = null,
            NormResourceKind? resourceKind = null,
            int maximumResults = 100)
        {
            Text = (text ?? string.Empty).Trim();
            Environment = environment;
            DepthMeters = depthMeters;
            VariantCode = (variantCode ?? string.Empty).Trim();
            ResourceKind = resourceKind;
            MaximumResults = maximumResults;
        }

        public string Text { get; }
        public NormEnvironment? Environment { get; }
        public decimal? DepthMeters { get; }
        public string VariantCode { get; }
        public NormResourceKind? ResourceKind { get; }
        public int MaximumResults { get; }
    }

    public enum NormSearchMatchKind
    {
        FilterOnly = 0,
        ExactKey = 1,
        ExactTitle = 2,
        Contains = 3,
        TokenFuzzy = 4
    }

    public sealed class NormSearchResult
    {
        internal NormSearchResult(
            string packageId,
            string packageVersion,
            string packageChecksum,
            NormDefinition definition,
            string key,
            string title,
            string workUnit,
            RegulationSourceLocator source,
            NormEnvironment environment,
            IEnumerable<decimal> depthMeters,
            NormSearchMatchKind matchKind,
            int score)
        {
            PackageId = packageId;
            PackageVersion = packageVersion;
            PackageChecksum = packageChecksum;
            Definition = definition;
            Key = key;
            Title = title;
            WorkUnit = workUnit;
            Source = source;
            Environment = environment;
            DepthMeters = new ReadOnlyCollection<decimal>(depthMeters.ToList());
            MatchKind = matchKind;
            Score = score;
        }

        public string PackageId { get; }
        public string PackageVersion { get; }
        public string PackageChecksum { get; }
        public NormDefinition Definition { get; }
        public string Key { get; }
        public string Title { get; }
        public string WorkUnit { get; }
        public RegulationSourceLocator Source { get; }
        public bool HasDetailedRates => Definition != null;
        public NormEnvironment Environment { get; }
        public IReadOnlyList<decimal> DepthMeters { get; }
        public NormSearchMatchKind MatchKind { get; }
        public int Score { get; }
    }

    public sealed class NormSearchIndex
    {
        private static readonly Regex DepthPattern = new Regex(
            "(?<number>[0-9]+(?:[.,][0-9]+)?)\\s*m(?:\\b|$)",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        private static readonly Regex NormKeyPattern = new Regex(
            "^NORM-[0-9]{3}\\.[0-9]{4}$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        private readonly string packageId;
        private readonly string packageVersion;
        private readonly string packageChecksum;
        private readonly IReadOnlyList<SearchDocument> documents;

        public NormSearchIndex(
            string packageId,
            string packageVersion,
            string packageChecksum,
            NormCatalog catalog)
        {
            this.packageId = Require(packageId, nameof(packageId));
            this.packageVersion = Require(packageVersion, nameof(packageVersion));
            this.packageChecksum = Require(packageChecksum, nameof(packageChecksum));
            if (!RegulationPackageValidator.IsSha256(this.packageChecksum))
                throw new ArgumentException("Package checksum khong hop le.", nameof(packageChecksum));
            if (catalog == null)
                throw new ArgumentNullException(nameof(catalog));

            documents = new ReadOnlyCollection<SearchDocument>(catalog.Definitions
                .Select(definition => new SearchDocument(
                    definition,
                    definition.Key,
                    definition.Title,
                    definition.WorkUnit,
                    definition.Source,
                    definition.Variants,
                    definition.Rates,
                    ClassifyEnvironment(definition.Key),
                    ExtractDepths(definition.Title)))
                .ToList());
        }

        public NormSearchIndex(
            string packageId,
            string packageVersion,
            string packageChecksum,
            RegulationDataModule normModule)
        {
            this.packageId = Require(packageId, nameof(packageId));
            this.packageVersion = Require(packageVersion, nameof(packageVersion));
            this.packageChecksum = Require(packageChecksum, nameof(packageChecksum));
            if (!RegulationPackageValidator.IsSha256(this.packageChecksum))
                throw new ArgumentException("Package checksum khong hop le.", nameof(packageChecksum));
            if (normModule == null)
                throw new ArgumentNullException(nameof(normModule));
            if (normModule.Kind != RegulationModuleKind.Norm)
                throw new ArgumentException("Module khong phai Norm.", nameof(normModule));
            RegulationDataValidationResult validation = RegulationDataValidator.Validate(normModule);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(normModule));

            var values = new List<SearchDocument>();
            foreach (RegulationDataRecord record in normModule.Records.Where(record =>
                string.Equals(record.RecordType, "NormCatalog", StringComparison.Ordinal)))
            {
                NormDefinition definition = null;
                if (record.Data.Contains("variantCodes=") && record.Data.Contains("rates="))
                    definition = NormRecordParser.Parse(record);
                values.Add(new SearchDocument(
                    definition,
                    record.Key,
                    record.Title,
                    record.Unit,
                    record.Source,
                    definition?.Variants ?? new string[0],
                    definition?.Rates ?? new NormResourceRate[0],
                    ClassifyEnvironment(record.Key),
                    ExtractDepths(record.Title)));
            }
            if (values.Count == 0)
                throw new FormatException("Module Norm khong co NormCatalog.");
            documents = new ReadOnlyCollection<SearchDocument>(values);
        }

        public IReadOnlyList<NormSearchResult> Search(NormSearchQuery query)
        {
            if (query == null)
                throw new ArgumentNullException(nameof(query));
            if (query.MaximumResults < 1 || query.MaximumResults > 500)
                throw new ArgumentOutOfRangeException(
                    nameof(query),
                    "MaximumResults phai tu 1 den 500.");
            if (query.DepthMeters.HasValue && query.DepthMeters.Value < 0m)
                throw new ArgumentOutOfRangeException(nameof(query), "Do sau khong duoc am.");

            string normalizedQuery = Normalize(query.Text);
            string[] queryTokens = Tokenize(normalizedQuery);
            bool exactKeyQuery = documents.Any(document => string.Equals(
                document.Key,
                query.Text,
                StringComparison.OrdinalIgnoreCase));
            if (!exactKeyQuery && NormKeyPattern.IsMatch(query.Text))
                return new NormSearchResult[0];
            var results = new List<NormSearchResult>();
            foreach (SearchDocument document in documents)
            {
                if (exactKeyQuery && !string.Equals(
                    document.Key,
                    query.Text,
                    StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                if (query.Environment.HasValue &&
                    document.Environment != query.Environment.Value)
                {
                    continue;
                }
                if (query.DepthMeters.HasValue && !document.Depths.Any(depth =>
                    Math.Abs(depth - query.DepthMeters.Value) <= 0.0001m))
                {
                    continue;
                }
                if (query.VariantCode.Length > 0 && !document.Variants.Any(
                    variant => string.Equals(
                        variant,
                        query.VariantCode,
                        StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }
                if (query.ResourceKind.HasValue && !document.Rates.Any(
                    rate => rate.Kind == query.ResourceKind.Value))
                {
                    continue;
                }

                NormSearchMatchKind matchKind;
                int score;
                if (!TryScore(document, query.Text, normalizedQuery, queryTokens, out matchKind, out score))
                    continue;
                results.Add(new NormSearchResult(
                    packageId,
                    packageVersion,
                    packageChecksum,
                    document.Definition,
                    document.Key,
                    document.Title,
                    document.WorkUnit,
                    document.Source,
                    document.Environment,
                    document.Depths,
                    matchKind,
                    score));
            }

            return new ReadOnlyCollection<NormSearchResult>(results
                .OrderBy(result => result.Score)
                .ThenBy(result => result.Key, StringComparer.Ordinal)
                .Take(query.MaximumResults)
                .ToList());
        }

        public static IReadOnlyList<NormSearchResult> SearchMany(
            IEnumerable<NormSearchIndex> indexes,
            NormSearchQuery query)
        {
            if (query == null)
                throw new ArgumentNullException(nameof(query));
            var results = (indexes ?? Enumerable.Empty<NormSearchIndex>())
                .Where(index => index != null)
                .SelectMany(index => index.Search(query))
                .OrderBy(result => result.Score)
                .ThenBy(result => result.Key, StringComparer.Ordinal)
                .ThenBy(result => result.PackageId, StringComparer.Ordinal)
                .ThenBy(result => result.PackageVersion, StringComparer.Ordinal)
                .Take(query.MaximumResults)
                .ToList();
            return new ReadOnlyCollection<NormSearchResult>(results);
        }

        private static bool TryScore(
            SearchDocument document,
            string rawQuery,
            string normalizedQuery,
            string[] queryTokens,
            out NormSearchMatchKind matchKind,
            out int score)
        {
            if (normalizedQuery.Length == 0)
            {
                matchKind = NormSearchMatchKind.FilterOnly;
                score = 100;
                return true;
            }
            if (string.Equals(document.Key, rawQuery.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                matchKind = NormSearchMatchKind.ExactKey;
                score = 0;
                return true;
            }
            if (string.Equals(document.NormalizedTitle, normalizedQuery, StringComparison.Ordinal))
            {
                matchKind = NormSearchMatchKind.ExactTitle;
                score = 5;
                return true;
            }
            if (document.NormalizedKey.Contains(normalizedQuery) ||
                document.NormalizedTitle.Contains(normalizedQuery))
            {
                matchKind = NormSearchMatchKind.Contains;
                score = document.NormalizedKey.Contains(normalizedQuery) ? 10 : 20;
                return true;
            }

            int fuzzyScore = 0;
            foreach (string queryToken in queryTokens)
            {
                int best = int.MaxValue;
                foreach (string documentToken in document.Tokens)
                {
                    if (documentToken == queryToken)
                    {
                        best = 0;
                        break;
                    }
                    if (documentToken.StartsWith(queryToken, StringComparison.Ordinal) ||
                        queryToken.StartsWith(documentToken, StringComparison.Ordinal))
                    {
                        best = Math.Min(best, 1);
                        continue;
                    }
                    int maximumDistance = queryToken.Length >= 6 ? 2 : 1;
                    int distance = LevenshteinDistance(queryToken, documentToken, maximumDistance);
                    if (distance <= maximumDistance)
                        best = Math.Min(best, distance + 1);
                }
                if (best == int.MaxValue)
                {
                    matchKind = default(NormSearchMatchKind);
                    score = 0;
                    return false;
                }
                fuzzyScore += best;
            }

            matchKind = NormSearchMatchKind.TokenFuzzy;
            score = 30 + fuzzyScore;
            return true;
        }

        private static int LevenshteinDistance(string left, string right, int maximum)
        {
            if (Math.Abs(left.Length - right.Length) > maximum)
                return maximum + 1;
            int[] previous = Enumerable.Range(0, right.Length + 1).ToArray();
            int[] current = new int[right.Length + 1];
            for (int leftIndex = 1; leftIndex <= left.Length; leftIndex++)
            {
                current[0] = leftIndex;
                int rowMinimum = current[0];
                for (int rightIndex = 1; rightIndex <= right.Length; rightIndex++)
                {
                    int substitution = previous[rightIndex - 1] +
                        (left[leftIndex - 1] == right[rightIndex - 1] ? 0 : 1);
                    current[rightIndex] = Math.Min(
                        Math.Min(previous[rightIndex] + 1, current[rightIndex - 1] + 1),
                        substitution);
                    rowMinimum = Math.Min(rowMinimum, current[rightIndex]);
                }
                if (rowMinimum > maximum)
                    return maximum + 1;
                int[] swap = previous;
                previous = current;
                current = swap;
            }
            return previous[right.Length];
        }

        private static NormEnvironment ClassifyEnvironment(string key)
        {
            if (key.StartsWith("NORM-000.", StringComparison.Ordinal))
                return NormEnvironment.Survey;
            if (key.StartsWith("NORM-010.", StringComparison.Ordinal) ||
                key.StartsWith("NORM-020.", StringComparison.Ordinal))
            {
                return NormEnvironment.Land;
            }
            if (key.StartsWith("NORM-030.", StringComparison.Ordinal))
                return NormEnvironment.InlandWater;
            if (key.StartsWith("NORM-040.", StringComparison.Ordinal))
                return NormEnvironment.Sea;
            throw new FormatException("Khong phan loai duoc moi truong cho " + key + ".");
        }

        private static decimal[] ExtractDepths(string title)
        {
            var values = new List<decimal>();
            foreach (Match match in DepthPattern.Matches(Normalize(title)))
            {
                string text = match.Groups["number"].Value.Replace(',', '.');
                decimal value;
                if (decimal.TryParse(
                    text,
                    NumberStyles.AllowDecimalPoint,
                    CultureInfo.InvariantCulture,
                    out value))
                {
                    values.Add(value);
                }
            }
            return values.Distinct().OrderBy(value => value).ToArray();
        }

        private static string Normalize(string value)
        {
            string decomposed = (value ?? string.Empty)
                .Replace('đ', 'd')
                .Replace('Đ', 'D')
                .Normalize(NormalizationForm.FormD);
            var builder = new StringBuilder(decomposed.Length);
            bool previousSpace = true;
            foreach (char character in decomposed)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(character) == UnicodeCategory.NonSpacingMark)
                    continue;
                char normalized = char.ToLowerInvariant(character);
                if (char.IsLetterOrDigit(normalized) || normalized == '.' || normalized == ',')
                {
                    builder.Append(normalized);
                    previousSpace = false;
                }
                else if (!previousSpace)
                {
                    builder.Append(' ');
                    previousSpace = true;
                }
            }
            return builder.ToString().Trim();
        }

        private static string[] Tokenize(string normalized)
        {
            return (normalized ?? string.Empty)
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static string Require(string value, string parameterName)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.Length == 0)
                throw new ArgumentException("Gia tri khong duoc trong.", parameterName);
            return normalized;
        }

        private sealed class SearchDocument
        {
            public SearchDocument(
                NormDefinition definition,
                string key,
                string title,
                string workUnit,
                RegulationSourceLocator source,
                IEnumerable<string> variants,
                IEnumerable<NormResourceRate> rates,
                NormEnvironment environment,
                IEnumerable<decimal> depths)
            {
                Definition = definition;
                Key = key;
                Title = title;
                WorkUnit = workUnit;
                Source = source;
                Variants = variants.ToArray();
                Rates = rates.ToArray();
                Environment = environment;
                Depths = depths.ToArray();
                NormalizedKey = Normalize(key);
                NormalizedTitle = Normalize(title);
                string searchable = string.Join(" ", new[]
                {
                    key,
                    title,
                    workUnit,
                    string.Join(" ", Variants),
                    string.Join(" ", Rates.Select(rate => rate.ResourceCode)),
                    source.DocumentId,
                    source.Section
                });
                Tokens = Tokenize(Normalize(searchable))
                    .Distinct(StringComparer.Ordinal)
                    .ToArray();
            }

            public NormDefinition Definition { get; }
            public string Key { get; }
            public string Title { get; }
            public string WorkUnit { get; }
            public RegulationSourceLocator Source { get; }
            public string[] Variants { get; }
            public NormResourceRate[] Rates { get; }
            public NormEnvironment Environment { get; }
            public decimal[] Depths { get; }
            public string NormalizedKey { get; }
            public string NormalizedTitle { get; }
            public string[] Tokens { get; }
        }
    }
}
