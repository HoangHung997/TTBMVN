using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum EstimateAppendixIssueCode
    {
        InvalidIdentity = 1,
        DuplicateLine = 2,
        InvalidQuantity = 3,
        MissingUnitRate = 4,
        MixedPriceProfile = 5
    }

    public sealed class EstimateAppendixLineRequest
    {
        public EstimateAppendixLineRequest(
            string lineId,
            string groupKey,
            string description,
            string unit,
            decimal quantity,
            decimal acceptedQuantity,
            UnitRateCalculationResult unitRate,
            string quantitySourceReference,
            string outputReference)
        {
            LineId = (lineId ?? string.Empty).Trim();
            GroupKey = (groupKey ?? string.Empty).Trim();
            Description = (description ?? string.Empty).Trim();
            Unit = (unit ?? string.Empty).Trim();
            Quantity = quantity;
            AcceptedQuantity = acceptedQuantity;
            UnitRate = unitRate;
            QuantitySourceReference = (quantitySourceReference ?? string.Empty).Trim();
            OutputReference = (outputReference ?? string.Empty).Trim();
        }

        public string LineId { get; }
        public string GroupKey { get; }
        public string Description { get; }
        public string Unit { get; }
        public decimal Quantity { get; }
        public decimal AcceptedQuantity { get; }
        public UnitRateCalculationResult UnitRate { get; }
        public string QuantitySourceReference { get; }
        public string OutputReference { get; }
        public bool RequiresUnitRate => Quantity != 0m || AcceptedQuantity != 0m;
    }

    public sealed class EstimateAppendixCalculationRequest
    {
        public EstimateAppendixCalculationRequest(
            string packageId,
            string packageVersion,
            string packageChecksum,
            IEnumerable<EstimateAppendixLineRequest> lines)
        {
            PackageId = (packageId ?? string.Empty).Trim();
            PackageVersion = (packageVersion ?? string.Empty).Trim();
            PackageChecksum = (packageChecksum ?? string.Empty).Trim();
            Lines = new ReadOnlyCollection<EstimateAppendixLineRequest>(
                (lines ?? Enumerable.Empty<EstimateAppendixLineRequest>())
                    .Where(line => line != null)
                    .ToList());
        }

        public string PackageId { get; }
        public string PackageVersion { get; }
        public string PackageChecksum { get; }
        public IReadOnlyList<EstimateAppendixLineRequest> Lines { get; }
    }

    public sealed class EstimateAppendixValidationIssue
    {
        internal EstimateAppendixValidationIssue(
            EstimateAppendixIssueCode issueCode,
            string lineId,
            string message)
        {
            IssueCode = issueCode;
            LineId = lineId ?? string.Empty;
            Message = message ?? string.Empty;
        }

        public EstimateAppendixIssueCode IssueCode { get; }
        public string LineId { get; }
        public string Message { get; }
    }

    public sealed class EstimateAppendixValidationResult
    {
        internal EstimateAppendixValidationResult(IEnumerable<EstimateAppendixValidationIssue> issues)
        {
            Issues = new ReadOnlyCollection<EstimateAppendixValidationIssue>(issues.ToList());
        }

        public IReadOnlyList<EstimateAppendixValidationIssue> Issues { get; }
        public bool IsValid => Issues.Count == 0;
    }

    public sealed class EstimateAppendixValidationException : InvalidOperationException
    {
        internal EstimateAppendixValidationException(EstimateAppendixValidationResult validation)
            : base(string.Join(" ", validation.Issues.Select(issue => issue.Message)))
        {
            Validation = validation;
        }

        public EstimateAppendixValidationResult Validation { get; }
    }

    public sealed class EstimateAppendixLineResult
    {
        internal EstimateAppendixLineResult(EstimateAppendixLineRequest request)
        {
            Request = request;
            decimal materialRate = request.UnitRate?.MaterialAmountVnd ?? 0m;
            decimal laborRate = request.UnitRate?.LaborAmountVnd ?? 0m;
            decimal machineRate = request.UnitRate?.MachineAmountVnd ?? 0m;
            MaterialAmountVnd = request.Quantity * materialRate;
            LaborAmountVnd = request.Quantity * laborRate;
            MachineAmountVnd = request.Quantity * machineRate;
            AcceptedMaterialAmountVnd = request.AcceptedQuantity * materialRate;
            AcceptedLaborAmountVnd = request.AcceptedQuantity * laborRate;
            AcceptedMachineAmountVnd = request.AcceptedQuantity * machineRate;
        }

        public EstimateAppendixLineRequest Request { get; }
        public decimal MaterialAmountVnd { get; }
        public decimal LaborAmountVnd { get; }
        public decimal MachineAmountVnd { get; }
        public decimal TotalAmountVnd => MaterialAmountVnd + LaborAmountVnd + MachineAmountVnd;
        public decimal AcceptedMaterialAmountVnd { get; }
        public decimal AcceptedLaborAmountVnd { get; }
        public decimal AcceptedMachineAmountVnd { get; }
        public decimal AcceptedTotalAmountVnd =>
            AcceptedMaterialAmountVnd + AcceptedLaborAmountVnd + AcceptedMachineAmountVnd;
    }

    public sealed class EstimateAppendixGroupResult
    {
        internal EstimateAppendixGroupResult(
            string groupKey,
            IEnumerable<EstimateAppendixLineResult> lines)
        {
            GroupKey = groupKey;
            Lines = new ReadOnlyCollection<EstimateAppendixLineResult>(lines.ToList());
            MaterialAmountVnd = Lines.Sum(line => line.MaterialAmountVnd);
            LaborAmountVnd = Lines.Sum(line => line.LaborAmountVnd);
            MachineAmountVnd = Lines.Sum(line => line.MachineAmountVnd);
            AcceptedMaterialAmountVnd = Lines.Sum(line => line.AcceptedMaterialAmountVnd);
            AcceptedLaborAmountVnd = Lines.Sum(line => line.AcceptedLaborAmountVnd);
            AcceptedMachineAmountVnd = Lines.Sum(line => line.AcceptedMachineAmountVnd);
        }

        public string GroupKey { get; }
        public IReadOnlyList<EstimateAppendixLineResult> Lines { get; }
        public decimal MaterialAmountVnd { get; }
        public decimal LaborAmountVnd { get; }
        public decimal MachineAmountVnd { get; }
        public decimal TotalAmountVnd => MaterialAmountVnd + LaborAmountVnd + MachineAmountVnd;
        public decimal AcceptedMaterialAmountVnd { get; }
        public decimal AcceptedLaborAmountVnd { get; }
        public decimal AcceptedMachineAmountVnd { get; }
        public decimal AcceptedTotalAmountVnd =>
            AcceptedMaterialAmountVnd + AcceptedLaborAmountVnd + AcceptedMachineAmountVnd;
    }

    public sealed class EstimateAppendixCalculationResult
    {
        internal EstimateAppendixCalculationResult(
            EstimateAppendixCalculationRequest request,
            IEnumerable<EstimateAppendixLineResult> lines)
        {
            PackageId = request.PackageId;
            PackageVersion = request.PackageVersion;
            PackageChecksum = request.PackageChecksum;
            Lines = new ReadOnlyCollection<EstimateAppendixLineResult>(lines.ToList());
            Groups = new ReadOnlyCollection<EstimateAppendixGroupResult>(
                Lines.GroupBy(line => line.Request.GroupKey, StringComparer.Ordinal)
                    .Select(group => new EstimateAppendixGroupResult(group.Key, group))
                    .ToList());
            MaterialAmountVnd = Lines.Sum(line => line.MaterialAmountVnd);
            LaborAmountVnd = Lines.Sum(line => line.LaborAmountVnd);
            MachineAmountVnd = Lines.Sum(line => line.MachineAmountVnd);
            AcceptedMaterialAmountVnd = Lines.Sum(line => line.AcceptedMaterialAmountVnd);
            AcceptedLaborAmountVnd = Lines.Sum(line => line.AcceptedLaborAmountVnd);
            AcceptedMachineAmountVnd = Lines.Sum(line => line.AcceptedMachineAmountVnd);
            UnitRateProfileId = Lines.Select(line => line.Request.UnitRate)
                .Where(rate => rate != null)
                .Select(rate => rate.PriceProfileId)
                .FirstOrDefault() ?? string.Empty;
            UnitRateProfileVersion = Lines.Select(line => line.Request.UnitRate)
                .Where(rate => rate != null)
                .Select(rate => rate.PriceProfileVersion)
                .FirstOrDefault() ?? string.Empty;
            UnitRateProfileChecksum = Lines.Select(line => line.Request.UnitRate)
                .Where(rate => rate != null)
                .Select(rate => rate.PriceProfileChecksum)
                .FirstOrDefault() ?? string.Empty;
        }

        public string PackageId { get; }
        public string PackageVersion { get; }
        public string PackageChecksum { get; }
        public string UnitRateProfileId { get; }
        public string UnitRateProfileVersion { get; }
        public string UnitRateProfileChecksum { get; }
        public IReadOnlyList<EstimateAppendixLineResult> Lines { get; }
        public IReadOnlyList<EstimateAppendixGroupResult> Groups { get; }
        public decimal MaterialAmountVnd { get; }
        public decimal LaborAmountVnd { get; }
        public decimal MachineAmountVnd { get; }
        public decimal TotalAmountVnd => MaterialAmountVnd + LaborAmountVnd + MachineAmountVnd;
        public decimal AcceptedMaterialAmountVnd { get; }
        public decimal AcceptedLaborAmountVnd { get; }
        public decimal AcceptedMachineAmountVnd { get; }
        public decimal AcceptedTotalAmountVnd =>
            AcceptedMaterialAmountVnd + AcceptedLaborAmountVnd + AcceptedMachineAmountVnd;
        public string RoundingRule => "none-invariant-decimal";
    }

    public static class EstimateAppendixCalculator
    {
        public static EstimateAppendixValidationResult Validate(
            EstimateAppendixCalculationRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            var issues = new List<EstimateAppendixValidationIssue>();
            if (request.PackageId.Length == 0 ||
                request.PackageVersion.Length == 0 ||
                request.PackageChecksum.Length == 0)
            {
                issues.Add(new EstimateAppendixValidationIssue(
                    EstimateAppendixIssueCode.InvalidIdentity,
                    string.Empty,
                    "Thieu danh tinh RegulationPackage cua phu luc."));
            }

            var lineIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string profileId = null;
            string profileVersion = null;
            string profileChecksum = null;
            foreach (EstimateAppendixLineRequest line in request.Lines)
            {
                if (line.LineId.Length == 0 || line.GroupKey.Length == 0)
                {
                    issues.Add(new EstimateAppendixValidationIssue(
                        EstimateAppendixIssueCode.InvalidIdentity,
                        line.LineId,
                        "Dong phu luc phai co LineId va GroupKey."));
                }
                else if (!lineIds.Add(line.LineId))
                {
                    issues.Add(new EstimateAppendixValidationIssue(
                        EstimateAppendixIssueCode.DuplicateLine,
                        line.LineId,
                        "LineId phu luc bi trung: " + line.LineId + "."));
                }

                if (line.Quantity < 0m || line.AcceptedQuantity < 0m)
                {
                    issues.Add(new EstimateAppendixValidationIssue(
                        EstimateAppendixIssueCode.InvalidQuantity,
                        line.LineId,
                        "Khoi luong phu luc khong duoc am: " + line.LineId + "."));
                }
                if (line.RequiresUnitRate && line.UnitRate == null)
                {
                    issues.Add(new EstimateAppendixValidationIssue(
                        EstimateAppendixIssueCode.MissingUnitRate,
                        line.LineId,
                        "Dong co khoi luong phai co don gia: " + line.LineId + "."));
                }
                if (line.UnitRate == null)
                    continue;

                if (profileId == null)
                {
                    profileId = line.UnitRate.PriceProfileId;
                    profileVersion = line.UnitRate.PriceProfileVersion;
                    profileChecksum = line.UnitRate.PriceProfileChecksum;
                }
                else if (!string.Equals(profileId, line.UnitRate.PriceProfileId, StringComparison.Ordinal) ||
                    !string.Equals(profileVersion, line.UnitRate.PriceProfileVersion, StringComparison.Ordinal) ||
                    !string.Equals(profileChecksum, line.UnitRate.PriceProfileChecksum, StringComparison.Ordinal))
                {
                    issues.Add(new EstimateAppendixValidationIssue(
                        EstimateAppendixIssueCode.MixedPriceProfile,
                        line.LineId,
                        "Khong duoc tron nhieu PriceProfile trong mot phu luc."));
                }
            }
            return new EstimateAppendixValidationResult(issues);
        }

        public static EstimateAppendixCalculationResult Calculate(
            EstimateAppendixCalculationRequest request)
        {
            EstimateAppendixValidationResult validation = Validate(request);
            if (!validation.IsValid)
                throw new EstimateAppendixValidationException(validation);

            return new EstimateAppendixCalculationResult(
                request,
                request.Lines.Select(line => new EstimateAppendixLineResult(line)));
        }
    }
}
