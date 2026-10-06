using System;

namespace ExcelAddIn1.Core
{
    public enum DaodatFeasibilityStatus
    {
        Ok,
        InvalidInput,
        BelowMinimum,
        AboveMaximum
    }

    public sealed class DaodatFeasibilityResult
    {
        public DaodatFeasibilityStatus Status { get; set; }
        public double AverageVolume { get; set; }
        public string Suggestion { get; set; }
        public bool CanGenerate => Status == DaodatFeasibilityStatus.Ok;
    }

    public static class DaodatFeasibility
    {
        private const double Tolerance = 0.000001;

        public static DaodatFeasibilityResult Analyze(int count, double targetSum, double volumeMin, double volumeMax)
        {
            var result = new DaodatFeasibilityResult();

            if (count <= 0)
            {
                result.Status = DaodatFeasibilityStatus.InvalidInput;
                result.Suggestion = "So luong tin hieu phai lon hon 0.";
                return result;
            }

            if (targetSum <= 0 || volumeMin <= 0 || volumeMax <= 0 || volumeMin > volumeMax)
            {
                result.Status = DaodatFeasibilityStatus.InvalidInput;
                result.AverageVolume = count > 0 ? targetSum / count : 0;
                result.Suggestion = "Du lieu V, V min hoac V max khong hop le.";
                return result;
            }

            result.AverageVolume = targetSum / count;
            if (result.AverageVolume < volumeMin - Tolerance)
            {
                result.Status = DaodatFeasibilityStatus.BelowMinimum;
                result.Suggestion = "Gia tri trung binh dang thap hon V min. Hay giam can duoi H/d/r hoac tang so luong N.";
                return result;
            }

            if (result.AverageVolume > volumeMax + Tolerance)
            {
                result.Status = DaodatFeasibilityStatus.AboveMaximum;
                result.Suggestion = "Gia tri trung binh dang cao hon V max. Hay tang can tren H/d/r hoac giam so luong N.";
                return result;
            }

            result.Status = DaodatFeasibilityStatus.Ok;
            result.Suggestion = "Gia tri nam trong khoang V co the tao.";
            return result;
        }

        public static string BuildFailureMessage(
            string label,
            double targetSum,
            int count,
            double volumeMin,
            double volumeMax,
            int maxAttempts)
        {
            DaodatFeasibilityResult feasibility = Analyze(count, targetSum, volumeMin, volumeMax);
            string suggestion = feasibility.CanGenerate
                ? "Gia tri nam trong khoang V nhung chua tim duoc bo so phu hop. Hay tang so lan thu toi da, hoac noi rong can H/d/r mot chut."
                : feasibility.Suggestion;

            return $"Khong tao duoc ket qua cho {label} sau {maxAttempts} lan thu.\n" +
                   $"Tong can tao: {targetSum:0.###}; N: {count}; trung binh moi dong: {feasibility.AverageVolume:0.###}.\n" +
                   $"V co the tao: {volumeMin:0.###} den {volumeMax:0.###}.\n" +
                   suggestion;
        }
    }
}
