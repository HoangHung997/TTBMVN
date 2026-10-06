using System;

namespace ExcelAddIn1.Core
{
    public static class DaodatMath
    {
        public static double CalculateVolume(double d1, double r1, double d2, double r2, double h)
        {
            double a1 = d1 * r1;
            double a2 = d2 * r2;
            return ((a1 + a2 + Math.Sqrt(a1 * a2)) * h) / 3;
        }

        public static double CalculateHeightFromVolume(double volume, double d1, double r1, double d2, double r2)
        {
            double a1 = d1 * r1;
            double a2 = d2 * r2;
            double denominator = a1 + a2 + Math.Sqrt(a1 * a2);
            if (denominator <= 0)
                throw new ArgumentException("Khong tinh duoc H tu V vi thong so d/r khong hop le.");

            return 3 * volume / denominator;
        }
    }
}
