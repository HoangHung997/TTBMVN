using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelAddIn1
{
    public static partial class RandomDaodat
    {
        private static List<(double d1, double r1, double d2, double r2, double H)> GenerateResultsForTarget(
            int count,
            double targetSum,
            DaodatRandomSettings settings,
            string label)
        {
            EnsureTargetFeasible(label, count, targetSum, settings);

            List<double> randomParts = GenerateRandomParts(count, settings.VolumeMin, settings.VolumeMax, targetSum, settings.MaxAttempts);
            if (randomParts == null)
                throw new InvalidOperationException(BuildFailureMessage(label, targetSum, count, settings));

            int success = 0;
            var cts = new CancellationTokenSource();
            var options = new ParallelOptions
            {
                CancellationToken = cts.Token,
                MaxDegreeOfParallelism = GetOptimalParallelism(settings)
            };

            ConcurrentBag<(double d1, double r1, double d2, double r2, double H)> results =
                new ConcurrentBag<(double, double, double, double, double)>();

            try
            {
                Parallel.For(0, settings.MaxAttempts, options, (attempt, state) =>
                {
                    if (Volatile.Read(ref success) == 1)
                    {
                        state.Stop();
                        return;
                    }

                    var trialResults = new List<(double d1, double r1, double d2, double r2, double H)>();
                    double sumV = 0;

                    for (int i = 0; i < count; i++)
                    {
                        if (GenerateParametersForVolume(randomParts[i], out double d1, out double r1, out double d2, out double r2, out double H, settings))
                        {
                            double v = Math.Round(DaodatRandomSettings.CalculateVolume(d1, r1, d2, r2, H), 2);
                            trialResults.Add((d1, r1, d2, r2, H));
                            sumV += v;
                        }
                        else
                        {
                            break;
                        }
                    }

                    if (trialResults.Count == count && Math.Abs(sumV - targetSum) <= 0.001)
                    {
                        if (Interlocked.CompareExchange(ref success, 1, 0) == 0)
                        {
                            results = new ConcurrentBag<(double, double, double, double, double)>(trialResults);
                            cts.Cancel();
                            state.Stop();
                        }
                    }
                });
            }
            catch (OperationCanceledException)
            {
            }

            if (Volatile.Read(ref success) == 0 || results.Count == 0)
                throw new InvalidOperationException(BuildFailureMessage(label, targetSum, count, settings));

            return results.ToList();
        }
    }
}
