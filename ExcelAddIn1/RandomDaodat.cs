
using ExcelAddIn1.Funtion;
using ExcelAddIn1.Core;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class DaodatRandomSettings
    {
        public double D1Min { get; set; }
        public double D1Max { get; set; }
        public double R1Min { get; set; }
        public double R1Max { get; set; }
        public double D2Min { get; set; }
        public double D2Max { get; set; }
        public double R2Min { get; set; }
        public double R2Max { get; set; }
        public double HMin { get; set; }
        public double HMax { get; set; }
        public int MaxAttempts { get; set; } = 10000;
        public int MaxParallelWorkers { get; set; } = 0;

        public double VolumeMin => CalculateVolume(D1Min, R1Min, D2Min, R2Min, HMin);
        public double VolumeMax => CalculateVolume(D1Max, R1Max, D2Max, R2Max, HMax);

        public DaodatRandomSettings Clone()
        {
            return new DaodatRandomSettings
            {
                D1Min = D1Min,
                D1Max = D1Max,
                R1Min = R1Min,
                R1Max = R1Max,
                D2Min = D2Min,
                D2Max = D2Max,
                R2Min = R2Min,
                R2Max = R2Max,
                HMin = HMin,
                HMax = HMax,
                MaxAttempts = MaxAttempts,
                MaxParallelWorkers = MaxParallelWorkers
            };
        }

        public void Validate()
        {
            ValidateRange("d1", D1Min, D1Max);
            ValidateRange("r1", R1Min, R1Max);
            ValidateRange("d2", D2Min, D2Max);
            ValidateRange("r2", R2Min, R2Max);
            ValidateRange("H", HMin, HMax);

            if (MaxAttempts < 1)
                throw new ArgumentException("So lan thu toi da phai lon hon 0.");

            if (MaxParallelWorkers < 0)
                throw new ArgumentException("So luong toi da khong duoc nho hon 0.");
        }

        public static DaodatRandomSettings CreateDefault(bool is5m)
        {
            return is5m
                ? new DaodatRandomSettings
                {
                    D1Min = 2.7,
                    D1Max = 3.3,
                    R1Min = 2.2,
                    R1Max = 2.8,
                    D2Min = 1.3,
                    D2Max = 1.7,
                    R2Min = 0.8,
                    R2Max = 1.2,
                    HMin = 0.1,
                    HMax = 10,
                    MaxAttempts = 10000,
                    MaxParallelWorkers = 0
                }
                : new DaodatRandomSettings
                {
                    D1Min = 1.7,
                    D1Max = 2.3,
                    R1Min = 1.2,
                    R1Max = 1.8,
                    D2Min = 1.1,
                    D2Max = 1.5,
                    R2Min = 0.6,
                    R2Max = 1.0,
                    HMin = 0.1,
                    HMax = 10,
                    MaxAttempts = 10000,
                    MaxParallelWorkers = 0
                };
        }

        public static double CalculateVolume(double d1, double r1, double d2, double r2, double h)
        {
            return DaodatMath.CalculateVolume(d1, r1, d2, r2, h);
        }

        private static void ValidateRange(string name, double min, double max)
        {
            if (min <= 0 || max <= 0)
                throw new ArgumentException($"Can {name} phai lon hon 0.");

            if (min > max)
                throw new ArgumentException($"Can duoi {name} khong duoc lon hon can tren.");
        }
    }

    public class DaodatRunOptions
    {
        public bool Is5m { get; set; }
        public bool InsertRows { get; set; }
        public bool CreateNewSheet { get; set; }
        public bool LinkBack { get; set; }
        public bool UseOutputStartCell { get; set; }
        public string OutputStartAddress { get; set; }
        public bool HasTableData { get; set; }
        public bool ProjectIncludes5m { get; set; }
        public bool UseDefaultOutputRow { get; set; } = true;
        public string TableDataAddress { get; set; }
        public int OutputStartRow { get; set; }
        public string NewSheetName { get; set; }
        public string ExistingSheetName { get; set; }
        public Dictionary<string, string> SourceDataColumns { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public List<DaodatColumnMapping> ColumnMappings { get; set; } = new List<DaodatColumnMapping>();

        public DaodatRunOptions Clone()
        {
            return new DaodatRunOptions
            {
                Is5m = Is5m,
                InsertRows = InsertRows,
                CreateNewSheet = CreateNewSheet,
                LinkBack = LinkBack,
                UseOutputStartCell = UseOutputStartCell,
                OutputStartAddress = OutputStartAddress,
                HasTableData = HasTableData,
                ProjectIncludes5m = ProjectIncludes5m,
                UseDefaultOutputRow = UseDefaultOutputRow,
                TableDataAddress = TableDataAddress,
                OutputStartRow = OutputStartRow,
                NewSheetName = NewSheetName,
                ExistingSheetName = ExistingSheetName,
                SourceDataColumns = new Dictionary<string, string>(SourceDataColumns ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase),
                ColumnMappings = (ColumnMappings ?? new List<DaodatColumnMapping>()).Select(x => x.Clone()).ToList()
            };
        }
    }

    public class DaodatColumnMapping
    {
        public string Key { get; set; }
        public string SourceColumn { get; set; }
        public string OutputColumn { get; set; }
        public string FormatLocal { get; set; }

        public DaodatColumnMapping Clone()
        {
            return new DaodatColumnMapping
            {
                Key = Key,
                SourceColumn = SourceColumn,
                OutputColumn = OutputColumn,
                FormatLocal = FormatLocal
            };
        }
    }

    public static partial class RandomDaodat
    {
        private static int randomSeed = Environment.TickCount;
        private static readonly ThreadLocal<Random> threadRandom =
            new ThreadLocal<Random>(() => new Random(Interlocked.Increment(ref randomSeed)));
        private static DaodatRandomSettings settings3m = DaodatRandomSettings.CreateDefault(false);
        private static DaodatRandomSettings settings5m = DaodatRandomSettings.CreateDefault(true);
        private static DaodatRunOptions runOptions3m = new DaodatRunOptions { Is5m = false };
        private static DaodatRunOptions runOptions5m = new DaodatRunOptions { Is5m = true };

        static RandomDaodat()
        {
            LoadPersistedSettings();
        }

        public static DaodatRandomSettings GetSettings(bool is5m)
        {
            return (is5m ? settings5m : settings3m).Clone();
        }

        public static void SaveSettings(bool is5m, DaodatRandomSettings settings)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            settings.Validate();

            if (is5m)
                settings5m = settings.Clone();
            else
                settings3m = settings.Clone();

            SavePersistedSettings();
        }

        public static DaodatRunOptions GetRunOptions(bool is5m)
        {
            return (is5m ? runOptions5m : runOptions3m).Clone();
        }

        public static void SaveRunOptions(DaodatRunOptions runOptions)
        {
            if (runOptions == null)
                throw new ArgumentNullException(nameof(runOptions));

            DaodatRunOptions savedOptions = runOptions.Clone();
            savedOptions.OutputStartAddress = string.IsNullOrWhiteSpace(savedOptions.OutputStartAddress)
                ? null
                : savedOptions.OutputStartAddress.Trim();
            savedOptions.UseOutputStartCell = savedOptions.UseOutputStartCell && !string.IsNullOrEmpty(savedOptions.OutputStartAddress);

            if (savedOptions.Is5m)
                runOptions5m = savedOptions;
            else
                runOptions3m = savedOptions;

            SavePersistedSettings();
        }
    }
}
