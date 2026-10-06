using System;
using System.Globalization;
using System.Windows.Forms;

namespace ExcelAddIn1.Winform
{
    public class SmartNumericUpDown : NumericUpDown
    {
        private bool updatingScale;

        public SmartNumericUpDown()
        {
            Minimum = 0;
            Maximum = 1000000000;
            MaximumDecimalPlaces = 4;
            DecimalPlaces = 0;
            Increment = 1;
        }

        public int MaximumDecimalPlaces { get; set; }

        public void SetSmartValue(decimal value, string preferredText = null)
        {
            int decimalPlaces = CountDecimalPlaces(preferredText ?? FormatValue(value));
            SetScale(decimalPlaces);
            Value = ClampValue(value);
            UpdateSmartIncrement();
        }

        protected override void OnTextChanged(EventArgs e)
        {
            base.OnTextChanged(e);
            UpdateSmartIncrement();
        }

        protected override void OnValueChanged(EventArgs e)
        {
            base.OnValueChanged(e);
            UpdateSmartIncrement();
        }

        private void UpdateSmartIncrement()
        {
            if (updatingScale)
                return;

            int decimalPlaces = CountDecimalPlaces(Text);
            SetScale(decimalPlaces);
            Increment = decimalPlaces <= 0 ? 1 : Pow10Negative(decimalPlaces);
        }

        private void SetScale(int decimalPlaces)
        {
            decimalPlaces = Math.Max(0, Math.Min(MaximumDecimalPlaces, decimalPlaces));
            if (DecimalPlaces == decimalPlaces)
                return;

            try
            {
                updatingScale = true;
                DecimalPlaces = decimalPlaces;
            }
            finally
            {
                updatingScale = false;
            }
        }

        private decimal ClampValue(decimal value)
        {
            if (value < Minimum)
                return Minimum;
            if (value > Maximum)
                return Maximum;
            return value;
        }

        private decimal Pow10Negative(int decimalPlaces)
        {
            decimal result = 1;
            for (int i = 0; i < decimalPlaces; i++)
                result /= 10;
            return result;
        }

        private int CountDecimalPlaces(string text)
        {
            if (MaximumDecimalPlaces <= 0 || string.IsNullOrWhiteSpace(text))
                return 0;

            string decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            int index = text.LastIndexOf(decimalSeparator, StringComparison.Ordinal);
            if (index < 0 || index >= text.Length - 1)
                return 0;

            int count = 0;
            for (int i = index + 1; i < text.Length; i++)
            {
                if (!char.IsDigit(text[i]))
                    break;
                count++;
            }

            return Math.Min(count, MaximumDecimalPlaces);
        }

        private string FormatValue(decimal value)
        {
            return value.ToString("0.####", CultureInfo.CurrentCulture);
        }
    }
}
