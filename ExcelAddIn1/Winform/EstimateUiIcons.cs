using System;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace ExcelAddIn1.Winform
{
    internal enum EstimateUiIconKind
    {
        Clipboard,
        Check,
        Warning,
        Error,
        Database,
        Document,
        Folder,
        Settings,
        Refresh,
        Link
    }

    internal static class EstimateUiIcons
    {
        internal static Bitmap Create(
            EstimateUiIconKind kind,
            int size,
            Color color)
        {
            int actual = Math.Max(16, size);
            var bitmap = new Bitmap(actual, actual);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (var pen = new Pen(color, Math.Max(1.6f, actual / 12f)))
            using (var brush = new SolidBrush(color))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.Clear(Color.Transparent);
                float s = actual;
                switch (kind)
                {
                    case EstimateUiIconKind.Check:
                        graphics.DrawEllipse(pen, s * .12f, s * .12f, s * .76f, s * .76f);
                        graphics.DrawLines(pen, new[]
                        {
                            new PointF(s * .29f, s * .51f),
                            new PointF(s * .44f, s * .65f),
                            new PointF(s * .72f, s * .35f)
                        });
                        break;
                    case EstimateUiIconKind.Warning:
                        graphics.DrawPolygon(pen, new[]
                        {
                            new PointF(s * .50f, s * .10f),
                            new PointF(s * .90f, s * .84f),
                            new PointF(s * .10f, s * .84f)
                        });
                        graphics.DrawLine(pen, s * .50f, s * .34f, s * .50f, s * .59f);
                        graphics.FillEllipse(brush, s * .46f, s * .69f, s * .08f, s * .08f);
                        break;
                    case EstimateUiIconKind.Error:
                        graphics.DrawEllipse(pen, s * .12f, s * .12f, s * .76f, s * .76f);
                        graphics.DrawLine(pen, s * .34f, s * .34f, s * .66f, s * .66f);
                        graphics.DrawLine(pen, s * .66f, s * .34f, s * .34f, s * .66f);
                        break;
                    case EstimateUiIconKind.Database:
                        graphics.DrawEllipse(pen, s * .18f, s * .15f, s * .64f, s * .22f);
                        graphics.DrawArc(pen, s * .18f, s * .32f, s * .64f, s * .22f, 0, 180);
                        graphics.DrawArc(pen, s * .18f, s * .50f, s * .64f, s * .22f, 0, 180);
                        graphics.DrawLine(pen, s * .18f, s * .26f, s * .18f, s * .62f);
                        graphics.DrawLine(pen, s * .82f, s * .26f, s * .82f, s * .62f);
                        graphics.DrawArc(pen, s * .18f, s * .52f, s * .64f, s * .26f, 0, 180);
                        break;
                    case EstimateUiIconKind.Document:
                        graphics.DrawRectangle(pen, s * .22f, s * .12f, s * .55f, s * .74f);
                        graphics.DrawLine(pen, s * .31f, s * .35f, s * .67f, s * .35f);
                        graphics.DrawLine(pen, s * .31f, s * .49f, s * .67f, s * .49f);
                        graphics.DrawLine(pen, s * .31f, s * .63f, s * .61f, s * .63f);
                        break;
                    case EstimateUiIconKind.Folder:
                        graphics.DrawPolygon(pen, new[]
                        {
                            new PointF(s * .12f, s * .29f),
                            new PointF(s * .40f, s * .29f),
                            new PointF(s * .49f, s * .39f),
                            new PointF(s * .88f, s * .39f),
                            new PointF(s * .82f, s * .78f),
                            new PointF(s * .14f, s * .78f)
                        });
                        break;
                    case EstimateUiIconKind.Settings:
                        graphics.DrawEllipse(pen, s * .30f, s * .30f, s * .40f, s * .40f);
                        graphics.DrawEllipse(pen, s * .43f, s * .43f, s * .14f, s * .14f);
                        for (int i = 0; i < 8; i++)
                        {
                            double angle = Math.PI * 2 * i / 8;
                            float x1 = s * .5f + (float)Math.Cos(angle) * s * .27f;
                            float y1 = s * .5f + (float)Math.Sin(angle) * s * .27f;
                            float x2 = s * .5f + (float)Math.Cos(angle) * s * .39f;
                            float y2 = s * .5f + (float)Math.Sin(angle) * s * .39f;
                            graphics.DrawLine(pen, x1, y1, x2, y2);
                        }
                        break;
                    case EstimateUiIconKind.Refresh:
                        graphics.DrawArc(pen, s * .18f, s * .18f, s * .64f, s * .64f, 35, 245);
                        graphics.DrawLines(pen, new[]
                        {
                            new PointF(s * .78f, s * .22f),
                            new PointF(s * .80f, s * .42f),
                            new PointF(s * .62f, s * .34f)
                        });
                        break;
                    case EstimateUiIconKind.Link:
                        graphics.DrawArc(pen, s * .12f, s * .31f, s * .45f, s * .38f, 120, 230);
                        graphics.DrawArc(pen, s * .43f, s * .31f, s * .45f, s * .38f, -60, 230);
                        graphics.DrawLine(pen, s * .38f, s * .50f, s * .62f, s * .50f);
                        break;
                    case EstimateUiIconKind.Clipboard:
                    default:
                        graphics.DrawRectangle(pen, s * .20f, s * .22f, s * .60f, s * .66f);
                        graphics.DrawRectangle(pen, s * .34f, s * .12f, s * .32f, s * .18f);
                        graphics.DrawLine(pen, s * .32f, s * .45f, s * .68f, s * .45f);
                        graphics.DrawLine(pen, s * .32f, s * .60f, s * .68f, s * .60f);
                        break;
                }
            }
            return bitmap;
        }
    }
}
