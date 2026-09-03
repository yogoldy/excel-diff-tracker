using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace ExcelDiffTracker.App.Services;

public enum TrayState
{
    Idle,
    Processing,
    Paused,
    Warning
}

public static class AppIconFactory
{
    public static Icon Create(TrayState state)
    {
        using var bitmap = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(Color.Transparent);

        var primary = state switch
        {
            TrayState.Processing => Color.FromArgb(45, 132, 170),
            TrayState.Paused => Color.FromArgb(115, 123, 119),
            TrayState.Warning => Color.FromArgb(184, 100, 31),
            _ => Color.FromArgb(40, 122, 104)
        };
        using var fill = new SolidBrush(primary);
        using var white = new Pen(Color.White, 1.7f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round };

        graphics.FillRoundedRectangle(fill, new RectangleF(2, 2, 28, 28), 7);
        graphics.DrawRoundedRectangle(white, new RectangleF(5, 6, 11, 20), 1.5f);
        graphics.DrawLine(white, 5, 12, 16, 12);
        graphics.DrawLine(white, 5, 19, 16, 19);
        graphics.DrawLine(white, 10.5f, 6, 10.5f, 26);
        graphics.DrawLine(white, 13, 16, 17, 16);
        graphics.DrawLine(white, 17, 16, 23, 10);
        graphics.DrawLine(white, 17, 16, 25, 16);
        graphics.DrawLine(white, 17, 16, 23, 22);
        graphics.FillEllipse(Brushes.White, 22, 8.5f, 3, 3);
        graphics.FillEllipse(Brushes.White, 24, 14.5f, 3, 3);
        graphics.FillEllipse(Brushes.White, 22, 20.5f, 3, 3);
        switch (state)
        {
            case TrayState.Warning:
                graphics.FillEllipse(Brushes.White, 25, 25, 5, 5);
                break;
            case TrayState.Paused:
                graphics.DrawLine(white, 26, 25, 26, 29);
                graphics.DrawLine(white, 29, 25, 29, 29);
                break;
            case TrayState.Processing:
                graphics.DrawArc(white, 24, 24, 6, 6, 25, 275);
                break;
        }

        var handle = bitmap.GetHicon();
        try
        {
            using var temporary = Icon.FromHandle(handle);
            return (Icon)temporary.Clone();
        }
        finally
        {
            _ = DestroyIcon(handle);
        }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(nint handle);

    private static void FillRoundedRectangle(this Graphics graphics, Brush brush, RectangleF rectangle, float radius)
    {
        using var path = RoundedPath(rectangle, radius);
        graphics.FillPath(brush, path);
    }
    private static void DrawRoundedRectangle(this Graphics graphics, Pen pen, RectangleF rectangle, float radius)
    {
        using var path = RoundedPath(rectangle, radius);
        graphics.DrawPath(pen, path);
    }
    private static GraphicsPath RoundedPath(RectangleF rectangle, float radius)
    {
        var path = new GraphicsPath();
        var diameter = radius * 2;
        path.AddArc(rectangle.Left, rectangle.Top, diameter, diameter, 180, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Top, diameter, diameter, 270, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(rectangle.Left, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }
}
