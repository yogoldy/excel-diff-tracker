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
        using var white = new Pen(Color.White, 1.8f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
        using var outline = new Pen(Color.White, 1.5f);

        graphics.FillRoundedRectangle(fill, new RectangleF(2, 2, 28, 28), 7);
        var page = new RectangleF(7, 5.5f, 15, 20);
        graphics.DrawRoundedRectangle(outline, page, 2);
        graphics.DrawLine(white, 10, 11, 19, 11);
        graphics.DrawLine(white, 10, 15, 17, 15);
        graphics.DrawLine(white, 10, 19, 15, 19);
        graphics.DrawEllipse(outline, 17, 17, 10, 10);
        switch (state)
        {
            case TrayState.Warning:
                graphics.DrawLine(white, 22, 19, 22, 22.5f);
                graphics.FillEllipse(Brushes.White, 21.1f, 24, 1.8f, 1.8f);
                break;
            case TrayState.Paused:
                graphics.DrawLine(white, 20.5f, 19.5f, 20.5f, 24.5f);
                graphics.DrawLine(white, 23.5f, 19.5f, 23.5f, 24.5f);
                break;
            case TrayState.Processing:
                graphics.DrawArc(white, 19, 19, 6, 6, 25, 275);
                graphics.DrawLine(white, 24.5f, 19, 25.5f, 21);
                break;
            default:
                graphics.DrawLine(white, 19.5f, 22, 21.5f, 24);
                graphics.DrawLine(white, 21.5f, 24, 25, 19.8f);
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
