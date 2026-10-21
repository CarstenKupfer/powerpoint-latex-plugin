using System;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointLatexPlugin
{
    internal class LatexFieldDrawer : IDisposable
    {
        private const int WH_MOUSE = 7;

        private const int WM_MOUSEMOVE = 0x0200;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONUP = 0x0202;

        private readonly PowerPoint.Application application;

        private IntPtr hookId = IntPtr.Zero;
        private HookProc hookProc;

        private bool isDragging = false;
        private POINT startPoint;

        private PowerPoint.Shape previewShape;

        public LatexFieldDrawer(PowerPoint.Application application)
        {
            this.application = application;
            hookProc = MouseHookCallback;
        }

        public void Start()
        {
            Stop();

            hookId = SetWindowsHookEx(
                WH_MOUSE,
                hookProc,
                IntPtr.Zero,
                GetCurrentThreadId()
            );
        }

        public void Stop()
        {
            if (hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(hookId);
                hookId = IntPtr.Zero;
            }

            isDragging = false;
        }

        private IntPtr MouseHookCallback(
            int nCode,
            IntPtr wParam,
            IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int message = wParam.ToInt32();

                MOUSEHOOKSTRUCT mouseData =
                    Marshal.PtrToStructure<MOUSEHOOKSTRUCT>(lParam);

                if (message == WM_LBUTTONDOWN)
                {
                    if (IsPointInsideSlide(mouseData.pt))
                    {
                        startPoint = mouseData.pt;
                        isDragging = true;

                        CreatePreview(startPoint);

                        return new IntPtr(1);
                    }

                    Stop();
                }

                if (message == WM_MOUSEMOVE && isDragging)
                {
                    UpdatePreview(mouseData.pt);

                    return new IntPtr(1);
                }

                if (message == WM_LBUTTONUP && isDragging)
                {
                    POINT endPoint = mouseData.pt;

                    isDragging = false;

                    DeletePreview();

                    Stop();

                    CreateLatexField(startPoint, endPoint);

                    return new IntPtr(1);
                }
            }

            return CallNextHookEx(
                hookId,
                nCode,
                wParam,
                lParam
            );
        }

        private void CreatePreview(POINT point)
        {
            PowerPoint.Slide slide =
                application.ActiveWindow.View.Slide;

            float x = ScreenToSlideX(point.X);
            float y = ScreenToSlideY(point.Y);

            previewShape = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle,
                x,
                y,
                1,
                1
            );

            // Keine Füllung
            previewShape.Fill.Visible =
                Office.MsoTriState.msoFalse;

            // Sichtbare Umrandung
            previewShape.Line.Visible =
                Office.MsoTriState.msoTrue;

            // Gestrichelte Linie
            previewShape.Line.DashStyle =
                Office.MsoLineDashStyle.msoLineDash;
        }
        private void UpdatePreview(POINT currentPoint)
        {
            if (previewShape == null)
                return;

            float startX =
                ScreenToSlideX(startPoint.X);

            float startY =
                ScreenToSlideY(startPoint.Y);

            float currentX =
                ScreenToSlideX(currentPoint.X);

            float currentY =
                ScreenToSlideY(currentPoint.Y);

            float left =
                Math.Min(startX, currentX);

            float top =
                Math.Min(startY, currentY);

            float width =
                Math.Abs(currentX - startX);

            float height =
                Math.Abs(currentY - startY);

            previewShape.Left = left;
            previewShape.Top = top;

            previewShape.Width =
                Math.Max(width, 1);

            previewShape.Height =
                Math.Max(height, 1);
        }

        private void DeletePreview()
        {
            if (previewShape != null)
            {
                previewShape.Delete();
                previewShape = null;
            }
        }


        private bool IsPointInsideSlide(POINT point)
        {
            try
            {
                PowerPoint.DocumentWindow window =
                    application.ActiveWindow;

                PowerPoint.Presentation presentation =
                    application.ActivePresentation;

                float slideWidth =
                    presentation.PageSetup.SlideWidth;

                float slideHeight =
                    presentation.PageSetup.SlideHeight;

                float left =
                    window.PointsToScreenPixelsX(0);

                float top =
                    window.PointsToScreenPixelsY(0);

                float right =
                    window.PointsToScreenPixelsX(slideWidth);

                float bottom =
                    window.PointsToScreenPixelsY(slideHeight);

                return point.X >= left &&
                       point.X <= right &&
                       point.Y >= top &&
                       point.Y <= bottom;
            }
            catch
            {
                return false;
            }
        }

        private void CreateLatexField(
            POINT start,
            POINT end)
        {
            PowerPoint.DocumentWindow window =
                application.ActiveWindow;

            PowerPoint.Slide slide =
                window.View.Slide;

            float startX = ScreenToSlideX(start.X);
            float startY = ScreenToSlideY(start.Y);

            float endX = ScreenToSlideX(end.X);
            float endY = ScreenToSlideY(end.Y);

            float left = Math.Min(startX, endX);
            float top = Math.Min(startY, endY);

            float width = Math.Abs(endX - startX);
            float height = Math.Abs(endY - startY);

            if (width < 20)
                width = 200;

            if (height < 20)
                height = 40;

            PowerPoint.Shape textBox =
                slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    left,
                    top,
                    width,
                    height
                );

            textBox.Tags.Add("LatexField", "true");

            textBox.TextFrame.TextRange.Select();

            System.Windows.Forms.MessageBox.Show(
                "LaTeX Field successfully detected!\nTag: "
                + textBox.Tags["LatexField"]
            );
        }

        private float ScreenToSlideX(int screenX)
        {
            PowerPoint.DocumentWindow window =
                application.ActiveWindow;

            float zero =
                window.PointsToScreenPixelsX(0);

            float hundred =
                window.PointsToScreenPixelsX(100);

            float pixelsPerPoint =
                (hundred - zero) / 100f;

            return (screenX - zero) / pixelsPerPoint;
        }

        private float ScreenToSlideY(int screenY)
        {
            PowerPoint.DocumentWindow window =
                application.ActiveWindow;

            float zero =
                window.PointsToScreenPixelsY(0);

            float hundred =
                window.PointsToScreenPixelsY(100);

            float pixelsPerPoint =
                (hundred - zero) / 100f;

            return (screenY - zero) / pixelsPerPoint;
        }

        public void Dispose()
        {
            Stop();
        }

        private delegate IntPtr HookProc(
            int nCode,
            IntPtr wParam,
            IntPtr lParam
        );

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEHOOKSTRUCT
        {
            public POINT pt;
            public IntPtr hwnd;
            public uint wHitTestCode;
            public UIntPtr dwExtraInfo;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(
            int idHook,
            HookProc lpfn,
            IntPtr hMod,
            uint dwThreadId
        );

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(
            IntPtr hhk
        );

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(
            IntPtr hhk,
            int nCode,
            IntPtr wParam,
            IntPtr lParam
        );

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
    }
}