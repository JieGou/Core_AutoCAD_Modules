using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using mpSettings;

namespace ModPlus.Windows
{
    /// <summary>
    /// Логика взаимодействия для NewFunctionsShow.xaml
    /// </summary>
    public partial class NewFunctionsShow
    {
        public NewFunctionsShow()
        {
            InitializeComponent();
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
                );
        }

        private void BtFunctionVideo_OnClick(object sender, RoutedEventArgs e)
        {
            var bt = sender as Button;
            if (!string.IsNullOrEmpty(bt?.Tag.ToString()))
            {
                OpenVideo win = new OpenVideo(bt.Tag.ToString());
                win.ShowDialog();
            }
        }
    }
    /// <summary>
    /// Специальный декоратор, служащий для "добавления четкости" изображению
    /// Элемент Image нужно "обернуть" в этот декоратор
    /// </summary>
    public class DevicePixels : Decorator
    {
        public DevicePixels()
        {
            SnapsToDevicePixels = true;
            this.HorizontalAlignment = HorizontalAlignment.Left;
            this.VerticalAlignment = VerticalAlignment.Top;
            LayoutUpdated += OnLayoutUpdated;
        }

        private void OnLayoutUpdated(object sender, EventArgs e)
        {
            Point pixelOffset = GetPixelOffset();
            if (!AreClose(pixelOffset, _pixelOffset))
            {
                _pixelOffset = pixelOffset;
                InvalidateVisual();
            }
        }

        protected override void OnRender(DrawingContext dc)
        {
        }

        protected override Size ArrangeOverride(Size arrangeSize)
        {
            UIElement child = Child;
            if (child != null)
            {
                Rect finalRect = HelperDeflateRect(new Rect(arrangeSize), _pixelOffset);
                child.Arrange(finalRect);
            }
            return arrangeSize;
        }

        protected override Size MeasureOverride(Size constraint)
        {
            UIElement child = this.Child;
            if (child != null)
            {
                Size availableSize = new Size(Math.Max(0.0, constraint.Width - _pixelOffset.X), Math.Max(0.0, constraint.Height - _pixelOffset.Y));
                child.Measure(availableSize);
                Size desiredSize = child.DesiredSize;
                return new Size(desiredSize.Width + _pixelOffset.X, desiredSize.Height + _pixelOffset.Y);
            }
            return new Size();
        }

        private static Rect HelperDeflateRect(Rect rt, Point offset)
        {
            return new Rect(
            rt.Left + offset.X,
            rt.Top + offset.Y,
            Math.Max(0.0, rt.Width - offset.X),
            Math.Max(0.0, rt.Height - offset.Y));
        }

        // Gets the matrix that will convert a point from "above" the
        // coordinate space of a visual into the the coordinate space
        // "below" the visual.
        private static Matrix GetVisualTransform(Visual v)
        {
            if (v != null)
            {
                Matrix m = Matrix.Identity;

                Transform transform = VisualTreeHelper.GetTransform(v);
                if (transform != null)
                {
                    Matrix cm = transform.Value;
                    m = Matrix.Multiply(m, cm);
                }

                Vector offset = VisualTreeHelper.GetOffset(v);
                m.Translate(offset.X, offset.Y);

                return m;
            }

            return Matrix.Identity;
        }

        private static Point TryApplyVisualTransform(Point point, Visual v, bool inverse, bool throwOnError, out bool success)
        {
            success = true;
            if (v != null)
            {
                Matrix visualTransform = GetVisualTransform(v);
                if (inverse)
                {
                    if (!throwOnError && !visualTransform.HasInverse)
                    {
                        success = false;
                        return new Point(0, 0);
                    }
                    visualTransform.Invert();
                }
                point = visualTransform.Transform(point);
            }
            return point;
        }

        private static Point ApplyVisualTransform(Point point, Visual v, bool inverse)
        {
            bool success;
            return TryApplyVisualTransform(point, v, inverse, true, out success);
        }


        private Point GetPixelOffset()
        {
            Point pixelOffset = new Point();

            PresentationSource ps = PresentationSource.FromVisual(this);
            if (ps != null)
            {
                Visual rootVisual = ps.RootVisual;

                // Transform (0,0) from this element up to pixels.
                pixelOffset = TransformToAncestor(rootVisual).Transform(pixelOffset);
                pixelOffset = ApplyVisualTransform(pixelOffset, rootVisual, false);
                pixelOffset = ps.CompositionTarget.TransformToDevice.Transform(pixelOffset);

                // Round the origin to the nearest whole pixel.
                pixelOffset.X = Math.Ceiling(pixelOffset.X);
                pixelOffset.Y = Math.Ceiling(pixelOffset.Y);

                // Transform the whole-pixel back to this element.
                pixelOffset = ps.CompositionTarget.TransformFromDevice.Transform(pixelOffset);
                pixelOffset = ApplyVisualTransform(pixelOffset, rootVisual, true);
                pixelOffset = rootVisual.TransformToDescendant(this).Transform(pixelOffset);
            }
            return pixelOffset;
        }
        private static bool AreClose(Point point1, Point point2)
        {
            return AreClose(point1.X, point2.X) && AreClose(point1.Y, point2.Y);
        }

        private static bool AreClose(double value1, double value2)
        {
            if (value1 == value2)
            {
                return true;
            }
            double delta = value1 - value2;
            return ((delta < 1.53E-06) && (delta > -1.53E-06));
        }
        private Point _pixelOffset;
    }
}
