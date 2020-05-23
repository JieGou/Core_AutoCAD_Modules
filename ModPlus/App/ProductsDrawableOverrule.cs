#pragma warning disable 1591

namespace ModPlus.App
{
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.GraphicsInterface;
    using ModPlusAPI;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

    /// <inheritdoc />
    public class ProductsDrawableOverrule : DrawableOverrule
    {
        public static ProductsDrawableOverrule ProductsDrawableOverruleInstance;

        public static ProductsDrawableOverrule Instance()
        {
            return ProductsDrawableOverruleInstance ?? (ProductsDrawableOverruleInstance = new ProductsDrawableOverrule());
        }

        /// <inheritdoc />
        public override bool WorldDraw(Drawable drawable, WorldDraw wd)
        {
            if (!wd.Context.IsPlotGeneration) // Не в состоянии печати!
            {
                var ent = drawable as Entity;
                if (ent != null && !ent.IsAProxy && !ent.IsCancelling && !ent.IsDisposed && !ent.IsErased)
                {
                    try
                    {
                        if (ent.IsModPlusProduct())
                        {
                            double height = (double)AcApp.GetSystemVariable("VIEWSIZE");
                            var scale = height / 1500;
                            var offset = 20;
                            var plane = ent.GetPlane();

                            Matrix3d ucs = AcApp.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem;
                            var extents = ent.GeometricExtents;
                            var pt = extents.MaxPoint.TransformBy(ucs.Inverse());
                            Point3dCollection points = new Point3dCollection();
                            points.Add(new Point3d(pt.X + ((offset + 00) * scale), pt.Y + ((offset + 00) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 00) * scale), pt.Y + ((offset + 10) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 10) * scale), pt.Y + ((offset + 10) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 10) * scale), pt.Y + ((offset + 30) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 00) * scale), pt.Y + ((offset + 30) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 00) * scale), pt.Y + ((offset + 40) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 30) * scale), pt.Y + ((offset + 40) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 30) * scale), pt.Y + ((offset + 30) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 20) * scale), pt.Y + ((offset + 30) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 20) * scale), pt.Y + ((offset + 10) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 30) * scale), pt.Y + ((offset + 10) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 30) * scale), pt.Y + ((offset + 00) * scale), plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + ((offset + 00) * scale), pt.Y + ((offset + 00) * scale), plane.PointOnPlane.Z));

                            short backupColor = wd.SubEntityTraits.Color;
                            FillType backupFillType = wd.SubEntityTraits.FillType;
                            wd.SubEntityTraits.FillType = FillType.FillAlways;
                            wd.SubEntityTraits.Color = 150;
                            wd.Geometry.Polygon(points);
                            wd.SubEntityTraits.FillType = FillType.FillNever;
                            
                            // restore
                            wd.SubEntityTraits.Color = backupColor;
                            wd.SubEntityTraits.FillType = backupFillType;
                        }
                    }
                    catch (System.Exception exception)
                    {
                        // not showing. Only sending by AppMetrica
                        Statistic.SendException(exception);
                    }
                }
            }

            return base.WorldDraw(drawable, wd);
        }
    }
}
