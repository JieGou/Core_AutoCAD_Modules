#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using ModPlusAPI;
#pragma warning disable 1591

namespace ModPlus.App
{
    public class MpProductsDrawableOverrule : DrawableOverrule
    {
        public static MpProductsDrawableOverrule MpProductsDrawableOverruleInstance;
        public static MpProductsDrawableOverrule Instance()
        {
            return MpProductsDrawableOverruleInstance ?? (MpProductsDrawableOverruleInstance = new MpProductsDrawableOverrule());
        }

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
                            double height = (double) AcApp.GetSystemVariable("VIEWSIZE");
                            var scale = height / 1500;
                            var offset = 20;
                            var plane = ent.GetPlane();

                            Matrix3d ucs = AcApp.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem;
                            var extents = ent.GeometricExtents;
                            //var pt = extents.MaxPoint;
                            var pt = extents.MaxPoint.TransformBy(ucs.Inverse());
                            Point3dCollection points = new Point3dCollection();
                            points.Add(new Point3d(pt.X + (offset + 00) * scale, pt.Y + (offset + 00) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 00) * scale, pt.Y + (offset + 10) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 10) * scale, pt.Y + (offset + 10) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 10) * scale, pt.Y + (offset + 30) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 00) * scale, pt.Y + (offset + 30) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 00) * scale, pt.Y + (offset + 40) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 30) * scale, pt.Y + (offset + 40) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 30) * scale, pt.Y + (offset + 30) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 20) * scale, pt.Y + (offset + 30) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 20) * scale, pt.Y + (offset + 10) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 30) * scale, pt.Y + (offset + 10) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 30) * scale, pt.Y + (offset + 00) * scale, plane.PointOnPlane.Z));
                            points.Add(new Point3d(pt.X + (offset + 00) * scale, pt.Y + (offset + 00) * scale, plane.PointOnPlane.Z));

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

    public class MpProductIconFunctions
    {
        /// <summary>Включить идентификационные иконки для примитивов, имеющих расширенные данные продуктов ModPlus</summary>
        [CommandMethod("mpShowProductIcons")]
        public static void ShowIcon()
        {
            if (MpProductsDrawableOverrule.MpProductsDrawableOverruleInstance == null)
            {
                UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "mpProductInsert", "ShowIcon",
                    true.ToString(), true);
                Overrule.AddOverrule(RXObject.GetClass(typeof(Entity)), MpProductsDrawableOverrule.Instance(), true);
                Overrule.Overruling = true;
                AcApp.DocumentManager.MdiActiveDocument.Editor.Regen();
            }
        }
        /// <summary>Отключить идентификационные иконки для примитивов, имеющих расширенные данные продуктов ModPlus</summary>
        [CommandMethod("mpHideProductIcons")]
        public void HideIcon()
        {
            if (MpProductsDrawableOverrule.MpProductsDrawableOverruleInstance != null)
            {
                UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "mpProductInsert", "ShowIcon",
                    false.ToString(), true);
                Overrule.RemoveOverrule(RXObject.GetClass(typeof(Entity)), MpProductsDrawableOverrule.Instance());
                MpProductsDrawableOverrule.MpProductsDrawableOverruleInstance = null;
                AcApp.DocumentManager.MdiActiveDocument.Editor.Regen();
            }
        }
    }
}
