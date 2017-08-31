
using System;
using Autodesk.AutoCAD.Colors;
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
using ModPlusAPI.Windows;

namespace ModPlus.App
{
    public class MpProductsDrawableOverrule : DrawableOverrule
    {
        protected static MpProductsDrawableOverrule _mpProductsDrawableOverrule;
        public static MpProductsDrawableOverrule Instance()
        {
            return _mpProductsDrawableOverrule ?? (_mpProductsDrawableOverrule = new MpProductsDrawableOverrule());
        }

        public override bool WorldDraw(Drawable drawable, WorldDraw wd)
        {
            var ent = drawable as Entity;
            if (ent != null)
                if (ent.IsModPlusProduct())
                {
                    double height = (double)AcApp.GetSystemVariable("VIEWSIZE");
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
                    wd.SubEntityTraits.FillType = FillType.FillAlways;
                    wd.SubEntityTraits.Color = 150;
                    wd.Geometry.Polygon(points);
                    wd.SubEntityTraits.FillType = FillType.FillNever;
                    wd.SubEntityTraits.Color = backupColor;
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
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "mpProductInsert", "ShowIcon", true.ToString(), true);
            Overrule.AddOverrule(RXObject.GetClass(typeof(Entity)), MpProductsDrawableOverrule.Instance(), true);
            Overrule.Overruling = true;
            AcApp.DocumentManager.MdiActiveDocument.Editor.Regen();
        }
        /// <summary>Отключить идентификационные иконки для примитивов, имеющих расширенные данные продуктов ModPlus</summary>
        [CommandMethod("mpHideProductIcons")]
        public void HideIcon()
        {
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "mpProductInsert", "ShowIcon", false.ToString(), true);
            Overrule.RemoveOverrule(RXObject.GetClass(typeof(Entity)), MpProductsDrawableOverrule.Instance());
            AcApp.DocumentManager.MdiActiveDocument.Editor.Regen();
        }
    }
}
