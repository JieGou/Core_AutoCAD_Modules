using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;

namespace ModPlus.MPCOHelpers.Overrules
{
    public class MPCOGrips
    {
        /// <summary>
        /// Виды ручек для примитива
        /// </summary>
        public enum MPCOEntityGripType
        {
            /// <summary>
            /// Обычная точка
            /// </summary>
            Point = 1,
            /// <summary>
            /// Отображение плюса
            /// </summary>
            Plus,
            /// <summary>
            /// Отображение минуса
            /// </summary>
            Minus,
            /// <summary>
            /// Положение текста
            /// </summary>
            Text,
            /// <summary>
            /// Список (выпадающий список)
            /// </summary>
            List
        }
        public abstract class MPCOGripData : GripData
        {
            /// <summary>
            /// Тип ручки примитива
            /// </summary>
            public MPCOEntityGripType GripType { get; set; }

            public override bool ViewportDraw(ViewportDraw worldDraw, ObjectId entityId, DrawType type, Point3d? imageGripPoint,
                int gripSizeInPixels)
            {
                CoordinateSystem3d eCS = GetECS(entityId);
                Point2d numPixelsInUnitSquare = worldDraw.Viewport.GetNumPixelsInUnitSquare(GripPoint);
                double num = (double)gripSizeInPixels / numPixelsInUnitSquare.X;
                Point3dCollection point3dCollections = new Point3dCollection();
                point3dCollections.Add((GripPoint - (num * eCS.Xaxis)) - (num * eCS.Yaxis));
                point3dCollections.Add((GripPoint - (num * eCS.Xaxis)) + (num * eCS.Yaxis));
                point3dCollections.Add((GripPoint + (num * eCS.Xaxis)) + (num * eCS.Yaxis));
                point3dCollections.Add((GripPoint + (num * eCS.Xaxis)) - (num * eCS.Yaxis));
                worldDraw.SubEntityTraits.FillType = FillType.FillAlways;
                worldDraw.Geometry.Polygon(point3dCollections);
                worldDraw.SubEntityTraits.FillType = FillType.FillNever;
                worldDraw.SubEntityTraits.TrueColor = new EntityColor(0, 0, 0);
                worldDraw.Geometry.Polygon(point3dCollections);
                return true;
            }

            protected CoordinateSystem3d GetECS(ObjectId entityId)
            {
                CoordinateSystem3d coordinateSystem3D = new CoordinateSystem3d(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis);
                if (!entityId.IsNull)
                {
                    using (OpenCloseTransaction openCloseTransaction = new OpenCloseTransaction())
                    {
                        Entity obj = (Entity)openCloseTransaction.GetObject(entityId, OpenMode.ForRead);
                        if (obj != null)
                        {
                            //if (obj.IsPlanar || obj is MLeader)
                            //{
                            Plane plane = obj.GetPlane();
                            Plane plane1 = new Plane(plane.PointOnPlane, plane.Normal);
                            coordinateSystem3D = plane1.GetCoordinateSystem();
                            //}
                            //if (obj is Wipeout)
                            //{
                            //    CoordinateSystem3d orientation = (obj as Wipeout).Orientation;
                            //    Plane plane2 = new Plane(orientation.Origin, orientation.Zaxis.GetNormal());
                            //    coordinateSystem3d = plane2.GetCoordinateSystem();
                            //}
                        }
                        openCloseTransaction.Commit();
                    }
                }
                return coordinateSystem3D;
            }
        }
    }
}
