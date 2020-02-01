namespace ModPlus.Helpers
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

    /// <summary>Различные вспомогательные методы для работы в AutoCAD</summary>
    public static class AutocadHelpers
    {
        /// <summary>ObjectId блока для стрелки</summary>
        /// <param name="newArrName">Имя блока для стрелки</param>
        /// <returns>ObjectId нового блока стрелки</returns>
        public static ObjectId GetArrowObjectId(string newArrName)
        {
            ObjectId arrObjId;
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // Получаем текущее значение переменной DIMBLK
            var oldArrName = AcApp.GetSystemVariable("DIMBLK") as string;

            // Устанавливаем новое значение DIMBLK
            // (если такой блок отсутствует в чертеже, то он будет создан)
            AcApp.SetSystemVariable("DIMBLK", newArrName);

            // Возвращаем предыдущее значение переменной DIMBLK
            if (!string.IsNullOrEmpty(oldArrName))
                AcApp.SetSystemVariable("DIMBLK", oldArrName);

            // Теперь получаем objectId блока
            var tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId,OpenMode.ForRead);
                arrObjId = bt[newArrName];
                tr.Commit();
            }

            return arrObjId;
        }

        /// <summary>ObjectId одного из стандартного блока для стрелки</summary>
        /// <returns>ObjectId стандартного блока стрелки</returns>
        public static ObjectId GetArrowObjectId(StandardArrowhead standardArrowhead)
        {
            var newArrName = string.Empty;
            if (standardArrowhead != StandardArrowhead.closedFilled)
                newArrName = standardArrowhead.ToString();
            ObjectId arrObjId;
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // Получаем текущее значение переменной DIMBLK
            var oldArrName = AcApp.GetSystemVariable("DIMBLK") as string;

            // Устанавливаем новое значение DIMBLK
            // (если такой блок отсутствует в чертеже, то он будет создан)
            AcApp.SetSystemVariable("DIMBLK", newArrName);

            // Возвращаем предыдущее значение переменной DIMBLK
            if (!string.IsNullOrEmpty(oldArrName))
                AcApp.SetSystemVariable("DIMBLK", oldArrName);

            // Теперь получаем objectId блока
            var tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                arrObjId = bt[newArrName];
                tr.Commit();
            }

            return arrObjId;
        }

        /// <summary>Стандартные стрелки AutoCAD</summary>
        [SuppressMessage("ReSharper", "InconsistentNaming")]
        public enum StandardArrowhead
        {
            /// <summary>Заполненная замкнутая</summary>
            closedFilled,

            /// <summary>Точка</summary>
            _DOT,

            /// <summary>Малая точка</summary>
            _DOTSMALL,

            /// <summary>Контурная точка</summary>
            _DOTBLANK,

            /// <summary>Указатель исходной точки</summary>
            _ORIGIN,

            /// <summary>Указатель исходной точки 2</summary>
            _ORIGIN2,

            /// <summary>Разомкнутая</summary>
            _OPEN,

            /// <summary>Прямой угол</summary>
            _OPEN90,

            /// <summary>Разомкнутая 30</summary>
            _OPEN30,

            /// <summary>Замкнутая</summary>
            _CLOSED,

            /// <summary>Контурная малая точка</summary>
            _SMALL,

            /// <summary>Ничего</summary>
            _NONE,

            /// <summary>Засечка</summary>
            _OBLIQUE,

            /// <summary>Заполненный прямоугольник</summary>
            _BOXFILLED,

            /// <summary>Прямоугольник</summary>
            _BOXBLANK,

            /// <summary>Контурная замкнутая</summary>
            _CLOSEDBLANK,

            /// <summary>Треугольник базы отсчета</summary>
            _DATUMBLANK,

            /// <summary>Интеграл</summary>
            _INTEGRAL,

            /// <summary>Двойная засечка</summary>
            _ARCHTICK
        }
        
        /// <summary>
        /// Перевод точки из пользовательской системы координат в мировую
        /// </summary>
        public static Point3d UcsToWcs(Point3d pt)
        {
            var m = GetUcsMatrix(HostApplicationServices.WorkingDatabase);
            return pt.TransformBy(m);
        }

        private static bool IsPaperSpace(Database db)
        {
            if (db.TileMode)
                return false;
            var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;
            return db.PaperSpaceVportId == ed.CurrentViewportObjectId;
        }

        private static Matrix3d GetUcsMatrix(Database db)
        {
            Point3d origin;
            Vector3d xAxis, yAxis;
            if (IsPaperSpace(db))
            {
                origin = db.Pucsorg;
                xAxis = db.Pucsxdir;
                yAxis = db.Pucsydir;
            }
            else
            {
                origin = db.Ucsorg;
                xAxis = db.Ucsxdir;
                yAxis = db.Ucsydir;
            }

            var zAxis = xAxis.CrossProduct(yAxis);
            return Matrix3d.AlignCoordinateSystem(
                Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                origin, xAxis, yAxis, zAxis);
        }

        /// <summary>Зуммировать объекты</summary>
        /// <param name="objIds">Однострочный массив ObjectId зуммируемых объектов</param>
        public static void ZoomToEntities(ObjectId[] objIds)
        {
            var editor = AcApp.DocumentManager.MdiActiveDocument.Editor;

            var psr = editor.SelectImplied();
            ObjectId[] selected = null;
            if (psr.Status == PromptStatus.OK)
                selected = psr.Value.GetObjectIds();
            editor.SetImpliedSelection(objIds);

            Autodesk.AutoCAD.Internal.Utils.ZoomObjects(true);

            editor.SetImpliedSelection(selected);
        }

        /// <summary>Получение "реального" имени блока</summary>
        /// <param name="tr">Транзакция</param>
        /// <param name="bref">Вхождение блока</param>
        /// <returns></returns>
        public static string EffectiveBlockName(Transaction tr, BlockReference bref)
        {
            BlockTableRecord btr;
            if (bref.IsDynamicBlock | bref.Name.StartsWith("*U", StringComparison.InvariantCultureIgnoreCase))
            {
                btr = tr.GetObject(bref.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            }
            else
            {
                btr = tr.GetObject(bref.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            }

            return btr?.Name;
        }

        /// <summary>Получение аннотативного масштаба по имени из текущей базы данных (HostApplicationServices.WorkingDatabase)</summary>
        /// <param name="name">Имя масштаба</param>
        /// <returns>Аннотативный масштаб с таким именем или текущий масштаб в БД</returns>
        public static AnnotationScale GetAnnotationScaleByName(string name)
        {
            var db = HostApplicationServices.WorkingDatabase;
            var ocm = db.ObjectContextManager;
            if (ocm != null)
            {
                var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");
                if (occ != null)
                {
                    foreach (var objectContext in occ)
                    {
                        var asc = objectContext as AnnotationScale;
                        if (asc?.Name == name)
                            return asc;
                    }
                }
            }

            return db.Cannoscale;
        }

        /// <summary>
        /// Проверка, что редактор доступен для приема команд. Если не доступен, то будет показано сообщение
        /// </summary>
        public static bool IsEditorIsQuiescent(Editor editor)
        {
            if (editor.IsQuiescent)
                return true;

            MessageBox.Show(Language.GetItem("AutocadDlls", "msg23"), MessageBoxIcon.Alert);
            return false;
        }
    }
}
