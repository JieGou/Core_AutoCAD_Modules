// ReSharper disable InconsistentNaming
namespace ModPlus
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.Colors;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.GraphicsInterface;
    using Autodesk.AutoCAD.Runtime;
    using Autodesk.AutoCAD.Windows;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using Windows.MiniPlugins;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
    using MenuItem = Autodesk.AutoCAD.Windows.MenuItem;
    using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;
    using Viewport = Autodesk.AutoCAD.DatabaseServices.Viewport;

    /// <summary>
    /// Мини-функции
    /// </summary>
    public class MiniFunctions
    {
        private const string LangItem = "AutocadDlls";

        /// <summary>
        /// Включение/отключение всех контекстных меню мини-функций в зависимости от настроек
        /// </summary>
        public static void LoadUnloadContextMenu()
        {
            // ent by block
            var entByBlockObjContMen = !bool.TryParse(UserConfigFile.GetValue("EntByBlockOCM"), out bool b) || b;
            if (entByBlockObjContMen)
                MiniFunctionsContextMenuExtensions.EntByBlockObjectContextMenu.Attach();
            else
                MiniFunctionsContextMenuExtensions.EntByBlockObjectContextMenu.Detach();

            // nested ent layer
            var nestedEntLayerObjContMen = !bool.TryParse(UserConfigFile.GetValue("NestedEntLayerOCM"), out b) || b;
            if (nestedEntLayerObjContMen)
                MiniFunctionsContextMenuExtensions.NestedEntLayerObjectContextMenu.Attach();
            else
                MiniFunctionsContextMenuExtensions.NestedEntLayerObjectContextMenu.Detach();

            // Fast block
            var fastBlocksContextMenu = !bool.TryParse(UserConfigFile.GetValue("FastBlocksCM"), out b) || b;
            if (fastBlocksContextMenu)
                MiniFunctionsContextMenuExtensions.FastBlockContextMenu.Attach();
            else
                MiniFunctionsContextMenuExtensions.FastBlockContextMenu.Detach();

            // VP to MS
            var VPtoMSObjConMen = !bool.TryParse(UserConfigFile.GetValue("VPtoMS"), out b) || b;
            if (VPtoMSObjConMen)
                MiniFunctionsContextMenuExtensions.VPtoMSObjectContextMenu.Attach();
            else
                MiniFunctionsContextMenuExtensions.VPtoMSObjectContextMenu.Detach();

            // wipeout vertex edit
            /*
             * Так как не получается создать контекстное меню конкретно на класс Wipeout (возможно в поздних версиях устранили),
             * то приходится делать через подписку на событие и создание меню у Entity
             */
            var wipeoutEditOCM = !bool.TryParse(UserConfigFile.GetValue("WipeoutEditOCM"), out b) || b; // true
            if (wipeoutEditOCM)
            {
                AcApp.DocumentManager.DocumentCreated += WipeoutEditOCM_Documents_DocumentCreated;
                AcApp.DocumentManager.DocumentActivated += WipeoutEditOCM_Documents_DocumentActivated;

                foreach (Document document in AcApp.DocumentManager)
                {
                    document.ImpliedSelectionChanged += WipeoutEditOCM_Document_ImpliedSelectionChanged;
                }
            }
            else
            {
                AcApp.DocumentManager.DocumentCreated -= WipeoutEditOCM_Documents_DocumentCreated;
                AcApp.DocumentManager.DocumentActivated -= WipeoutEditOCM_Documents_DocumentActivated;

                foreach (Document document in AcApp.DocumentManager)
                {
                    document.ImpliedSelectionChanged -= WipeoutEditOCM_Document_ImpliedSelectionChanged;
                }
            }
        }

        /// <summary>
        /// Отключить все контекстные меню
        /// </summary>
        public static void UnloadAll()
        {
            MiniFunctionsContextMenuExtensions.EntByBlockObjectContextMenu.Detach();
            MiniFunctionsContextMenuExtensions.NestedEntLayerObjectContextMenu.Detach();
            MiniFunctionsContextMenuExtensions.FastBlockContextMenu.Detach();
            MiniFunctionsContextMenuExtensions.VPtoMSObjectContextMenu.Detach();

            // wipeout vertex edit
            /*
             * Так как не получается создать контекстное меню конкретно на класс Wipeout (возможно в поздних версиях устранили),
             * то приходится делать через подписку на событие и создание меню у Entity
             */
            AcApp.DocumentManager.DocumentCreated -= WipeoutEditOCM_Documents_DocumentCreated;
            AcApp.DocumentManager.DocumentActivated -= WipeoutEditOCM_Documents_DocumentActivated;

            foreach (Document document in AcApp.DocumentManager)
            {
                document.ImpliedSelectionChanged -= WipeoutEditOCM_Document_ImpliedSelectionChanged;
            }
        }

        private static void WipeoutEditOCM_Documents_DocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            if (e.Document != null)
            {
                e.Document.ImpliedSelectionChanged -= WipeoutEditOCM_Document_ImpliedSelectionChanged;
                e.Document.ImpliedSelectionChanged += WipeoutEditOCM_Document_ImpliedSelectionChanged;
            }
        }

        private static void WipeoutEditOCM_Documents_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            if (e.Document != null)
            {
                e.Document.ImpliedSelectionChanged -= WipeoutEditOCM_Document_ImpliedSelectionChanged;
                e.Document.ImpliedSelectionChanged += WipeoutEditOCM_Document_ImpliedSelectionChanged;
            }
        }

        private static void WipeoutEditOCM_Document_ImpliedSelectionChanged(object sender, EventArgs e)
        {
            PromptSelectionResult psr = AcApp.DocumentManager.MdiActiveDocument.Editor.SelectImplied();
            bool detach = true;
            if (psr.Value != null && psr.Value.Count == 1)
            {
                using (AcApp.DocumentManager.MdiActiveDocument.LockDocument())
                {
                    using (OpenCloseTransaction tr = new OpenCloseTransaction())
                    {
                        foreach (SelectedObject selectedObject in psr.Value)
                        {
                            if (selectedObject.ObjectId == ObjectId.Null)
                                continue;
                            var obj = tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead);
                            if (obj is Wipeout)
                            {
                                MiniFunctionsContextMenuExtensions.WipeoutEditObjectContextMenu.Attach();
                                detach = false;
                            }
                        }

                        tr.Commit();
                    }
                }
            }

            if (detach)
                MiniFunctionsContextMenuExtensions.WipeoutEditObjectContextMenu.Detach();
        }

        #region Nested entities ByBlock

        /// <summary>
        /// Задать вхождения ПоБлоку
        /// </summary>
        [CommandMethod("ModPlus", "mpEntByBlock", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void EntriesByBlock()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            try
            {
                var selectedObjects = ed.SelectImplied();
                if (selectedObjects.Value == null)
                {
                    var pso = new PromptSelectionOptions
                    {
                        MessageForAdding = $"\n{Language.GetItem(LangItem, "msg2")}",
                        MessageForRemoval = "\n",
                        AllowSubSelections = false,
                        AllowDuplicates = false
                    };
                    var psr = ed.GetSelection(pso);
                    if (psr.Status != PromptStatus.OK)
                        return;
                    selectedObjects = psr;
                }

                if (selectedObjects.Value.Count > 0)
                {
                    using (var tr = doc.TransactionManager.StartTransaction())
                    {
                        foreach (SelectedObject so in selectedObjects.Value)
                        {
                            var selEnt = tr.GetObject(so.ObjectId, OpenMode.ForRead);
                            if (selEnt is BlockReference blockReference)
                            {
                                ChangeProperties(blockReference.BlockTableRecord);
                            }
                        }

                        tr.Commit();
                    }

                    ed.Regen();
                }
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private static void ChangeProperties(ObjectId objectId)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;

            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(objectId, OpenMode.ForRead);

                foreach (var entId in btr)
                {
                    var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;

                    if (ent != null)
                    {
                        var br = ent as BlockReference;
                        if (br != null)
                        {
                            // recursive
                            ChangeProperties(br.BlockTableRecord);
                        }
                        else
                        {
                            ent.UpgradeOpen();
                            ent.Color = Color.FromColorIndex(ColorMethod.ByBlock, 0);
                            ent.LineWeight = LineWeight.ByBlock;
                            ent.Linetype = "ByBlock";
                            ent.DowngradeOpen();
                        }
                    }
                }

                tr.Commit();
            }
        }

        #endregion

        #region Nested entities Layer

        [CommandMethod("ModPlus", "mpNestedEntLayer", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void EntriesLayer()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            try
            {
                var selectedObjects = ed.SelectImplied();
                if (selectedObjects.Value == null)
                {
                    var pso = new PromptSelectionOptions
                    {
                        MessageForAdding = $"\n{Language.GetItem(LangItem, "msg2")}",
                        MessageForRemoval = "\n",
                        AllowSubSelections = false,
                        AllowDuplicates = false
                    };
                    var psr = ed.GetSelection(pso);
                    if (psr.Status != PromptStatus.OK)
                        return;
                    selectedObjects = psr;
                }

                if (selectedObjects.Value.Count > 0)
                {
                    var selectLayerWin = new SelectLayer();
                    if (selectLayerWin.ShowDialog() == true && selectLayerWin.LbLayers.SelectedIndex != -1)
                    {
                        var selectedLayer = (SelectLayer.SelLayer)selectLayerWin.LbLayers.SelectedItem;
                        using (var tr = doc.TransactionManager.StartTransaction())
                        {
                            foreach (SelectedObject so in selectedObjects.Value)
                            {
                                var selEnt = tr.GetObject(so.ObjectId, OpenMode.ForRead);
                                if (selEnt is BlockReference blockReference)
                                {
                                    ChangeLayer(blockReference.BlockTableRecord, selectedLayer.LayerId);
                                }
                            }

                            tr.Commit();
                        }

                        ed.Regen();
                    }
                }
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private static void ChangeLayer(ObjectId objectId, ObjectId layerId)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;

            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(objectId, OpenMode.ForRead);

                foreach (var entId in btr)
                {
                    var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;

                    if (ent != null)
                    {
                        var br = ent as BlockReference;
                        if (br != null)
                        {
                            // recursive
                            ChangeProperties(br.BlockTableRecord);
                        }
                        else
                        {
                            ent.UpgradeOpen();
                            ent.SetLayerId(layerId, true);
                            ent.DowngradeOpen();
                        }
                    }
                }

                tr.Commit();
            }
        }

        #endregion
        
        #region VP to MS

        [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
        private static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);

        [CommandMethod("ModPlus", "mpVPtoMS", CommandFlags.UsePickSet | CommandFlags.Redraw)]
        public void VPtoMS()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            try
            {
                var objectId = ObjectId.Null;
                var selection = ed.SelectImplied();
                if (selection.Value.Count == 2)
                {
                    using (var tr = doc.TransactionManager.StartTransaction())
                    {
                        foreach (SelectedObject selectedObject in selection.Value)
                        {
                            if (tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead) is Viewport)
                                continue;
                            objectId = selectedObject.ObjectId;
                        }

                        tr.Dispose();
                    }
                }
                else if (selection.Value.Count == 1)
                {
                    using (var tr = doc.TransactionManager.StartTransaction())
                    {
                        if (tr.GetObject(selection.Value[0].ObjectId, OpenMode.ForRead) is Viewport)
                            objectId = selection.Value[0].ObjectId;
                        tr.Dispose();
                    }
                }
                else
                {
                    var peo = new PromptEntityOptions($"\n{Language.GetItem(LangItem, "msg3")}");
                    peo.SetRejectMessage("\nReject");
                    peo.AllowNone = false;
                    peo.AllowObjectOnLockedLayer = true;
                    peo.AddAllowedClass(typeof(Viewport), true);
                    peo.AddAllowedClass(typeof(Polyline), true);
                    peo.AddAllowedClass(typeof(Polyline2d), true);
                    peo.AddAllowedClass(typeof(Curve), true);
                    var per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK)
                        return;
                    objectId = per.ObjectId;
                }

                if (objectId == ObjectId.Null)
                    return;
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    Viewport viewport;
                    var ent = tr.GetObject(objectId, OpenMode.ForRead);
                    if (ent is Viewport)
                    {
                        viewport = ent as Viewport;
                    }
                    else
                    {
                        var lm = LayoutManager.Current;
                        var vpid = lm.GetNonRectangularViewportIdFromClipId(objectId);
                        if (vpid != ObjectId.Null)
                        {
                            var clipVp = tr.GetObject(vpid, OpenMode.ForRead);
                            if (clipVp is Viewport)
                                viewport = clipVp as Viewport;
                            else
                                return;
                        }
                        else
                        {
                            MessageBox.Show(Language.GetItem(LangItem, "msg4"), MessageBoxIcon.Alert);
                            return;
                        }
                    }

                    // Переключаемся в пространство листа
                    ed.SwitchToPaperSpace();

                    // Номер текущего видового экрана
                    var vpNumber = (short)viewport.Number;

                    // Получаем точки границ текущего ВЭ
                    var psVpPnts = VPcontuorPoints(viewport, tr);

                    // Если есть точки, продолжаем
                    if (psVpPnts.Count == 0)
                        return;

                    // Обновляем вид
                    ed.UpdateScreen();

                    // Переходим внутрь активного ВЭ
                    ed.SwitchToModelSpace();

                    // Проверяем состояние ВЭ
                    if (viewport.Number > 0)
                    {
                        // Переключаемся в обрабатываемый ВЭ                                            
                        AcApp.SetSystemVariable("CVPORT", vpNumber);

                        // Переводим в модель граничные точки текущего ВЭ
                        var msVpPnts = ConvertVPcontuorPointsToMSpoints(psVpPnts);

                        var pline = new Polyline();
                        for (var i = 0; i < msVpPnts.Count; i++)
                        {
                            pline.AddVertexAt(i, new Point2d(msVpPnts[i].X, msVpPnts[i].Y), 0.0, 0.0, 0.0);
                        }

                        pline.Closed = true;

                        // свойства
                        pline.Layer = "0";

                        var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        var btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        btr?.AppendEntity(pline);
                        tr.AddNewlyCreatedDBObject(pline, true);
                    }

                    // now switch back to PS
                    ed.SwitchToPaperSpace();

                    tr.Commit();
                }

                // clear selection
                ed.SetImpliedSelection(new ObjectId[0]);
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private Point3dCollection VPcontuorPoints(Viewport viewport, Transaction tr)
        {
            // Коллекция для точек видового экрана
            Point3dCollection psVpPnts = new Point3dCollection();

            // Если видовой экран стандартный прямоугольный
            if (!viewport.NonRectClipOn)
            {
                // Получаем его точки
                viewport.GetGripPoints(psVpPnts, new IntegerCollection(), new IntegerCollection());

                // Выстраиваем точки в правильном порядке, по умолчанию они крест-накрест
                Point3d tmp = psVpPnts[2];
                psVpPnts[2] = psVpPnts[1];
                psVpPnts[1] = tmp;
            }

            // Если видовой экран подрезанный - получаем примитив, по которому он подрезан
            else
            {
                using (Entity ent = tr.GetObject(viewport.NonRectClipEntityId, OpenMode.ForRead) as Entity)
                {
                    // Если это полилиния - извлекаем ее точки
                    if (ent is Polyline pline)
                    {
                        for (int i = 0; i < pline.NumberOfVertices; i++)
                            psVpPnts.Add(pline.GetPoint3dAt(i));
                    }
                    else if (ent is Polyline2d pline2d)
                    {
                        foreach (ObjectId vertId in pline2d)
                        {
                            using (Vertex2d vert = tr.GetObject(vertId, OpenMode.ForRead) as Vertex2d)
                            {
                                if (!psVpPnts.Contains(vert.Position))
                                    psVpPnts.Add(vert.Position);
                            }
                        }
                    }
                    else if (ent is Curve curve)
                    {
                        double
                            startParam = curve.StartParam,
                            endParam = curve.EndParam,
                            delParam = (endParam - startParam) / 100;

                        for (double curParam = startParam; curParam < endParam; curParam += delParam)
                        {
                            Point3d curPt = curve.GetPointAtParameter(curParam);
                            psVpPnts.Add(curPt);
                            curParam += delParam;
                        }
                    }
                    else
                    {
                        MessageBox.Show(
                            $"{Language.GetItem(LangItem, "msg5")}\n{Language.GetItem(LangItem, "msg6")}{viewport.Number}, {Language.GetItem(LangItem, "msg7")}{ent}\n");
                    }
                }
            }

            return psVpPnts;
        }

        private Point3dCollection ConvertVPcontuorPointsToMSpoints(Point3dCollection psVpPnts)
        {
            Point3dCollection msVpPnts = new Point3dCollection();

            // Преобразование точки из PS в MS
            ResultBuffer rbPSDCS = new ResultBuffer(new TypedValue(5003, 3));
            ResultBuffer rbDCS = new ResultBuffer(new TypedValue(5003, 2));
            ResultBuffer rbWCS = new ResultBuffer(new TypedValue(5003, 0));
            double[] retPoint = new double[] { 0, 0, 0 };

            foreach (Point3d pnt in psVpPnts)
            {
                // преобразуем из DCS пространства Листа (PSDCS) RTSHORT=3
                // в DCS пространства Модели текущего Видового Экрана RTSHORT=2
                acedTrans(pnt.ToArray(), rbPSDCS.UnmanagedObject, rbDCS.UnmanagedObject, 0,
                    retPoint);

                // Преобразуем из DCS пространства Модели текущего Видового Экрана RTSHORT=2
                // в WCS RTSHORT=0
                acedTrans(retPoint, rbDCS.UnmanagedObject, rbWCS.UnmanagedObject, 0, retPoint);

                // Добавляем точку в коллекцию
                Point3d newPt = new Point3d(retPoint);
                if (!msVpPnts.Contains(newPt))
                    msVpPnts.Add(newPt);
            }

            return msVpPnts;
        }

        #endregion

        #region Edit wipeout

        /// <summary>
        /// Добавить вершину маскировке
        /// </summary>
        [CommandMethod("ModPlus", "mpAddVertexToWipeout", CommandFlags.UsePickSet)]
        public void AddVertexToWipeout()
        {
            try
            {
                Document doc = AcApp.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                var selectedObjects = ed.SelectImplied();
                var selectedId = ObjectId.Null;
                if (selectedObjects.Value == null || selectedObjects.Value.Count > 1)
                {
                    PromptEntityOptions peo = new PromptEntityOptions($"\n{Language.GetItem(LangItem, "msg20")}:");
                    peo.SetRejectMessage("\nWrong!");
                    peo.AllowNone = false;
                    peo.AddAllowedClass(typeof(Wipeout), true);

                    var ent = ed.GetEntity(peo);
                    if (ent.Status != PromptStatus.OK)
                        return;

                    selectedId = ent.ObjectId;
                }
                else
                {
                    selectedId = selectedObjects.Value[0].ObjectId;
                }

                if (selectedId != ObjectId.Null)
                    AddVertexToCurrentWipeout(selectedId);
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        /// <summary>
        /// Удалить вершину маскировки
        /// </summary>
        [CommandMethod("ModPlus", "mpRemoveVertexFromWipeout", CommandFlags.UsePickSet)]
        public void RemoveVertexToWipeout()
        {
            try
            {
                Document doc = AcApp.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                var selectedObjects = ed.SelectImplied();
                var selectedId = ObjectId.Null;
                if (selectedObjects.Value == null || selectedObjects.Value.Count > 1)
                {
                    PromptEntityOptions peo = new PromptEntityOptions($"\n{Language.GetItem(LangItem, "msg20")}:");
                    peo.SetRejectMessage("\nWrong!");
                    peo.AllowNone = false;
                    peo.AddAllowedClass(typeof(Wipeout), true);

                    var ent = ed.GetEntity(peo);
                    if (ent.Status != PromptStatus.OK)
                        return;

                    selectedId = ent.ObjectId;
                }
                else
                {
                    selectedId = selectedObjects.Value[0].ObjectId;
                }

                if (selectedId != ObjectId.Null)
                    RemoveVertexFromCurrentWipeout(selectedId);
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private static void AddVertexToCurrentWipeout(ObjectId wipeoutId)
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            var loop = true;
            while (loop)
            {
                using (doc.LockDocument())
                {
                    using (Transaction tr = doc.TransactionManager.StartTransaction())
                    {
                        Wipeout wipeout = tr.GetObject(wipeoutId, OpenMode.ForWrite) as Wipeout;
                        if (wipeout != null)
                        {
                            var points3D = wipeout.GetVertices();

                            Polyline polyline = new Polyline();
                            for (int i = 0; i < points3D.Count; i++)
                            {
                                polyline.AddVertexAt(i, new Point2d(points3D[i].X, points3D[i].Y), 0.0, 0.0, 0.0);
                            }

                            var jig = new AddVertexJig();
                            var jres = jig.StartJig(polyline);
                            if (jres.Status != PromptStatus.OK)
                            {
                                loop = false;
                            }
                            else
                            {
                                polyline.AddVertexAt(jig.Vertex() + 1, jig.PickedPoint(), 0.0, 0.0, 0.0);
                                var new2DPoints = new Point2dCollection();
                                for (int i = 0; i < polyline.NumberOfVertices; i++)
                                {
                                    new2DPoints.Add(polyline.GetPoint2dAt(i));
                                }

                                wipeout.SetFrom(new2DPoints, polyline.Normal);
                            }
                        }

                        tr.Commit();
                    }
                }
            }
        }

        private static void RemoveVertexFromCurrentWipeout(ObjectId wipeoutId)
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            var loop = true;
            while (loop)
            {
                using (doc.LockDocument())
                {
                    using (Transaction tr = doc.TransactionManager.StartTransaction())
                    {
                        Wipeout wipeout = tr.GetObject(wipeoutId, OpenMode.ForWrite) as Wipeout;
                        if (wipeout != null)
                        {
                            var points3D = wipeout.GetVertices();
                            if (points3D.Count > 3)
                            {
                                Polyline polyline = new Polyline();
                                for (int i = 0; i < points3D.Count; i++)
                                {
                                    polyline.AddVertexAt(i, new Point2d(points3D[i].X, points3D[i].Y), 0.0, 0.0, 0.0);
                                }

                                var pickedPt = ed.GetPoint($"\n{Language.GetItem(LangItem, "msg22")}:");
                                if (pickedPt.Status != PromptStatus.OK)
                                {
                                    loop = false;
                                }
                                else
                                {
                                    var pt = polyline.GetClosestPointTo(pickedPt.Value, false);
                                    var param = polyline.GetParameterAtPoint(pt);
                                    var vertex = Convert.ToInt32(Math.Truncate(param));
                                    polyline.RemoveVertexAt(vertex);
                                    var new2DPoints = new Point2dCollection();
                                    for (int i = 0; i < polyline.NumberOfVertices; i++)
                                    {
                                        new2DPoints.Add(polyline.GetPoint2dAt(i));
                                    }

                                    wipeout.SetFrom(new2DPoints, polyline.Normal);
                                }
                            }
                            else
                            {
                                // message
                                loop = false;
                            }
                        }

                        tr.Commit();
                    }
                }
            }
        }

        private class AddVertexJig : DrawJig
        {
            private Point3d _prevPoint;
            private Point3d _currentPoint;
            private Point3d _startPoint;
            private Polyline _pLine;
            private int _vertex;

            public PromptResult StartJig(Polyline pLine)
            {
                _pLine = pLine;
                _prevPoint = _pLine.GetPoint3dAt(0);
                _startPoint = _pLine.GetPoint3dAt(0);

                return AcApp.DocumentManager.MdiActiveDocument.Editor.Drag(this);
            }

            public int Vertex()
            {
                return _vertex;
            }

            public Point2d PickedPoint()
            {
                return new Point2d(_currentPoint.X, _currentPoint.Y);
            }

            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                var ppo = new JigPromptPointOptions($"\n{Language.GetItem(LangItem, "msg21")}:")
                {
                    BasePoint = _startPoint,
                    UseBasePoint = true,
                    UserInputControls = UserInputControls.Accept3dCoordinates
                                        | UserInputControls.NullResponseAccepted
                                        | UserInputControls.AcceptOtherInputString
                                        | UserInputControls.NoNegativeResponseAccepted
                };

                var ppr = prompts.AcquirePoint(ppo);

                if (ppr.Status != PromptStatus.OK)
                    return SamplerStatus.Cancel;

                if (ppr.Status == PromptStatus.OK)
                {
                    _currentPoint = ppr.Value;

                    if (CursorHasMoved())
                    {
                        _prevPoint = _currentPoint;
                        return SamplerStatus.OK;
                    }

                    return SamplerStatus.NoChange;
                }

                return SamplerStatus.NoChange;
            }

            protected override bool WorldDraw(WorldDraw draw)
            {
                var mods = System.Windows.Forms.Control.ModifierKeys;
                var control = (mods & System.Windows.Forms.Keys.Control) > 0;
                var pt = _pLine.GetClosestPointTo(_currentPoint, false);
                var param = _pLine.GetParameterAtPoint(pt);
                _vertex = Convert.ToInt32(Math.Truncate(param));
                var maxVx = _pLine.NumberOfVertices - 1;
                if (control)
                {
                    if (_vertex < maxVx)
                        _vertex++;
                }

                if (_vertex != maxVx)
                {
                    // Если вершина не последня
                    var line1 = new Line(_pLine.GetPoint3dAt(_vertex), _currentPoint);
                    draw.Geometry.Draw(line1);
                    var line2 = new Line(_pLine.GetPoint3dAt(_vertex + 1), _currentPoint);
                    draw.Geometry.Draw(line2);
                }
                else
                {
                    var line1 = new Line(_pLine.GetPoint3dAt(_vertex), _currentPoint);
                    draw.Geometry.Draw(line1);
                    if (_pLine.Closed)
                    {
                        // Если полилиния замкнута, то рисуем отрезок к первой вершине
                        var line2 = new Line(_pLine.GetPoint3dAt(0), _currentPoint);
                        draw.Geometry.Draw(line2);
                    }
                }

                return true;
            }

            private bool CursorHasMoved()
            {
                return _currentPoint.DistanceTo(_prevPoint) > Tolerance.Global.EqualPoint;
            }
        }

        #endregion

        public class MiniFunctionsContextMenuExtensions
        {
            public static class EntByBlockObjectContextMenu
            {
                public static ContextMenuExtension ContextMenu;

                public static void Attach()
                {
                    if (ContextMenu == null)
                    {
                        // For Entity
                        ContextMenu = new ContextMenuExtension();
                        var miEnt = new MenuItem(Language.GetItem(LangItem, "h49"));
                        miEnt.Click += StartFunction;
                        ContextMenu.MenuItems.Add(miEnt);

                        var rxcEnt = RXObject.GetClass(typeof(BlockReference));
                        Application.AddObjectContextMenuExtension(rxcEnt, ContextMenu);
                    }
                }

                public static void Detach()
                {
                    if (ContextMenu != null)
                    {
                        var rxcEnt = RXObject.GetClass(typeof(BlockReference));
                        Application.RemoveObjectContextMenuExtension(rxcEnt, ContextMenu);
                        ContextMenu = null;
                    }
                }

                public static void StartFunction(object o, EventArgs e)
                {
                    AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute(
                        "_.mpEntByBlock ", false, false, false);
                }
            }

            public static class NestedEntLayerObjectContextMenu
            {
                public static ContextMenuExtension ContextMenu;

                public static void Attach()
                {
                    if (ContextMenu == null)
                    {
                        // For Entity
                        ContextMenu = new ContextMenuExtension();
                        var miEnt = new MenuItem(Language.GetItem(LangItem, "h59"));
                        miEnt.Click += StartFunction;
                        ContextMenu.MenuItems.Add(miEnt);

                        var rxcEnt = RXObject.GetClass(typeof(BlockReference));
                        Application.AddObjectContextMenuExtension(rxcEnt, ContextMenu);
                    }
                }

                public static void Detach()
                {
                    if (ContextMenu != null)
                    {
                        var rxcEnt = RXObject.GetClass(typeof(BlockReference));
                        Application.RemoveObjectContextMenuExtension(rxcEnt, ContextMenu);
                        ContextMenu = null;
                    }
                }

                public static void StartFunction(object o, EventArgs e)
                {
                    AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute(
                        "_.mpNestedEntLayer ", false, false, false);
                }
            }

            public static class FastBlockContextMenu
            {
                public static ContextMenuExtension ContextMenu;

                public static void Attach()
                {
                    if (ContextMenu == null)
                    {
                        if (File.Exists(UserConfigFile.FullFileName))
                        {
                            var configXml = UserConfigFile.ConfigFileXml;
                            var settingsXml = configXml?.Element("Settings");
                            var fastBlocksXml = settingsXml?.Element("mpFastBlocks");
                            if (fastBlocksXml != null)
                            {
                                if (fastBlocksXml.Elements("FastBlock").Any())
                                {
                                    ContextMenu = new ContextMenuExtension { Title = Language.GetItem(LangItem, "h50") };
                                    foreach (var fbXml in fastBlocksXml.Elements("FastBlock"))
                                    {
                                        var mi = new MenuItem(fbXml.Attribute("Name").Value);
                                        mi.Click += Mi_Click;
                                        ContextMenu.MenuItems.Add(mi);
                                    }

                                    Application.AddDefaultContextMenuExtension(ContextMenu);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show(Language.GetItem(LangItem, "err4"), MessageBoxIcon.Close);
                        }
                    }
                }

                public static void Detach()
                {
                    if (ContextMenu != null)
                    {
                        Application.RemoveDefaultContextMenuExtension(ContextMenu);
                        ContextMenu = null;
                    }
                }

                private static void Mi_Click(object sender, EventArgs e)
                {
                    try
                    {
                        if (sender is MenuItem mi)
                        {
                            if (File.Exists(UserConfigFile.FullFileName))
                            {
                                var configXml = UserConfigFile.ConfigFileXml;
                                var settingsXml = configXml?.Element("Settings");
                                var fastBlocksXml = settingsXml?.Element("mpFastBlocks");
                                if (fastBlocksXml != null)
                                {
                                    if (fastBlocksXml.Elements("FastBlock").Any())
                                    {
                                        foreach (var fbXml in fastBlocksXml.Elements("FastBlock"))
                                        {
                                            if (fbXml.Attribute("Name").Value.Equals(mi.Text))
                                            {
                                                InsertBlock(
                                                    fbXml.Attribute("File").Value,
                                                    fbXml.Attribute("BlockName").Value);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show(Language.GetItem(LangItem, "err4"), MessageBoxIcon.Close);
                            }
                        }
                    }
                    catch (System.Exception exception)
                    {
                        ExceptionBox.Show(exception);
                    }
                }

                private static void InsertBlock(string file, string blockName)
                {
                    DocumentCollection dm = AcApp.DocumentManager;
                    Editor ed = dm.MdiActiveDocument.Editor;
                    Database destDb = dm.MdiActiveDocument.Database;
                    Database sourceDb = new Database(false, true);

                    // Read the DWG into a side database
                    sourceDb.ReadDwgFile(file, FileShare.Read, true, string.Empty);

                    // Create a variable to store the list of block identifiers
                    ObjectIdCollection blockIds = new ObjectIdCollection();
                    using (dm.MdiActiveDocument.LockDocument())
                    {
                        using (var sourceT = sourceDb.TransactionManager.StartTransaction())
                        {
                            // Open the block table
                            BlockTable bt = (BlockTable)sourceT.GetObject(sourceDb.BlockTableId, OpenMode.ForRead, false);

                            // Check each block in the block table
                            foreach (ObjectId btrId in bt)
                            {
                                BlockTableRecord btr = (BlockTableRecord)sourceT.GetObject(btrId, OpenMode.ForRead, false);

                                // Only add named & non-layout blocks to the copy list
                                if (btr.Name.Equals(blockName))
                                {
                                    blockIds.Add(btrId);
                                    break;
                                }

                                btr.Dispose();
                            }
                        }

                        // Copy blocks from source to destination database
                        IdMapping mapping = new IdMapping();
                        sourceDb.WblockCloneObjects(
                            blockIds,
                            destDb.BlockTableId,
                            mapping,
                            DuplicateRecordCloning.Replace,
                            false);
                        sourceDb.Dispose();

                        // Вставка

                        using (var tr = destDb.TransactionManager.StartTransaction())
                        {
                            BlockTable bt = (BlockTable)tr.GetObject(destDb.BlockTableId, OpenMode.ForRead, false);
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[blockName], OpenMode.ForRead, false);

                            var blkId = BlockInsertion.InsertBlockRef(0, tr, destDb, ed, bt[blockName]);

                            tr.Commit();
                        }
                    }
                }
            }

            public static class BlockInsertion
            {
                /// <summary>
                /// Вставка блока с атрибутами
                /// </summary>
                /// <param name="promptCounter">0 - только вставка, 1 - с поворотом</param>
                /// <param name="tr">Транзакция</param>
                /// <param name="db">База данных чертежа</param>
                /// <param name="ed">Editor</param>
                /// <param name="blkdefid">ObjectId блока</param>
                /// <param name="atts">Список имен атрибутов</param>
                /// <returns></returns>
                public static ObjectId InsertBlockRef(
                    int promptCounter,
                    Transaction tr,
                    Database db,
                    Editor ed,
                    ObjectId blkdefid)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    BlockReference blkref = new BlockReference(Point3d.Origin, blkdefid);
                    ObjectId id = btr.AppendEntity(blkref);
                    tr.AddNewlyCreatedDBObject(blkref, true);
                    BlockRefJig jig = new BlockRefJig(blkref);
                    jig.SetPromptCounter(0);
                    PromptResult res = ed.Drag(jig);
                    if (res.Status == PromptStatus.OK)
                    {
                        if (promptCounter == 1)
                        {
                            jig.SetPromptCounter(promptCounter);
                            res = ed.Drag(jig);
                            if (res.Status == PromptStatus.OK)
                            {
                                return id;
                            }
                        }
                        else
                        {
                            return id;
                        }
                    }

                    blkref.Erase();
                    return ObjectId.Null;
                }

                internal class BlockRefJig : EntityJig
                {
                    Point3d m_Position, m_BasePoint;
                    double m_Angle;
                    int m_PromptCounter;
                    Matrix3d m_Ucs;
                    Matrix3d m_Mat;
                    Editor ed = AcApp.DocumentManager.MdiActiveDocument.Editor;

                    public BlockRefJig(BlockReference blkref)
                    : base(blkref)
                    {
                        m_Position = new Point3d();
                        m_Angle = 0;
                        m_Ucs = ed.CurrentUserCoordinateSystem;
                        Update();
                    }

                    protected override SamplerStatus Sampler(JigPrompts prompts)
                    {
                        switch (m_PromptCounter)
                        {
                            case 0:
                                {
                                    JigPromptPointOptions jigOpts = new JigPromptPointOptions(
                                        $"\n{Language.GetItem(LangItem, "msg8")}");
                                    jigOpts.UserInputControls =
                                    UserInputControls.Accept3dCoordinates |
                                    UserInputControls.NoZeroResponseAccepted |
                                    UserInputControls.NoNegativeResponseAccepted;
                                    PromptPointResult res = prompts.AcquirePoint(jigOpts);
                                    Point3d pnt = res.Value;
                                    if (pnt != m_Position)
                                    {
                                        m_Position = pnt;
                                        m_BasePoint = m_Position;
                                    }
                                    else
                                    {
                                        return SamplerStatus.NoChange;
                                    }

                                    if (res.Status == PromptStatus.Cancel)
                                        return SamplerStatus.Cancel;

                                    return SamplerStatus.OK;
                                }

                            case 1:
                                {
                                    JigPromptAngleOptions jigOpts = new JigPromptAngleOptions(
                                        $"\n{Language.GetItem(LangItem, "msg9")}");
                                    jigOpts.UserInputControls =
                                    UserInputControls.Accept3dCoordinates |
                                    UserInputControls.NoNegativeResponseAccepted |
                                    UserInputControls.GovernedByUCSDetect |
                                    UserInputControls.UseBasePointElevation;
                                    jigOpts.Cursor = CursorType.RubberBand;
                                    jigOpts.UseBasePoint = true;
                                    jigOpts.BasePoint = m_BasePoint;
                                    PromptDoubleResult res = prompts.AcquireAngle(jigOpts);
                                    double angleTemp = res.Value;
                                    if (angleTemp != m_Angle)
                                        m_Angle = angleTemp;
                                    else
                                        return SamplerStatus.NoChange;
                                    if (res.Status == PromptStatus.Cancel)
                                        return SamplerStatus.Cancel;

                                    return SamplerStatus.OK;
                                }

                            default:
                                return SamplerStatus.NoChange;
                        }
                    }

                    protected override bool Update()
                    {
                        try
                        {
                            /*Ucs?Jig???:
                            * 1.?????Wcs???,?//xy??
                            * 2.?????????Ucs?
                            * 3.????Ucs???
                            * 4.?????Wcs
                            */
                            BlockReference blkref = (BlockReference)Entity;
                            blkref.Normal = Vector3d.ZAxis;
                            blkref.Position = m_Position.TransformBy(ed.CurrentUserCoordinateSystem);
                            blkref.Rotation = m_Angle;
                            blkref.TransformBy(m_Ucs);
                        }
                        catch
                        {
                            return false;
                        }

                        return true;
                    }

                    public void SetPromptCounter(int i)
                    {
                        if (i == 0 || i == 1)
                        {
                            m_PromptCounter = i;
                        }
                    }
                }
            }

            public static class VPtoMSObjectContextMenu
            {
                public static ContextMenuExtension ContextMenuForVP;
                public static ContextMenuExtension ContextMenuForCurve;

                public static void Attach()
                {
                    if (ContextMenuForVP == null)
                    {
                        ContextMenuForVP = new ContextMenuExtension();
                        var miEnt = new MenuItem(Language.GetItem(LangItem, "h51"));
                        miEnt.Click += StartFunction;
                        ContextMenuForVP.MenuItems.Add(miEnt);

                        var rxcEnt = RXObject.GetClass(typeof(Viewport));
                        Application.AddObjectContextMenuExtension(rxcEnt, ContextMenuForVP);
                    }

                    if (ContextMenuForCurve == null)
                    {
                        ContextMenuForCurve = new ContextMenuExtension();
                        var miEnt = new MenuItem(Language.GetItem(LangItem, "h51"));
                        miEnt.Click += StartFunction;
                        ContextMenuForCurve.MenuItems.Add(miEnt);

                        var rxcEnt = RXObject.GetClass(typeof(Curve));
                        Application.AddObjectContextMenuExtension(rxcEnt, ContextMenuForCurve);
                    }
                }

                public static void Detach()
                {
                    if (ContextMenuForVP != null)
                    {
                        var rxcEnt = RXObject.GetClass(typeof(Viewport));
                        Application.RemoveObjectContextMenuExtension(rxcEnt, ContextMenuForVP);
                        ContextMenuForVP = null;
                    }

                    if (ContextMenuForCurve != null)
                    {
                        var rxcEnt = RXObject.GetClass(typeof(Curve));
                        Application.RemoveObjectContextMenuExtension(rxcEnt, ContextMenuForCurve);
                        ContextMenuForCurve = null;
                    }
                }

                public static void StartFunction(object o, EventArgs e)
                {
                    AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute(
                        "_.mpVPtoMS ", false, false, false);
                }
            }

            public static class WipeoutEditObjectContextMenu
            {
                public static ContextMenuExtension ContextMenu;

                public static void Attach()
                {
                    if (ContextMenu == null)
                    {
                        ContextMenu = new ContextMenuExtension();
                        var mi1 = new MenuItem(Language.GetItem(LangItem, "h54"));
                        mi1.Click += Mi1_Click;
                        ContextMenu.MenuItems.Add(mi1);
                        var mi2 = new MenuItem(Language.GetItem(LangItem, "h55"));
                        mi2.Click += Mi2_Click;
                        ContextMenu.MenuItems.Add(mi2);
                    }

                    var rxClass = RXObject.GetClass(typeof(Entity));
                    Application.AddObjectContextMenuExtension(rxClass, ContextMenu);
                }

                public static void Detach()
                {
                    if (ContextMenu != null)
                    {
                        var rxcEnt = RXObject.GetClass(typeof(Entity));
                        Application.RemoveObjectContextMenuExtension(rxcEnt, ContextMenu);

                        // ContextMenu = null;
                    }
                }

                private static void Mi2_Click(object sender, EventArgs e)
                {
                    AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute(
                        "_.mpRemoveVertexFromWipeout ", false, false, false);
                }

                private static void Mi1_Click(object sender, EventArgs e)
                {
                    AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute(
                        "_.mpAddVertexToWipeout ", false, false, false);
                }
            }
        }
    }
}
