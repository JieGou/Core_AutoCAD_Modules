#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using ModPlusAPI;
using ModPlusAPI.Windows;

namespace ModPlus
{
    public class MiniFunctions
    {
        private const string LangItem = "AutocadDlls";

        public static void LoadUnloadContextMenues()
        {
            // ent by block
            // ent by block
            var entByBlockObjContMen = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "EntByBlockOCM"), out bool b) || b;
            if (entByBlockObjContMen)
                ContextMenues.EntByBlockObjectContextMenu.Attach();
            else ContextMenues.EntByBlockObjectContextMenu.Detach();
            // Fast block
            var fastBlocksContextMenu = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "FastBlocksCM"), out b) || b;
            if (fastBlocksContextMenu)
                ContextMenues.FastBlockContextMenu.Attach();
            else ContextMenues.FastBlockContextMenu.Detach();
            // VP to MS
            var VPtoMSObjConMen = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "VPtoMS"), out b) || b;
            if (VPtoMSObjConMen)
                ContextMenues.VPtoMSobjectContextMenu.Attach();
            else ContextMenues.VPtoMSobjectContextMenu.Detach();
        }

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
                        MessageForAdding = "\n" + Language.GetItem(LangItem, "msg2"),
                        MessageForRemoval = "\n",
                        AllowSubSelections = false,
                        AllowDuplicates = false
                    };
                    var psr = ed.GetSelection(pso);
                    if (psr.Status != PromptStatus.OK) return;
                    selectedObjects = psr;
                }
                if (selectedObjects.Value.Count > 0)
                {
                    using (var tr = doc.TransactionManager.StartTransaction())
                    {
                        foreach (SelectedObject so in selectedObjects.Value)
                        {
                            var selEnt = tr.GetObject(so.ObjectId, OpenMode.ForRead);
                            if (selEnt is BlockReference)
                            {
                                ChangeProperties((selEnt as BlockReference).BlockTableRecord);
                            }
                        }
                        tr.Commit();
                    }
                    ed.Regen();
                }
            }
            catch (System.Exception exception)
            {
                ExceptionBox.ShowForConfigurator(exception);
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

#if ac2010
        [DllImport("acad.exe", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
        private static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);
#elif ac2013
        [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
        private static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);
#endif
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
                    var peo = new PromptEntityOptions("\n" + Language.GetItem(LangItem, "msg3"));
                    peo.SetRejectMessage("\nReject");
                    peo.AllowNone = false;
                    peo.AllowObjectOnLockedLayer = true;
                    peo.AddAllowedClass(typeof(Viewport), true);
                    peo.AddAllowedClass(typeof(Polyline), true);
                    peo.AddAllowedClass(typeof(Polyline2d), true);
                    peo.AddAllowedClass(typeof(Curve), true);
                    var per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK) return;
                    objectId = per.ObjectId;
                }
                if (objectId == ObjectId.Null) return;
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
                            else return;
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
                    if (psVpPnts.Count == 0) return;
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
                ExceptionBox.ShowForConfigurator(exception);
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
                    if (ent is Polyline)
                    {
                        Polyline pline = ent as Polyline;
                        for (int i = 0; i < pline.NumberOfVertices; i++) psVpPnts.Add(pline.GetPoint3dAt(i));
                    }
                    else if (ent is Polyline2d)
                    {
                        Polyline2d pline2d = ent as Polyline2d;
                        foreach (ObjectId vertId in pline2d)
                        {
                            using (Vertex2d vert = tr.GetObject(vertId, OpenMode.ForRead) as Vertex2d)
                            {
                                if (!psVpPnts.Contains(vert.Position)) psVpPnts.Add(vert.Position);
                            }
                        }

                    }
                    else if (ent is Curve)
                    {
                        Curve curve = ent as Curve;
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
                        MessageBox.Show(Language.GetItem(LangItem, "msg5")
                            + "\n" + Language.GetItem(LangItem, "msg6") + viewport.Number + ", " +
                            Language.GetItem(LangItem, "msg7") + ent + "\n");
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
                if (!msVpPnts.Contains(newPt)) msVpPnts.Add(newPt);
            }
            return msVpPnts;
        }

        public class ContextMenues
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
                        Autodesk.AutoCAD.ApplicationServices.Application.AddObjectContextMenuExtension(rxcEnt, ContextMenu);
                    }
                }
                public static void Detach()
                {
                    if (ContextMenu != null)
                    {
                        var rxcEnt = RXObject.GetClass(typeof(BlockReference));
                        Autodesk.AutoCAD.ApplicationServices.Application.RemoveObjectContextMenuExtension(rxcEnt, ContextMenu);
                        ContextMenu = null;
                    }
                }
                public static void StartFunction(object o, EventArgs e)
                {
                    AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute(
                        "_.mpEntByBlock ", false, false, false);
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
                                    Autodesk.AutoCAD.ApplicationServices.Application.AddDefaultContextMenuExtension(
                                        ContextMenu);
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
                        Autodesk.AutoCAD.ApplicationServices.Application.RemoveDefaultContextMenuExtension(ContextMenu);
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
                                                    fbXml.Attribute("BlockName").Value
                                                    );
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
                    sourceDb.ReadDwgFile(file, FileShare.Read, true, "");
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
                        sourceDb.WblockCloneObjects(blockIds,
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
                    }// lock
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
                        else return id;
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
                                    JigPromptPointOptions jigOpts = new JigPromptPointOptions("\n" + Language.GetItem(LangItem, "msg8"));
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
                                        return SamplerStatus.NoChange;
                                    if (res.Status == PromptStatus.Cancel)
                                        return SamplerStatus.Cancel;

                                    return SamplerStatus.OK;
                                }
                            case 1:
                                {
                                    JigPromptAngleOptions jigOpts = new JigPromptAngleOptions("\n" + Language.GetItem(LangItem, "msg9"));
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

            public static class VPtoMSobjectContextMenu
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
                        Autodesk.AutoCAD.ApplicationServices.Application.AddObjectContextMenuExtension(rxcEnt, ContextMenuForVP);
                    }
                    if (ContextMenuForCurve == null)
                    {
                        ContextMenuForCurve = new ContextMenuExtension();
                        var miEnt = new MenuItem(Language.GetItem(LangItem, "h51"));
                        miEnt.Click += StartFunction;
                        ContextMenuForCurve.MenuItems.Add(miEnt);

                        var rxcEnt = RXObject.GetClass(typeof(Curve));
                        Autodesk.AutoCAD.ApplicationServices.Application.AddObjectContextMenuExtension(rxcEnt, ContextMenuForCurve);
                    }
                }
                public static void Detach()
                {
                    if (ContextMenuForVP != null)
                    {
                        var rxcEnt = RXObject.GetClass(typeof(Viewport));
                        Autodesk.AutoCAD.ApplicationServices.Application.RemoveObjectContextMenuExtension(rxcEnt, ContextMenuForVP);
                        ContextMenuForVP = null;
                    }
                    if (ContextMenuForCurve != null)
                    {
                        var rxcEnt = RXObject.GetClass(typeof(Curve));
                        Autodesk.AutoCAD.ApplicationServices.Application.RemoveObjectContextMenuExtension(rxcEnt, ContextMenuForCurve);
                        ContextMenuForCurve = null;
                    }
                }
                public static void StartFunction(object o, EventArgs e)
                {
                    AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute(
                        "_.mpVPtoMS ", false, false, false);
                }
            }
        }
    }
}
