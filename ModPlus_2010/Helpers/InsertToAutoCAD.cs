#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using ModPlusAPI.Windows;

namespace ModPlus.Helpers
{
    /// <summary>Функции вставки/добавления в автокад</summary>
    public class InsertToAutoCad
    {
        /// <summary>Вставить строкове значение в ячейку таблицы автокада</summary>
        /// <param name="firstStr">Первая строка. Замена разделителя действует только для неё</param>
        /// <param name="secondString">Вторая строка. Необязательно. Замена разделителя не действует</param>
        /// <param name="useSeparator">Использовать разделитель (для цифровых значений). Заменяет запятую на точку, а затем на текущий разделитель</param>
        public static void AddStringToAutoCadTableCell(string firstStr, string secondString, bool useSeparator)
        {
            try
            {
                var db = AcApp.DocumentManager.MdiActiveDocument.Database;
                var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;
                using (AcApp.DocumentManager.MdiActiveDocument.LockDocument())
                {
                    var peo = new PromptEntityOptions("\nВыберите таблицу: ");
                    peo.SetRejectMessage("\nНеверный выбор! Это не таблица!");
                    peo.AddAllowedClass(typeof(Table), false);
                    var per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK)
                    {
                        return;
                    }
                    var tr = db.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        try
                        {
                            var entId = per.ObjectId;
                            var tbl = (Table)tr.GetObject(entId, OpenMode.ForWrite);
                            var ppo = new PromptPointOptions("\nВыберите ячейку: ");
                            var end = false;
                            var vector = new Vector3d(0.0, 0.0, 1.0);
                            while (end == false)
                            {
                                var ppr = ed.GetPoint(ppo);
                                if (ppr.Status != PromptStatus.OK) return;
                                try
                                {
                                    TableHitTestInfo tblhittestinfo =
                                        tbl.HitTest(ppr.Value, vector);
                                    if (tblhittestinfo.Type == TableHitTestType.Cell)
                                    {
                                        var cell = new Cell(tbl, tblhittestinfo.Row, tblhittestinfo.Column);
                                        if (useSeparator)
                                            cell.TextString = ModPlusAPI.IO.String.ReplaceSeparator(firstStr) + secondString;
                                        else cell.TextString = firstStr + secondString;
                                        end = true;
                                    }
                                } // try
                                catch
                                {
                                    MessageBox.Show("Не попали в ячейку!");
                                }
                            } // while
                            tr.Commit();
                        } //try
                        catch (Exception ex)
                        {
                            ExceptionBox.ShowForConfigurator(ex);
                        }
                    } //using tr
                } //using lock
            } //try
            catch (Exception ex)
            {
                ExceptionBox.ShowForConfigurator(ex);
            }
        }
        /// <summary>Вставка элемента спецификации в строку таблицы AutoCad</summary>
        /// <param name="pos">Поз.</param>
        /// <param name="designation">Обозначение</param>
        /// <param name="name">Наименование</param>
        /// <param name="massa">Масса</param>
        /// <param name="note">Примечание</param>
        /// <remarks>Вставка значений в таблицу ведется согласно количеству столбцов. Подробнее на сайте api.modplus.org</remarks>
        public static void AddSpecificationItemToTableRow(string pos, string designation, string name, string massa, string note)
        {
            try
            {
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;
                using (doc.LockDocument())
                {
                    var options = new PromptEntityOptions("\nВыберите таблицу: ");
                    options.SetRejectMessage("\nНеверный выбор! Это не таблица!");
                    options.AddAllowedClass(typeof(Table), false);
                    var entity = ed.GetEntity(options);
                    if (entity.Status == PromptStatus.OK)
                    {
                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var table = (Table)tr.GetObject(entity.ObjectId, OpenMode.ForWrite);
                            var columnsCount = table.Columns.Count;

                            if (columnsCount == 6)
                            {
                                var options2 = new PromptPointOptions("\nВыберите строку: ");
                                var flag = false;
                                var viewVector = new Vector3d(0.0, 0.0, 1.0);
                                while (!flag)
                                {
                                    var point = ed.GetPoint(options2);
                                    if (point.Status != PromptStatus.OK)
                                        return;
                                    try
                                    {
                                        var info = table.HitTest(point.Value, viewVector);
                                        if (info.Type == TableHitTestType.Cell)
                                        {
                                            // Заполняем
                                            // Поз.
                                            table.Cells[info.Row, 0].SetValue(pos, ParseOption.SetDefaultFormat);
                                            // Обозначение
                                            table.Cells[info.Row, 1].SetValue(designation, ParseOption.SetDefaultFormat);
                                            // Наименование
                                            table.Cells[info.Row, 2].SetValue(name, ParseOption.SetDefaultFormat);
                                            //Масса
                                            table.Cells[info.Row, table.Columns.Count - 2].SetValue(massa, ParseOption.SetDefaultFormat);
                                            // Примечание
                                            table.Cells[info.Row, table.Columns.Count - 1].SetValue(note, ParseOption.SetDefaultFormat);
                                            flag = true;
                                        }
                                    }
                                    catch
                                    {
                                        MessageBox.Show("Не попали в ячейку!");
                                    }
                                }
                            }
                            else if (columnsCount == 4)
                            {
                                var options2 = new PromptPointOptions("\nВыберите строку: ");
                                var flag = false;
                                var viewVector = new Vector3d(0.0, 0.0, 1.0);
                                while (!flag)
                                {
                                    var point = ed.GetPoint(options2);
                                    if (point.Status != PromptStatus.OK)
                                        return;
                                    try
                                    {
                                        TableHitTestInfo info = table.HitTest(point.Value, viewVector);
                                        if (info.Type == TableHitTestType.Cell)
                                        {
                                            //Cell cell;
                                            // Заполняем                                            
                                            // Наименование
                                            table.Cells[info.Row, 1].SetValue(name, ParseOption.SetDefaultFormat);
                                            //Масса
                                            table.Cells[info.Row, table.Columns.Count - 1].SetValue(massa, ParseOption.SetDefaultFormat);
                                            flag = true;
                                        }
                                    }
                                    catch
                                    {
                                        MessageBox.Show("Не попали в ячейку!");
                                    }
                                }
                            }
                            else if (columnsCount == 5)
                            {
                                var options2 = new PromptPointOptions("\nВыберите строку: ");
                                var flag = false;
                                var viewVector = new Vector3d(0.0, 0.0, 1.0);
                                while (!flag)
                                {
                                    var point = ed.GetPoint(options2);
                                    if (point.Status != PromptStatus.OK)
                                        return;
                                    try
                                    {
                                        var info = table.HitTest(point.Value, viewVector);
                                        if (info.Type == TableHitTestType.Cell)
                                        {
                                            //Cell cell;
                                            // Заполняем
                                            // Обозначение
                                            //cell = new Cell(table, info.Row, 1);
                                            //cell.TextString = designation;
                                            //table.SetTextString(info.Row, 1, designation);
                                            // Наименование
                                            //cell = new Cell(table, info.Row, 2);
                                            //cell.TextString = name;
                                            //table.SetTextString(info.Row, 2, name);

                                            flag = true;
                                        }
                                    }
                                    catch
                                    {
                                        MessageBox.Show("Не попали в ячейку!");
                                    }
                                }
                            }
                            else if (columnsCount == 7)
                            {
                                var options2 = new PromptPointOptions("\nВыберите строку: ");
                                var flag = false;
                                var viewVector = new Vector3d(0.0, 0.0, 1.0);
                                while (!flag)
                                {
                                    var point = ed.GetPoint(options2);
                                    if (point.Status != PromptStatus.OK)
                                        return;
                                    try
                                    {
                                        var info = table.HitTest(point.Value, viewVector);
                                        if (info.Type == TableHitTestType.Cell)
                                        {
                                            //Cell cell;
                                            // Заполняем
                                            // Обозначение
                                            //cell = new Cell(table, info.Row, 3);
                                            //cell.TextString = designation;
                                            //table.SetTextString(info.Row, 3, designation);
                                            // Наименование
                                            //cell = new Cell(table, info.Row, 4);
                                            //cell.TextString = name;
                                            //table.SetTextString(info.Row, 4, name);

                                            flag = true;
                                        }
                                    }
                                    catch
                                    {
                                        MessageBox.Show("Не попали в ячейку!");
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Неверное количество столбцов в таблице!", MessageBoxIcon.Alert);
                            }
                            tr.Commit();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionBox.ShowForConfigurator(ex);
            }
        }
        /// <summary>Вставка однострочного текста</summary>
        /// <param name="text">Содержимое однострочного текста</param>
        public static void InsertDbText(string text)
        {
            try
            {
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;
                using (AcApp.DocumentManager.MdiActiveDocument.LockDocument())
                {
                    var tr = doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, false);
                        var dtxt = new DBText();
                        dtxt.SetDatabaseDefaults();
                        dtxt.TextString = text;
                        dtxt.TransformBy(ed.CurrentUserCoordinateSystem);
                        var jig = new DTextJig(dtxt);
                        var pr = ed.Drag(jig);
                        if (pr.Status == PromptStatus.OK)
                        {
                            var ent = jig.GetEntity();
                            btr.AppendEntity(ent);
                            tr.AddNewlyCreatedDBObject(ent, true);
                            doc.TransactionManager.QueueForGraphicsFlush();
                        }
                        tr.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionBox.ShowForConfigurator(ex);
            }
        }

        /// <summary>Вставка мультивыноски</summary>
        /// <param name="txt">Содержимое мультивыноски</param>
        /// <param name="standardArrowhead">Стандартная стрелка. По умолчанию значение "_NONE"</param>
        public static void InsertMLeader(string txt, AutocadHelpers.StandardArrowhead standardArrowhead = AutocadHelpers.StandardArrowhead._NONE)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = HostApplicationServices.WorkingDatabase;
            using (doc.LockDocument())
            {
                var ppo = new PromptPointOptions("\nУкажите точку: ") { AllowNone = true };
                var ppr = ed.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK) return;
                // arrowHead
                var arrowHead = "_NONE";
                if (standardArrowhead != AutocadHelpers.StandardArrowhead.closedFilled)
                    arrowHead = standardArrowhead.ToString();
                // Создаем текст
                var jig = new MLeaderJig
                {
                    FirstPoint = AutocadHelpers.UcsToWcs(ppr.Value),
                    MlText = txt,
                    ArrowName = arrowHead
                };
                var res = ed.Drag(jig);
                if (res.Status == PromptStatus.OK)
                {
                    using (var tr = doc.TransactionManager.StartTransaction())
                    {
                        var btr = (BlockTableRecord)
                            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                        btr.AppendEntity(jig.MLeader());
                        tr.AddNewlyCreatedDBObject(jig.MLeader(), true);
                        tr.Commit();
                    }
                }
                doc.TransactionManager.QueueForGraphicsFlush();
            }
        }
        #region Jigs

        private class DTextJig : EntityJig
        {
            Point3d _mCenterPt, _mActualPoint;
            public DTextJig(DBText dbtxt)
                : base(dbtxt)
            {
                _mCenterPt = dbtxt.Position;
            }

            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                var jigOpts = new JigPromptPointOptions
                {
                    UserInputControls = (UserInputControls.Accept3dCoordinates |
                                         UserInputControls.NoZeroResponseAccepted |
                                         UserInputControls.AcceptOtherInputString |
                                         UserInputControls.NoNegativeResponseAccepted),
                    Message = "\nТочка вставки: "
                };
                var dres = prompts.AcquirePoint(jigOpts);
                if (_mActualPoint == dres.Value)
                {
                    return SamplerStatus.NoChange;
                }
                _mActualPoint = dres.Value;
                return SamplerStatus.OK;
            }

            protected override bool Update()
            {
                _mCenterPt = _mActualPoint;
                try
                {
                    ((DBText)Entity).Position = _mCenterPt;
                }
                catch (Exception)
                {
                    return false;
                }
                return true;
            }

            public Entity GetEntity()
            {
                return Entity;
            }

        }

        private class MLeaderJig : DrawJig
        {
            private MLeader _mleader;
            public Point3d FirstPoint;
            private Point3d _secondPoint;
            private Point3d _prevPoint;
            public string MlText;
            public string ArrowName;

            public MLeader MLeader()
            {
                return _mleader;
            }

            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                var jpo = new JigPromptPointOptions
                {
                    BasePoint = FirstPoint,
                    UseBasePoint = true,
                    UserInputControls = UserInputControls.Accept3dCoordinates |
                                        UserInputControls.GovernedByUCSDetect,
                    Message = "\nТочка вставки: "
                };

                var res = prompts.AcquirePoint(jpo);
                _secondPoint = res.Value;
                if (res.Status != PromptStatus.OK) return SamplerStatus.Cancel;
                if (CursorHasMoved())
                {
                    _prevPoint = _secondPoint;
                    return SamplerStatus.OK;
                }
                return SamplerStatus.NoChange;
            }
            private bool CursorHasMoved()
            {
                return _secondPoint.DistanceTo(_prevPoint) > 1e-6;
            }
            protected override bool WorldDraw(WorldDraw draw)
            {
                var wg = draw.Geometry;
                if (wg != null)
                {
                    ObjectId arrId = AutocadHelpers.GetArrowObjectId(ArrowName);

                    var mtxt = new MText();
                    mtxt.SetDatabaseDefaults();
                    mtxt.Contents = MlText;
                    mtxt.Location = _secondPoint;
                    mtxt.Annotative = AnnotativeStates.True;
                    mtxt.TransformBy(AcApp.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem);

                    _mleader = new MLeader();
                    var ldNum = _mleader.AddLeader();
                    _mleader.AddLeaderLine(ldNum);
                    _mleader.SetDatabaseDefaults();
                    _mleader.ContentType = ContentType.MTextContent;
                    _mleader.ArrowSymbolId = arrId;
                    _mleader.MText = mtxt;
                    _mleader.TextAlignmentType = TextAlignmentType.LeftAlignment;
                    _mleader.TextAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine;
                    _mleader.TextAngleType = TextAngleType.HorizontalAngle;
                    _mleader.EnableAnnotationScale = true;
                    _mleader.Annotative = AnnotativeStates.True;
                    _mleader.AddFirstVertex(ldNum, FirstPoint);
                    _mleader.AddLastVertex(ldNum, _secondPoint);
                    _mleader.LeaderLineType = LeaderType.StraightLeader;
                    _mleader.EnableDogleg = false;
                    _mleader.DoglegLength = 0.0;
                    _mleader.LandingGap = 1.0;
                    _mleader.TextHeight = double.Parse(AcApp.GetSystemVariable("TEXTSIZE").ToString());

                    draw.Geometry.Draw(_mleader);
                }
                return true;
            }
        }
        #endregion
    }
}
