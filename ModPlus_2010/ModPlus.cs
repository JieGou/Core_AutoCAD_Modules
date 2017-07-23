#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Xml.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using ModPlus.App;
using ModPlus.Windows;
using System.Windows.Forms.Integration;
using ModPlus.Helpers;
using ModPlusAPI;
using ModPlusAPI.Windows;
#pragma warning disable 1591

namespace ModPlus
{
    public class ModPlus : IExtensionApplication
    {
        private static bool _quiteLoad = false;
        // Инициализация плагина
        public void Initialize()
        {
            try
            {
                var sw = new Stopwatch();
                sw.Start();
                // Получим значение переменной "Тихая загрузка" в первую очередь
                _quiteLoad = GetQuiteLoad();
                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                // Файла конфигурации может не существовать при загрузке плагина!
                // Поэтому все, что связанно с работой с файлом конфигурации должно это учитывать!
                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;
                if (!CheckCadVersion())
                {
                    ed.WriteMessage("\n***************************");
                    ed.WriteMessage("\nВНИМАНИЕ!");
                    ed.WriteMessage("\nПопытка загрузки плагина в несоответсвующую версию автокада!");
                    ed.WriteMessage("\nЗавершение загрузки");
                    ed.WriteMessage("\n***************************");
                    return;
                }
                ed.WriteMessage("\n***************************");
                ed.WriteMessage("\nЗагрузка плагина ModPlus...");
                if (!_quiteLoad) ed.WriteMessage("\nЗагрузка рабочих компонентов...");
                // Принудительная загрузка сборок
                LoadAssms(ed);
                if (!_quiteLoad) ed.WriteMessage("\nЗагрузка баз данных...");
                LoadBaseAssemblies(ed);
                if (!_quiteLoad) ed.WriteMessage("\nИнициализация файла конфигурации:");
                UserConfigFile.InitConfigFile();
                if (!_quiteLoad) ed.WriteMessage("\nЗагрузка функций...");
                LoadFunctions(ed);
                // Строим: ленту, меню, плавающее меню
                // Загрузка ленты
                Autodesk.Windows.ComponentManager.ItemInitialized += ComponentManager_ItemInitialized;
                // Палитра
                if (ModPlusAPI.Variables.Palette)
                    MpPalette.CreatePalette();
                // Загрузка основного меню (с проверкой значения из файла настроек)
                MpMenuFunction.LoadMainMenu();
                // Загрузка окна Чертежи
                MpDrawingsFunction.LoadMainMenu();
                // Загрузка контекстных меню для мини-функций
                MiniFunctions.LoadUnloadContextMenues();
                // проверка загруженности модуля автообновления
                CheckAutoUpdaterLoaded();

                sw.Stop();
                ed.WriteMessage("\nЗагрузка плагина ModPlus завершена. Затрачено времени (мc): " + sw.ElapsedMilliseconds);
                ed.WriteMessage("\nПриятной работы!");
                ed.WriteMessage("\n***************************");
            }
            catch (System.Exception exception)
            {
                ExceptionBox.ShowForConfigurator(exception);
            }
        }
        public void Terminate()
        {
            
        }
        // Значение тихой загрузки
        private static bool GetQuiteLoad()
        {
            // Т.к. нужно значение получить до инициализации файла, то читать нужно напрямую
            try
            {
                var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
                if (key != null)
                    using (key)
                    {
                        var cfile = key.GetValue("ConfigFile") as string;
                        if (!string.IsNullOrEmpty(cfile))
                        {
                            var sfile = XElement.Load(cfile);
                            return
                                bool.Parse(
                                    // ReSharper disable once AssignNullToNotNullAttribute
                                    sfile.Element("Settings")?.Element("MainSet")?.Attribute("ChkQuietLoading")?.Value);
                        }
                    }
                return false;
            }
            catch
            {
                return false;
            }
        }
        // проверка соответсвия версии автокада
        private static bool CheckCadVersion()
        {
            var cadVer = AcApp.Version;
            return (cadVer.Major + "." + cadVer.Minor).Equals(MpVersionData.CurCadInternalVersion);
        }

        // Принудительная загрузка сборок
        // необходимых для работы
        private static void LoadAssms(Editor ed)
        {
            try
            {
                var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
                using (key)
                {
                    if (key != null)
                    {
                        var assemblies = key.GetValue("Dll").ToString().Split('/').ToList();
                        foreach (var file in Directory.GetFiles(key.GetValue("TopDir").ToString(), "*.dll", SearchOption.AllDirectories))
                        {
                            if (assemblies.Contains((new FileInfo(file)).Name))
                            {
                                if (!_quiteLoad) ed.WriteMessage("\n* Загрузка компонента: " + (new FileInfo(file)).Name);
                                Assembly.LoadFrom(file);
                            }
                        }

                    }
                }
            }
            catch (System.Exception exception)
            {
                ExceptionBox.ShowForConfigurator(exception);
            }
        }
        // Загрузка базы данных
        private static readonly List<string> BaseFiles = new List<string>
        {
            "mpBaseInt.dll", "mpMetall.dll", "mpConcrete.dll", "mpMaterial.dll", "mpOther.dll", "mpWood.dll", "mpProductInt.dll"
        };

        private static void LoadBaseAssemblies(Editor ed)
        {
            try
            {
                var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
                using (key)
                {
                    if (key != null)
                    {
                        var directory = Path.Combine(key.GetValue("TopDir").ToString(), "Data");
                        if (Directory.Exists(directory))
                        {
                            foreach (var baseFile in BaseFiles)
                            {
                                var file = Path.Combine(directory, baseFile);
                                if (File.Exists(file))
                                {
                                    if (!_quiteLoad) ed.WriteMessage("\n* Загрузка файла базы данных: " + baseFile);
                                    Assembly.LoadFrom(file);
                                }
                                else
                                    if (!_quiteLoad) ed.WriteMessage("\n* Не найден файл базы данных: " + baseFile);
                            }
                        }
                        else
                        {
                            if (!_quiteLoad) ed.WriteMessage("\nНе найдена папка Data!");
                        }
                    }
                }
            }
            catch (System.Exception exception)
            {
                ExceptionBox.ShowForConfigurator(exception);
            }
        }
        // пере/Инициализация файла настроек
        //private static bool InitConfigFile()
        //{
        //    try
        //    {
        //        UserConfigFile.InitConfigFile();
        //        // ReSharper disable once AssignNullToNotNullAttribute
        //        var curDir = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
        //        // ReSharper disable once AssignNullToNotNullAttribute
        //        var configPath = Path.Combine(curDir, "UserData");
        //        if (!Directory.Exists(configPath))
        //            Directory.CreateDirectory(configPath);
        //        // Сначала проверяем путь, указанный в реестре
        //        // Если файл есть - грузим, если нет, то создаем по стандартному пути
        //        var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
        //        if (key != null)
        //            using (key)
        //            {
        //                var cfile = key.GetValue("ConfigFile") as string;
        //                if (string.IsNullOrEmpty(cfile) | !File.Exists(cfile))
        //                {
        //                    MpSettings.LoadFile(Path.Combine(configPath, "mpConfig.mpcf"));
        //                    return false;
        //                }
        //                if (MpSettings.ValidateXml(cfile))
        //                    MpSettings.LoadFile(cfile);
        //                else
        //                {
        //                    CopySettingsFileFromBackUp(cfile);
        //                    MpSettings.LoadFile(cfile);
        //                }
        //                return true;
        //            }
        //        return false;
        //    }
        //    catch (System.Exception exception)
        //    {
        //        ExceptionBox.ShowForConfigurator(exception);
        //        return false;
        //    }
        //}

        //private static void MakeSettingsFileBackUp()
        //{
        //    try
        //    {
        //        var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
        //        if (key != null)
        //            using (key)
        //            {
        //                var cfile = key.GetValue("ConfigFile") as string;
        //                if (cfile != null & File.Exists(cfile))
        //                {
        //                    var fi = new FileInfo(cfile);
        //                    if (fi.DirectoryName != null)
        //                        File.Copy(cfile, Path.Combine(fi.DirectoryName, "mpConfig.backup"), true);
        //                }
        //            }
        //    }
        //    catch (System.Exception exception)
        //    {
        //        ExceptionBox.ShowForConfigurator(exception);
        //    }
        //}

        //private static void CopySettingsFileFromBackUp(string cfile)
        //{
        //    try
        //    {
        //        if (File.Exists(cfile))
        //        {
        //            var fi = new FileInfo(cfile);
        //            if (fi.DirectoryName != null)
        //            {
        //                var bf = Path.Combine(fi.DirectoryName, "mpConfig.backup");
        //                if (File.Exists(bf))
        //                    File.Copy(bf, cfile, true);
        //            }
        //        }
        //    }
        //    catch (System.Exception exception)
        //    {
        //        ExceptionBox.ShowForConfigurator(exception);
        //    }
        //}
        // Загрузка функций

        private static void LoadFunctions(Editor ed)
        {
            try
            {
                // Расположение файла конфигурации
                var confF = UserConfigFile.FullFileName;
                // Грузим
                var configFile = XElement.Load(confF);
                // Делаем итерацию по значениям в файле конфигурации
                var xElement = configFile.Element("Config");
                var el = xElement?.Element("Functions");
                if (el != null)
                    foreach (var conFunc in el.Elements("function"))
                    {
                        /* Так как после обновления добавится значение 
                         * ProductFor, то нужно проверять по нем, при наличии
                         */
                        var productForAttr = conFunc.Attribute("ProductFor");
                        if (productForAttr != null)
                            if (!productForAttr.Value.Equals("AutoCAD"))
                                continue;
                        var confFuncNameAttr = conFunc.Attribute("Name");
                        if (confFuncNameAttr != null)
                        {
                            /* Так как значение AvailCad будет являться устаревшим, НО
                            * пока не будет удалено, делаем двойной вариант проверки
                            */
                            var conFuncAvailCad = string.Empty;
                            var confFuncAvailCadAttr = conFunc.Attribute("AvailCad");
                            if (confFuncAvailCadAttr != null)
                                conFuncAvailCad = confFuncAvailCadAttr.Value;
                            var availProductExternalVersionAttr = conFunc.Attribute("AvailProductExternalVersion");
                            if (availProductExternalVersionAttr != null)
                                conFuncAvailCad = availProductExternalVersionAttr.Value;
                            if (!string.IsNullOrEmpty(conFuncAvailCad))
                            {
                                // Проверяем по версии автокада
                                if (conFuncAvailCad.Equals(MpVersionData.CurCadVers))
                                {
                                    // Добавляем если только функция включена и есть физически на диске!!!
                                    var conFuncOnOff = bool.TryParse(conFunc.Attribute("OnOff")?.Value, out bool b) && b; // false
                                    var conFuncFileAttr = conFunc.Attribute("File");
                                    // Т.к. атрибута File может не быть
                                    if (conFuncOnOff)
                                    {
                                        if (conFuncFileAttr != null)
                                        {
                                            if (File.Exists(conFuncFileAttr.Value))
                                            {
                                                if (!_quiteLoad)
                                                    ed.WriteMessage("\n* Загрузка функции: " + confFuncNameAttr.Value);
                                                var loadedFuncAssembly = Assembly.LoadFrom(conFuncFileAttr.Value);
                                                LoadFunctionsHelper.GetDataFromFunctionIntrface(loadedFuncAssembly);
                                            }
                                        }
                                        else
                                        {
                                            var findedFile = LoadFunctionsHelper.FindFile(confFuncNameAttr.Value);
                                            if (!string.IsNullOrEmpty(findedFile))
                                                if (File.Exists(findedFile))
                                                {
                                                    if (!_quiteLoad)
                                                        ed.WriteMessage("\n* Загрузка функции: " + confFuncNameAttr.Value);
                                                    var loadedFuncAssembly = Assembly.LoadFrom(findedFile);
                                                    LoadFunctionsHelper.GetDataFromFunctionIntrface(loadedFuncAssembly);
                                                }
                                        }
                                    }
                                }
                            }
                        }
                    }
            }
            catch (System.Exception exception)
            {
                ExceptionBox.ShowForConfigurator(exception);
            }
        }
        /// <summary>
        /// Обработчик события, который проверяет, что построилась лента
        /// И когда она построилась - уже грузим свою вкладку, если надо
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void ComponentManager_ItemInitialized(object sender, Autodesk.Windows.RibbonItemEventArgs e)
        {
            //now one Ribbon item is initialized, but the Ribbon control
            //may not be available yet, so check if before
            if (Autodesk.Windows.ComponentManager.Ribbon == null) return;
            //ok, create Ribbon
            if (ModPlusAPI.Variables.Ribbon)
                RibbonBuilder.BuildRibbon();
            else
                RibbonBuilder.RemoveRibbon();
            //and remove the event handler
            Autodesk.Windows.ComponentManager.ItemInitialized -=
                ComponentManager_ItemInitialized;
        }
        /// <summary>
        /// Проверка загруженности модуля автообновления
        /// </summary>
        private static void CheckAutoUpdaterLoaded()
        {
            try
            {
                var  loadWithWindows =!bool.TryParse(ModPlusAPI.Regestry.GetValue("AutoUpdater","LoadWithWindows"), out bool b) || b;
                if (loadWithWindows)
                {
                    // Если "грузить с виндой", то проверяем, что модуль запущен
                    // если не запущен - запускаем
                    var isOpen = Process.GetProcesses().Any(t => t.ProcessName == "mpAutoUpdater");
                    if (!isOpen)
                    {
                        // ReSharper disable once AssignNullToNotNullAttribute
                        var curDir = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
                        if (curDir != null)
                        {
                            var fileToStart = Path.Combine(curDir, "mpAutoUpdater.exe");
                            if (File.Exists(fileToStart))
                            {
                                Process.Start(fileToStart);
                            }
                        }
                    }
                }
            }
            catch
            {
                // ignored
            }
        }
    }
    public class MpCadHelpers
    {
        /// <summary>
        /// Функции получения данных из автокада
        /// </summary>
        public class GetFromAutoCad
        {
            /// <summary>
            /// Получает расстояние между двумя указанными точками в виде строки
            /// </summary>
            /// <returns>Строкове представление расстояния</returns>
            public static string GetLenByTwoPoint()
            {
                try
                {
                    using (AcApp.DocumentManager.MdiActiveDocument.LockDocument())
                    {
                        var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;
                        var pdo = new PromptDistanceOptions("\nПервая точка: ");
                        var pdr = ed.GetDistance(pdo);
                        return pdr.Status != PromptStatus.OK ? string.Empty : pdr.Value.ToString(CultureInfo.InvariantCulture);
                    }
                }
                catch (System.Exception ex)
                {
                    ExceptionBox.ShowForConfigurator(ex);
                    return string.Empty;
                }

            }

            /// <summary>
            /// Получение суммы длин выбранных примитивов
            /// </summary>
            /// <param name="sumLen">Сумма длин всех примитивов</param>
            /// <param name="entities">Поддерживаемые примитивы</param>
            /// <param name="count">Количество примитивов</param>
            /// <param name="lens">Сумма длин для каждого примитива</param>
            /// <param name="objectIds">Список ObjectId выбранных примитивов</param>
            public static void GetLenFromEntities(out double sumLen, out List<string> entities, out List<int> count, out List<double> lens, out List<List<ObjectId>> objectIds)
            {
                // Поддерживаемые примитивы
                entities = new List<string> { "Line", "Circle", "Polyline", "Arc", "Spline", "Ellipse" };
                // Выбранное количество
                count = new List<int> { 0, 0, 0, 0, 0, 0 };
                // Сумма длин
                sumLen = 0.0;
                lens = new List<double> { 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 };
                // Список ObjectId
                objectIds = new List<List<ObjectId>>
                {
                    new List<ObjectId>(),
                    new List<ObjectId>(),
                    new List<ObjectId>(),
                    new List<ObjectId>(),
                    new List<ObjectId>(),
                    new List<ObjectId>()
                };

                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                try
                {
                    var selRes = ed.SelectImplied();
                    // Если сначала ничего не выбрано, просим выбрать:
                    if (selRes.Status == PromptStatus.Error)
                    {
                        var selOpts = new PromptSelectionOptions
                        {
                            MessageForAdding =
                                "\n" + "Выберите отрезки, полилинии, окружности, дуги, эллипсы или сплайны: "
                        };
                        TypedValue[] values =
                            {
                                new TypedValue((int) DxfCode.Operator, "<OR"),
                                new TypedValue((int) DxfCode.Start, "LINE"),
                                new TypedValue((int) DxfCode.Start, "POLYLINE"),
                                new TypedValue((int) DxfCode.Start, "LWPOLYLINE"),
                                new TypedValue((int) DxfCode.Start, "CIRCLE"),
                                new TypedValue((int) DxfCode.Start, "ARC"),
                                new TypedValue((int) DxfCode.Start, "SPLINE"),
                                new TypedValue((int) DxfCode.Start, "ELLIPSE"),
                                new TypedValue((int) DxfCode.Operator, "OR>")
                            };
                        var sfilter = new SelectionFilter(values);
                        selRes = ed.GetSelection(selOpts, sfilter);
                    }
                    else ed.SetImpliedSelection(new ObjectId[0]);

                    if (selRes.Status == PromptStatus.OK)// Если выбраны объекты, тогда дальше
                    {
                        using (var tr = doc.TransactionManager.StartTransaction())
                        {
                            try
                            {
                                var objIds = selRes.Value.GetObjectIds();
                                foreach (var objId in objIds)
                                {
                                    var ent = (Entity)tr.GetObject(objId, OpenMode.ForRead);
                                    switch (ent.GetType().Name)
                                    {
                                        case "Line":
                                            count[0]++;
                                            objectIds[0].Add(objId);
                                            lens[0] += ((Line)ent).Length;
                                            break;
                                        case "Circle":
                                            count[1]++;
                                            objectIds[1].Add(objId);
                                            lens[1] += ((Circle)ent).Circumference;
                                            break;
                                        case "Polyline":
                                            count[2]++;
                                            objectIds[2].Add(objId);
                                            lens[2] += ((Autodesk.AutoCAD.DatabaseServices.Polyline)ent).Length;
                                            break;
                                        case "Arc":
                                            count[3]++;
                                            objectIds[3].Add(objId);
                                            lens[3] += ((Arc)ent).Length;
                                            break;
                                        case "Spline":
                                            count[4]++;
                                            objectIds[4].Add(objId);
                                            lens[4] +=
                                                      (((Curve)ent).GetDistanceAtParameter(((Curve)ent).EndParam) -
                                                       ((Curve)ent).GetDistanceAtParameter(((Curve)ent).StartParam));
                                            break;
                                        case "Ellipse":
                                            count[5]++;
                                            objectIds[5].Add(objId);
                                            lens[5] +=
                                                       (((Curve)ent).GetDistanceAtParameter(((Curve)ent).EndParam) -
                                                        ((Curve)ent).GetDistanceAtParameter(((Curve)ent).StartParam));
                                            break;
                                    }
                                    ent.Dispose();
                                }
                                // Общая сумма длин
                                sumLen += lens.Sum();
                                tr.Commit();
                            }
                            catch (System.Exception ex)
                            {
                                ExceptionBox.ShowForConfigurator(ex);
                                tr.Commit();
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ExceptionBox.ShowForConfigurator(ex);
                }
            }
            /// <summary>
            /// Получение суммы длин выбранных примитивов
            /// </summary>
            /// <param name="sumLen">Сумма длин всех примитивов</param>
            public static void GetLenFromEntities(out double sumLen)
            {
                // Сумма длин
                sumLen = 0.0;
                var lens = new List<double> { 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 };

                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                try
                {
                    var selRes = ed.SelectImplied();
                    // Если сначала ничего не выбрано, просим выбрать:
                    if (selRes.Status == PromptStatus.Error)
                    {
                        var selOpts = new PromptSelectionOptions
                        {
                            MessageForAdding =
                                "\n" + "Выберите отрезки, полилинии, окружности, дуги, эллипсы или сплайны: "
                        };
                        TypedValue[] values =
                            {
                                new TypedValue((int) DxfCode.Operator, "<OR"),
                                new TypedValue((int) DxfCode.Start, "LINE"),
                                new TypedValue((int) DxfCode.Start, "POLYLINE"),
                                new TypedValue((int) DxfCode.Start, "LWPOLYLINE"),
                                new TypedValue((int) DxfCode.Start, "CIRCLE"),
                                new TypedValue((int) DxfCode.Start, "ARC"),
                                new TypedValue((int) DxfCode.Start, "SPLINE"),
                                new TypedValue((int) DxfCode.Start, "ELLIPSE"),
                                new TypedValue((int) DxfCode.Operator, "OR>")
                            };
                        var sfilter = new SelectionFilter(values);
                        selRes = ed.GetSelection(selOpts, sfilter);
                    }
                    else ed.SetImpliedSelection(new ObjectId[0]);

                    if (selRes.Status == PromptStatus.OK)// Если выбраны объекты, тогда дальше
                    {
                        using (var tr = doc.TransactionManager.StartTransaction())
                        {
                            try
                            {
                                var objIds = selRes.Value.GetObjectIds();
                                foreach (var objId in objIds)
                                {
                                    var ent = (Entity)tr.GetObject(objId, OpenMode.ForRead);
                                    switch (ent.GetType().Name)
                                    {
                                        case "Line":
                                            lens[0] += ((Line)ent).Length;
                                            break;
                                        case "Circle":
                                            lens[1] += ((Circle)ent).Circumference;
                                            break;
                                        case "Polyline":
                                            lens[2] += ((Autodesk.AutoCAD.DatabaseServices.Polyline)ent).Length;
                                            break;
                                        case "Arc":
                                            lens[3] += ((Arc)ent).Length;
                                            break;
                                        case "Spline":
                                            lens[4] +=
                                                      (((Curve)ent).GetDistanceAtParameter(((Curve)ent).EndParam) -
                                                       ((Curve)ent).GetDistanceAtParameter(((Curve)ent).StartParam));
                                            break;
                                        case "Ellipse":
                                            lens[5] +=
                                                       (((Curve)ent).GetDistanceAtParameter(((Curve)ent).EndParam) -
                                                        ((Curve)ent).GetDistanceAtParameter(((Curve)ent).StartParam));
                                            break;
                                    }
                                    ent.Dispose();
                                }
                                // Общая сумма длин
                                sumLen += lens.Sum();
                                tr.Commit();
                            }
                            catch (System.Exception ex)
                            {
                                ExceptionBox.ShowForConfigurator(ex);
                                tr.Commit();
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ExceptionBox.ShowForConfigurator(ex);
                }
            }
        }
        /// <summary>
        /// Добавление 
        /// </summary>
        /// <param name="regAppName"></param>
        public static void AddRegAppTableRecord(string regAppName)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var tr = doc.TransactionManager.StartTransaction();
            using (tr)
            {
                var rat =
                  (RegAppTable)tr.GetObject(
                    db.RegAppTableId,
                    OpenMode.ForRead,
                    false
                  );
                if (!rat.Has(regAppName))
                {
                    rat.UpgradeOpen();
                    var ratr =
                      new RegAppTableRecord { Name = regAppName };
                    rat.Add(ratr);
                    tr.AddNewlyCreatedDBObject(ratr, true);
                }
                tr.Commit();
            }
        }
        /// <summary>
        /// Добавление текстовых расширенных данных в словарь
        /// </summary>
        /// <param name="dictionaryName">Словарь</param>
        /// <param name="value">Добавляемое значение</param>
        public static void SetStringXData(string dictionaryName, string value)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                var db = doc.Database;
                try
                {
                    using (doc.LockDocument())
                    {
                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var rec = new Xrecord
                            {
                                Data = new ResultBuffer(
                                    new TypedValue(Convert.ToInt32(DxfCode.Text), value))
                            };

                            var dict =
                                (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite, false);
                            dict.SetAt(dictionaryName, rec);
                            tr.AddNewlyCreatedDBObject(rec, true);
                            tr.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ExceptionBox.ShowForConfigurator(ex);
                }
            }
        }
        /// <summary>
        /// Получение текстовых расширенных данных из указанного словаря
        /// </summary>
        /// <param name="dictionaryName">Словарь</param>
        /// <returns>Значение</returns>
        public static string GetStringXData(string dictionaryName)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                var db = doc.Database;
                try
                {
                    using (doc.LockDocument())
                    {
                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var dict =
                                (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite, true);
                            var id = dict.GetAt(dictionaryName);
                            var rec = tr.GetObject(id, OpenMode.ForWrite, true) as Xrecord;
                            var value = string.Empty;
                            if (rec != null)
                                foreach (var rb in rec.Data.AsArray())
                                {
                                    value = rb.Value.ToString();
                                }

                            tr.Commit();
                            return value;
                        }
                    }
                } // try
                catch
                {
                    return string.Empty;
                }
            }
            return string.Empty;
        }
        /// <summary>
        /// Проверка наличия словаря расширенных данных
        /// </summary>
        /// <param name="dictionaryName">Словарь</param>
        /// <returns>True - словарь существует, False - словарь отсутствует</returns>
        public static bool HasXDataDictionary(string dictionaryName)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                var db = doc.Database;
                try
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        using (doc.LockDocument())
                        {
                            var dict =
                                (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite, true);
                            if (dict.Contains(dictionaryName))
                            {
                                tr.Commit();
                                return true;
                            }
                            tr.Commit();
                            return false;
                        }
                    }
                } // try
                catch (System.Exception ex)
                {
                    ExceptionBox.ShowForConfigurator(ex);
                    return false;
                }
            }
            return false;
        }
        /// <summary>
        /// Удаление 
        /// </summary>
        /// <param name="value"></param>
        public static void DeleteStringXData(string value)
        {
            var database = AcApp.DocumentManager.MdiActiveDocument.Database;
            try
            {
                using (AcApp.DocumentManager.MdiActiveDocument.LockDocument())
                {
                    using (var tr = database.TransactionManager.StartTransaction())
                    {
                        var newValue = new Xrecord();
                        var values = new[] { new TypedValue(Convert.ToInt32(DxfCode.XRefPath), value) };
                        newValue.Data = new ResultBuffer(values);
                        var dictionary = ((DBDictionary)tr.GetObject(database.NamedObjectsDictionaryId, OpenMode.ForWrite, false));
                        foreach (var obj in dictionary)
                        {
                            if (obj.Value.GetObject(OpenMode.ForRead) is Xrecord)
                            {
                                var rec = obj.Value.GetObject(OpenMode.ForRead) as Xrecord;
                                var rb = rec?.Data;
                                var tv = rb?.AsArray();
                                var rb2 = newValue.Data;
                                var tv2 = rb2.AsArray();
                                if (((TypedValue)tv.GetValue(0)).Value.Equals(
                                    ((TypedValue)tv2.GetValue(0)).Value))
                                {
                                    dictionary.Remove(obj.Key);
                                    break;
                                }
                            }
                        }
                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ExceptionBox.ShowForConfigurator(ex);
            }
        }
    }
    public static class XDataHelpersForProducts
    {
        private const string AppName = "ModPlusProduct";

        public static void SaveDataToEntity(object product, DBObject ent, Transaction tr)
        {
            var regTable = (RegAppTable)tr.GetObject(ent.Database.RegAppTableId, OpenMode.ForWrite);
            if (!regTable.Has(AppName))
            {
                var app = new RegAppTableRecord
                {
                    Name = AppName
                };
                regTable.Add(app);
                tr.AddNewlyCreatedDBObject(app, true);
            }

            using (var resBuf = SaveToResBuf(product))
            {
                ent.XData = resBuf;
            }
        }
        public static object NewFromEntity(Entity ent)
        {
            using (var resBuf = ent.GetXDataForApplication(AppName))
            {
                return resBuf == null ? null : NewFromResBuf(resBuf);
            }
        }
        private static object NewFromResBuf(ResultBuffer resBuf)
        {
            var bf = new BinaryFormatter { Binder = new MyBinder() };

            var ms = MyUtil.ResBufToStream(resBuf);

            var mbc = bf.Deserialize(ms);

            return mbc;
        }
        private static ResultBuffer SaveToResBuf(object product)
        {
            var bf = new BinaryFormatter();
            var ms = new MemoryStream();
            bf.Serialize(ms, product);
            ms.Position = 0;

            var resBuf = MyUtil.StreamToResBuf(ms, AppName);

            return resBuf;
        }
        sealed class MyBinder : SerializationBinder
        {
            public override Type BindToType(
              string assemblyName,
              string typeName)
            {
                return Type.GetType($"{typeName}, {assemblyName}");
            }
        }
        class MyUtil
        {
            const int KMaxChunkSize = 127;

            public static ResultBuffer StreamToResBuf(
              Stream ms, string appName)
            {
                var resBuf = new ResultBuffer(
                    new TypedValue(
                        (int)DxfCode.ExtendedDataRegAppName, appName));
                for (var i = 0; i < ms.Length; i += KMaxChunkSize)
                {
                    var length = (int)Math.Min(ms.Length - i, KMaxChunkSize);
                    var datachunk = new byte[length];
                    ms.Read(datachunk, 0, length);
                    resBuf.Add(
                      new TypedValue(
                        (int)DxfCode.ExtendedDataBinaryChunk, datachunk));
                }

                return resBuf;
            }

            public static MemoryStream ResBufToStream(ResultBuffer resBuf)
            {
                var ms = new MemoryStream();
                var values = resBuf.AsArray();

                // Start from 1 to skip application name

                for (var i = 1; i < values.Length; i++)
                {
                    var datachunk = (byte[])values[i].Value;
                    ms.Write(datachunk, 0, datachunk.Length);
                }
                ms.Position = 0;

                return ms;
            }
        }
    }
    // Работа с палитрой
    public static class MpPalette
    {
        public static PaletteSet MpPaletteSet;
        [CommandMethod("mpPalette")]
        public static void CreatePalette()
        {
            try
            {
                if (MpPaletteSet == null)
                {
                    MpPaletteSet = new PaletteSet("mpPalette", "mpPalette", new Guid("A9C907EF-6281-4FA2-9B6C-E0401E41BB74"));
                    MpPaletteSet.Load += _mpPaletteSet_Load;
                    MpPaletteSet.Save += _mpPaletteSet_Save;
                    AddRemovePaletts();
                    MpPaletteSet.Icon = GetEmbeddedIcon("ModPlus.Resources.mpIcon.ico");
                    MpPaletteSet.Style =
                        PaletteSetStyles.ShowPropertiesMenu |
                        PaletteSetStyles.ShowAutoHideButton |
                        PaletteSetStyles.ShowCloseButton;
                    MpPaletteSet.MinimumSize = new System.Drawing.Size(100, 300);
                    MpPaletteSet.DockEnabled = DockSides.Left | DockSides.Right;
                    MpPaletteSet.RecalculateDockSiteLayout();
                    MpPaletteSet.Visible = true;
                }
                else
                {
                    AddRemovePaletts();
                    MpPaletteSet.Visible = true;
                }
            }
            catch (System.Exception exception) { ExceptionBox.ShowForConfigurator(exception); }
        }

        private static void AddRemovePaletts()
        {
            // functions
            if (ModPlusAPI.Variables.FunctionsInPalette)
            {
                var hasP = false;
                foreach (Palette p in MpPaletteSet)
                {
                    if (p.Name.Equals("Функции")) hasP = true;
                }
                if (!hasP)
                {
                    var palette = new mpPaletteFunctions();
                    var host = new ElementHost
                    {
                        AutoSize = true,
                        Dock = System.Windows.Forms.DockStyle.Fill,
                        Child = palette
                    };
                    MpPaletteSet.Add("Функции", host);
                }
            }
            else
            {
                for (var i = 0; i < MpPaletteSet.Count; i++)
                {
                    if (MpPaletteSet[i].Name.Equals("Функции"))
                    {
                        MpPaletteSet.Remove(i);
                        break;
                    }
                }
            }
            // drawings
            if (ModPlusAPI.Variables.DrawingsInPalette)
            {
                var hasP = false;
                foreach (Palette p in MpPaletteSet)
                {
                    if (p.Name.Equals("Чертежи")) hasP = true;
                }
                if (!hasP)
                {
                    var palette = new mpPaletteDrawings();
                    var host = new ElementHost
                    {
                        AutoSize = true,
                        Dock = System.Windows.Forms.DockStyle.Fill,
                        Child = palette
                    };
                    MpPaletteSet.Add("Чертежи", host);
                }
            }
            else
            {
                for (var i = 0; i < MpPaletteSet.Count; i++)
                {
                    if (MpPaletteSet[i].Name.Equals("Чертежи"))
                    {
                        MpPaletteSet.Remove(i);
                        break;
                    }
                }
            }
        }

        private static void _mpPaletteSet_Save(object sender, PalettePersistEventArgs e)
        {
            var a = (double)e.ConfigurationSection.ReadProperty("ModPlusPalette", 22.3);
        }

        private static void _mpPaletteSet_Load(object sender, PalettePersistEventArgs e)
        {
            e.ConfigurationSection.WriteProperty("ModPlusPalette", 32.3);
        }

        private static Icon GetEmbeddedIcon(string sName)
        {
            // ReSharper disable once AssignNullToNotNullAttribute
            return new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream(sName));
        }
    }
}

