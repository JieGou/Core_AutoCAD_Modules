#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Xml.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
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
        private static bool _quiteLoad;
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
                ModPlusAPI.Statistic.SendPluginStarting("AutoCAD", MpVersionData.CurCadVers);
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
                ExceptionBox.Show(exception);
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
                ExceptionBox.Show(exception);
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
                ExceptionBox.Show(exception);
            }
        }
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
                ExceptionBox.Show(exception);
            }
        }
        /// <summary>
        /// Обработчик события, который проверяет, что построилась лента
        /// И когда она построилась - уже грузим свою вкладку, если надо
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void ComponentManager_ItemInitialized(object sender, Autodesk.Windows.RibbonItemEventArgs e)
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
                var  loadWithWindows =!bool.TryParse(Regestry.GetValue("AutoUpdater","LoadWithWindows"), out bool b) || b;
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
            catch (System.Exception exception)
            {
                Statistic.SendException(exception);
            }
        }
    }
    /// <summary>Вспомгательные методы работы с расширенными данными для функций из раздела "Продукты ModPlus"</summary>
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
    /// <summary>Методы создания и работы с палитрой ModPlus</summary>
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
                    MpPaletteSet = new PaletteSet("Палитра ModPlus", "mpPalette", new Guid("A9C907EF-6281-4FA2-9B6C-E0401E41BB75"));
                    MpPaletteSet.Load += _mpPaletteSet_Load;
                    MpPaletteSet.Save += _mpPaletteSet_Save;
                    AddRemovePaletts();
                    MpPaletteSet.Icon = GetEmbeddedIcon("ModPlus.Resources.mpIcon.ico");
                    MpPaletteSet.Style =
                        PaletteSetStyles.ShowPropertiesMenu |
                        PaletteSetStyles.ShowAutoHideButton |
                        PaletteSetStyles.ShowCloseButton;
                    MpPaletteSet.MinimumSize = new Size(100, 300);
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
            catch (System.Exception exception) { ExceptionBox.Show(exception); }
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
            // ReSharper disable once UnusedVariable
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

