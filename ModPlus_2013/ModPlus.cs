using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
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

namespace ModPlus
{
    public class ModPlus : IExtensionApplication
    {
        private const string LangItem = "AutocadDlls";

        private static bool _quiteLoad;
        // Инициализация плагина
        public void Initialize()
        {
            try
            {
                var sw = new Stopwatch();
                sw.Start();
                // inint lang
                if(!Language.Initialize()) return;
                // Получим значение переменной "Тихая загрузка" в первую очередь
                _quiteLoad = ModPlusAPI.Variables.QuietLoading;
                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                // Файла конфигурации может не существовать при загрузке плагина!
                // Поэтому все, что связанно с работой с файлом конфигурации должно это учитывать!
                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;
                if (!CheckCadVersion())
                {
                    ed.WriteMessage("\n***************************");
                    ed.WriteMessage("\n" + Language.GetItem(LangItem, "p1"));
                    ed.WriteMessage("\n" + Language.GetItem(LangItem, "p2"));
                    ed.WriteMessage("\n" + Language.GetItem(LangItem, "p3"));
                    ed.WriteMessage("\n***************************");
                    return;
                }
                Statistic.SendPluginStarting("AutoCAD", MpVersionData.CurCadVers);
                ed.WriteMessage("\n***************************");
                ed.WriteMessage("\n" + Language.GetItem(LangItem, "p4"));
                if (!_quiteLoad) ed.WriteMessage("\n" + Language.GetItem(LangItem, "p5"));
                // Принудительная загрузка сборок
                LoadAssms(ed);
                if (!_quiteLoad) ed.WriteMessage("\n" + Language.GetItem(LangItem, "p6"));
                LoadBaseAssemblies(ed);
                if (!_quiteLoad) ed.WriteMessage("\n" + Language.GetItem(LangItem, "p7"));
                UserConfigFile.InitConfigFile();
                if (!_quiteLoad) ed.WriteMessage("\n" + Language.GetItem(LangItem, "p8"));
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
                // Включение иконок для продуктов
                var showProductsIcon = bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings,
                    "mpProductInsert", "ShowIcon"), out var b) && b; //false
                if (showProductsIcon)
                    MpProductIconFunctions.ShowIcon();

                sw.Stop();
                ed.WriteMessage("\n" + Language.GetItem(LangItem, "p9") + " " + sw.ElapsedMilliseconds);
                ed.WriteMessage("\n" + Language.GetItem(LangItem, "p10"));
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
                foreach (var fileName in Constants.ExtensionsLibraries)
                {
                    var extDll = Path.Combine(Constants.ExtensionsDirectory, fileName);
                    if (File.Exists(extDll))
                    {
                        if (!_quiteLoad) ed.WriteMessage("\n* " + Language.GetItem(LangItem, "p11") + " " + fileName);
                        Assembly.LoadFrom(extDll);
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
                                    if (!_quiteLoad) ed.WriteMessage("\n* " + Language.GetItem(LangItem, "p12") + " " + baseFile);
                                    Assembly.LoadFrom(file);
                                }
                                else
                                    if (!_quiteLoad) ed.WriteMessage("\n* " + Language.GetItem(LangItem, "p13") + " " + baseFile);
                            }
                        }
                        else
                        {
                            if (!_quiteLoad) ed.WriteMessage("\n" + Language.GetItem(LangItem, "p14"));
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
                var funtionsKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus\\Functions");
                if (funtionsKey == null) return;
                using (funtionsKey)
                {
                    foreach (var functionKeyName in funtionsKey.GetSubKeyNames())
                    {
                        var functionKey = funtionsKey.OpenSubKey(functionKeyName);
                        if (functionKey == null) continue;
                        foreach (var availPrVersKeyName in functionKey.GetSubKeyNames())
                        {
                            // Если версия продукта не совпадает, то пропускаю
                            if (!availPrVersKeyName.Equals(MpVersionData.CurCadVers)) continue;
                            var availPrVersKey = functionKey.OpenSubKey(availPrVersKeyName);
                            if (availPrVersKey == null) continue;
                            // беру свойства функции из реестра
                            var file = availPrVersKey.GetValue("File") as string;
                            var onOff = availPrVersKey.GetValue("OnOff") as string;
                            var productFor = availPrVersKey.GetValue("ProductFor") as string;
                            if (string.IsNullOrEmpty(onOff) || string.IsNullOrEmpty(productFor)) continue;
                            if (!productFor.Equals("AutoCAD")) continue;
                            var isOn = !bool.TryParse(onOff, out var b) || b; // default - true
                            // Если "Продукт для" подходит, файл существует и функция включена - гружу
                            if (isOn)
                            {
                                if (!string.IsNullOrEmpty(file) && File.Exists(file))
                                {
                                    // load
                                    if (!_quiteLoad)
                                        ed.WriteMessage("\n* " + Language.GetItem(LangItem, "p15") + " " + functionKeyName);
                                    var localFuncAssembly = Assembly.LoadFrom(file);
                                    LoadFunctionsHelper.GetDataFromFunctionIntrface(localFuncAssembly);
                                }
                                else
                                {
                                    var findedFile = LoadFunctionsHelper.FindFile(functionKeyName);
                                    if (!string.IsNullOrEmpty(findedFile) && File.Exists(findedFile))
                                    {
                                        if (!_quiteLoad)
                                            ed.WriteMessage("\n* " + Language.GetItem(LangItem, "p15") + " " + functionKeyName);
                                        var localFuncAssembly = Assembly.LoadFrom(findedFile);
                                        LoadFunctionsHelper.GetDataFromFunctionIntrface(localFuncAssembly);
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
                var loadWithWindows = !bool.TryParse(Regestry.GetValue("AutoUpdater", "LoadWithWindows"), out bool b) || b;
                if (loadWithWindows)
                {
                    // Если "грузить с виндой", то проверяем, что модуль запущен
                    // если не запущен - запускаем
                    var isOpen = Process.GetProcesses().Any(t => t.ProcessName == "mpAutoUpdater");
                    if (!isOpen)
                    {
                            var fileToStart = Path.Combine(Constants.CurrentDirectory, "mpAutoUpdater.exe");
                            if (File.Exists(fileToStart))
                            {
                                Process.Start(fileToStart);
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

        public static bool IsModPlusProduct(this Entity ent)
        {
            using (var rb = ent.GetXDataForApplication(AppName))
            {
                return rb != null;
            }
        }
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
        private const string LangItem = "AutocadDlls";
        public static PaletteSet MpPaletteSet;
        [CommandMethod("mpPalette")]
        public static void CreatePalette()
        {
            try
            {
                if (MpPaletteSet == null)
                {
                    MpPaletteSet = new PaletteSet(Language.GetItem(LangItem, "h48"), "mpPalette", new Guid("A9C907EF-6281-4FA2-9B6C-E0401E41BB76"));
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
                    MpPaletteSet.Visible = true;
                }
            }
            catch (System.Exception exception) { ExceptionBox.Show(exception); }
        }
        
        private static void AddRemovePaletts()
        {
            if(MpPaletteSet == null) return;
            try
            {
                var funName = Language.GetItem(LangItem, "h19");
                var drwName = Language.GetItem(LangItem, "h20");
                // functions
                if (ModPlusAPI.Variables.FunctionsInPalette)
                {
                    var hasP = false;
                    foreach (Palette p in MpPaletteSet)
                    {
                        if (p.Name.Equals(funName)) hasP = true;
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
                        MpPaletteSet.Add(funName, host);
                    }
                }
                else
                {
                    for (var i = 0; i < MpPaletteSet.Count; i++)
                    {
                        if (MpPaletteSet[i].Name.Equals(funName))
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
                        if (p.Name.Equals(drwName)) hasP = true;
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
                        MpPaletteSet.Add(drwName, host);
                    }
                }
                else
                {
                    for (var i = 0; i < MpPaletteSet.Count; i++)
                    {
                        if (MpPaletteSet[i].Name.Equals(drwName))
                        {
                            MpPaletteSet.Remove(i);
                            break;
                        }
                    }
                }
            }
            catch
            {
                // ignore
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

