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
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using ModPlus.App;
using mpMsg;
using mpRegistry;
using mpSettings;
using MahApps.Metro;
using MahApps.Metro.Controls;
using ModPlus.Windows;
using System.Windows.Forms.Integration;
using mpPInterface;
using ModPlus.Helpers;
using Brushes = System.Windows.Media.Brushes;

namespace ModPlus
{
    public class ModPlus : IExtensionApplication
    {
        // ReSharper disable once RedundantDefaultMemberInitializer
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
                if (!_quiteLoad)
                    ed.WriteMessage(InitConfigFile() ? " OK" : " Файл не найден. Создание стандартного");
                else InitConfigFile();

                if (!_quiteLoad) ed.WriteMessage("\nИнициализация глобальных переменных...");
                MpVars.ReadVarsFromSettingsfile();
                if (!_quiteLoad) ed.WriteMessage("\nЗагрузка функций...");
                LoadFunctions(ed);
                // Строим: ленту, меню, плавающее меню
                // Загрузка ленты
                Autodesk.Windows.ComponentManager.ItemInitialized += ComponentManager_ItemInitialized;
                // Палитра
                if (MpVars.MpPalette)
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
                MpExWin.Show(exception);
            }
        }
        public void Terminate()
        {
            // Делаем копию файла настроек
            MakeSettingsFileBackUp();
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
        static void LoadAssms(Editor ed)
        {
            try
            {
                var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
                using (key)
                {
                    if (key != null)
                    {
                        var assemblies = key.GetValue("Dll").ToString().Split('/').ToList();
                        if(!assemblies.Contains("mpCustomThemes.dll"))
                            assemblies.Add("mpCustomThemes.dll");
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
                MpExWin.Show(exception);
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
                MpExWin.Show(exception);
            }
        }
        // пере/Инициализация файла настроек
        private static bool InitConfigFile()
        {
            try
            {
                // ReSharper disable once AssignNullToNotNullAttribute
                var curDir = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
                // ReSharper disable once AssignNullToNotNullAttribute
                var configPath = Path.Combine(curDir, "UserData");
                if (!Directory.Exists(configPath))
                    Directory.CreateDirectory(configPath);
                // Сначала проверяем путь, указанный в реестре
                // Если файл есть - грузим, если нет, то создаем по стандартному пути
                var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
                if (key != null)
                    using (key)
                    {
                        var cfile = key.GetValue("ConfigFile") as string;
                        if (string.IsNullOrEmpty(cfile) | !File.Exists(cfile))
                        {
                            MpSettings.LoadFile(Path.Combine(configPath, "mpConfig.mpcf"));
                            return false;
                        }
                        if (MpSettings.ValidateXml(cfile))
                            MpSettings.LoadFile(cfile);
                        else
                        {
                            CopySettingsFileFromBackUp(cfile);
                            MpSettings.LoadFile(cfile);
                        }
                        return true;
                    }
                return false;
            }
            catch (System.Exception exception)
            {
                MpExWin.Show(exception);
                return false;
            }
        }

        private static void MakeSettingsFileBackUp()
        {
            try
            {
                var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
                if (key != null)
                    using (key)
                    {
                        var cfile = key.GetValue("ConfigFile") as string;
                        if (cfile != null & File.Exists(cfile))
                        {
                            var fi = new FileInfo(cfile);
                            if (fi.DirectoryName != null)
                                File.Copy(cfile, Path.Combine(fi.DirectoryName, "mpConfig.backup"), true);
                        }
                    }
            }
            catch (System.Exception exception)
            {
                MpExWin.Show(exception);
            }
        }

        private static void CopySettingsFileFromBackUp(string cfile)
        {
            try
            {
                if (File.Exists(cfile))
                {
                    var fi = new FileInfo(cfile);
                    if (fi.DirectoryName != null)
                    {
                        var bf = Path.Combine(fi.DirectoryName, "mpConfig.backup");
                        if (File.Exists(bf))
                            File.Copy(bf, cfile, true);
                    }
                }
            }
            catch (System.Exception exception)
            {
                MpExWin.Show(exception);
            }
        }
        // Загрузка функций
        private static void LoadFunctions(Editor ed)
        {
            try
            {
                // Расположение файла конфигурации
                var confF = MpSettings.FullFileName;
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
                MpExWin.Show(exception);
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
            if (MpVars.MpRibbon)
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
                var loadWithWindows = true;
                // проверяем в реестре
                var registryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
                if (registryKey != null)
                    using (registryKey)
                    {
                        var autoUpdaterFolder = registryKey.OpenSubKey("AutoUpdater");
                        if (autoUpdaterFolder != null)
                            using (autoUpdaterFolder)
                            {
                                bool b;
                                loadWithWindows =
                                    !bool.TryParse(autoUpdaterFolder.GetValue("LoadWithWindows").ToString(), out b) || b;
                            }
                    }
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

    public class MpWindowHelpers
    {
        public static void OnWindowStartUp(MetroWindow window, string theme, string accentColor, string borderType)
        {
            // Загрузка тем
            ThemeManager.AddAppTheme("DarkBlue", new Uri("pack://application:,,,/mpCustomThemes;component/DarkBlue.xaml"));
            ChangeWindowTheme(theme, accentColor, window);
            ChangeWindowBordes(borderType, window);
        }
        // Вид границ окна
        public static void ChangeWindowBordes(string borderType, MetroWindow window)
        {
            try
            {
                if (borderType.Equals("AccentBorder"))
                {
                    window.BorderThickness = new Thickness(1);
                    window.SetResourceReference(Control.BorderBrushProperty, "AccentColorBrush");
                    window.GlowBrush = null;
                }
                if (borderType.Equals("NoBorder"))
                {
                    window.BorderThickness = new Thickness(0);
                    window.BorderBrush = null;
                    window.GlowBrush = null;
                }
                if (borderType.Equals("AccentGlowBrush"))
                {
                    window.BorderThickness = new Thickness(0);
                    window.SetResourceReference(MetroWindow.GlowBrushProperty, "AccentColorBrush");
                    window.BorderBrush = null;
                }
                if (borderType.Equals("AccentBorderGlowBrush"))
                {
                    window.BorderThickness = new Thickness(1);
                    window.SetResourceReference(MetroWindow.GlowBrushProperty, "AccentColorBrush");
                    window.BorderBrush = null;
                }
                if (borderType.Equals("Shadow") || string.IsNullOrEmpty(borderType))
                {
                    window.BorderThickness = new Thickness(0);
                    window.GlowBrush = Brushes.Black;
                    window.BorderBrush = null;
                }
                if (string.IsNullOrEmpty(borderType))
                {
                    window.BorderThickness = new Thickness(1);
                    window.SetResourceReference(Control.BorderBrushProperty, "AccentColorBrush");
                    window.GlowBrush = null;
                }
            }
            catch
            {
                window.BorderThickness = new Thickness(1);
                window.SetResourceReference(Control.BorderBrushProperty, "AccentColorBrush");
                window.GlowBrush = null;
            }
        }

        public static void ChangeWindowTheme(string theme, string accentColor, System.Windows.Window window)
        {
            //Theme
            try
            {
                ThemeManager.ChangeAppStyle(
                    window,
                    ThemeManager.Accents.First(x => x.Name.Equals(accentColor)),
                    ThemeManager.GetAppTheme(theme)
                    );
            }
            catch
            {
                ThemeManager.ChangeAppStyle(
                    window,
                    ThemeManager.Accents.First(x => x.Name.Equals("Blue")),
                    ThemeManager.AppThemes.First(x => x.Name.Equals("BaseLight"))
                    );
            }
        }
    }
    public class MpCadHelpers
    {
        // Проверка того, что функция купленна
        public static bool IsFunctionBought(string name, string availCad)
        {
            try
            {
                // Расположение файла конфигурации
                var confF = MpSettings.FullFileName;
                // Грузим
                var configFile = XElement.Load(confF);
                // Проверяем есть ли группа Config
                // Если нет, то false
                if (configFile.Element("Config") == null)
                    return false;
                var element = configFile.Element("Config");
                // Проверяем есть ли подгруппа Functions
                // Если нет, то false
                if (element != null && element.Element("Functions") == null)
                    return false;
                var confFuncsXel = element?.Element("Functions");
                // Проходим по функциям в файле
                if (confFuncsXel != null)
                    foreach (var func in confFuncsXel.Elements("function"))
                    {
                        var nameAttr = func.Attribute("Name");
                        if (nameAttr != null && nameAttr.Value.Equals(name))
                        {
                            var availVersion = string.Empty;
                            var availCadAttr = func.Attribute("AvailCad");
                            if (availCadAttr != null)
                                availVersion = availCadAttr.Value;
                            var availProdAttr = func.Attribute("AvailProductExternalVersion");
                            if (availProdAttr != null)
                                availVersion = availProdAttr.Value;
                            if (!string.IsNullOrEmpty(availVersion) & availVersion.Equals(availCad))
                            {
                                var activeKey = func.Attribute("ActiveKey")?.Value;
                                if (string.IsNullOrEmpty(activeKey)) return false;
                                if (activeKey == EncDec.MDString(availCad + MpVars.RegistryKey + name)) return true;
                                return false;
                            }
                        }
                    }
                return false;
            }
            catch (System.Exception exception)
            {
                MpExWin.Show(exception);
                return false;
            }
        }
        /// <summary>
        /// Получаем блок для стрелки
        /// </summary>
        /// <param name="newArrName">Название стрелки</param>
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
                var bt =
                    (BlockTable)tr.GetObject(
                    db.BlockTableId,
                    OpenMode.ForRead
                  );
                arrObjId = bt[newArrName];
                tr.Commit();
            }
            return arrObjId;
        }
        /// <summary>
        // Функция перевода точки из пользовательской
        // системы координат в мировую
        // (взята у Александра Ривилиса)
        /// </summary>
        public static Point3d UcsToWcs(Point3d pt)
        {
            var m = GetUcsMatrix(HostApplicationServices.WorkingDatabase);
            return pt.TransformBy(m);
        }
        public static bool IsPaperSpace(Database db)
        {
            if (db.TileMode) return false;
            var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;
            return db.PaperSpaceVportId == ed.CurrentViewportObjectId;
        }
        public static Matrix3d GetUcsMatrix(Database db)
        {
            Point3d origin;
            Vector3d xAxis, yAxis;
            if (IsPaperSpace(db))
            {
                origin = db.Pucsorg; xAxis = db.Pucsxdir; yAxis = db.Pucsydir;
            }
            else
            {
                origin = db.Ucsorg; xAxis = db.Ucsxdir; yAxis = db.Ucsydir;
            }
            var zAxis = xAxis.CrossProduct(yAxis);
            return Matrix3d.AlignCoordinateSystem(
              Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
              origin, xAxis, yAxis, zAxis);
        }
        /// <summary>
        /// Проверка - является ли строка числом.
        /// С учетом разделителя целой и дробной части
        /// </summary>
        /// <param name="text">Проверяемая строка</param>
        /// <param name="separator">Разделитель целой и дробной части</param>
        /// <returns>True - строка является числом. False - строка не является числом</returns>
        public static bool IsDigit(string text, string separator)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(text, "^[0-9" + separator + "]*$");
        }
        /// <summary>
        /// Функция разбивки строки на список с учетом открывающего и закрывающего символов
        /// </summary>
        /// <param name="str">Входная строка</param>
        /// <param name="symbol1">Первый символ (открывающий)</param>
        /// <param name="symbol2">Второй символ (закрывающий)</param>
        /// <returns>Список</returns>
        public static List<string> BreakString(string str, char symbol1, char symbol2)
        {
            var result = new List<string>();
            var k = -1;
            var sb = new StringBuilder();
            for (var i = 0; i < str.Length; i++)
            {
                if (str[i].Equals(symbol1))
                {
                    if (sb.Length > 0)
                        result.Insert(k, sb.ToString());
                    sb = new StringBuilder();
                    if (i > 1)
                        if (!str[i - 1].Equals(symbol2))
                            k++;
                }
                else if (str[i].Equals(symbol2))
                {
                    result.Insert(k, sb.ToString());
                    sb = new StringBuilder();
                    k++;
                }
                else
                {
                    if (k == -1)
                        k++;
                    sb.Append(str[i]);
                }
            }
            return result;
        }
        /// <summary>
        /// Показать текст в блокноте (без сохранения на диск)
        /// </summary>
        /// <param name="text">Отображамый текст (должен быть со всеми управляющими символами типа \n, \r)</param>
        public static void ShowTextWithNotepad(string text)
        {
            var process = new Process { StartInfo = { FileName = @"notepad.exe" }, EnableRaisingEvents = true };
            process.Start(); // It will start Notepad process
            process.WaitForInputIdle(10000);
            if (process.Responding) // If currently started process(notepad) is responding
            {
                System.Windows.Forms.SendKeys.SendWait(text);
                // It will Add all the text from text variable to notepad 
            }
        }
        /// <summary>
        /// Список в строку с определенным раздилителем
        /// </summary>
        /// <param name="list">Список</param>
        /// <param name="separator">Разделитель. Должен быть один знак!</param>
        public static string ListToStringWithSeparator(List<string> list, string separator)
        {
            var str = list.Aggregate(string.Empty, (current, listitem) => current + listitem + separator);
            str = str.Substring(0, str.Length - 1);
            return str;
        }
        /// <summary>
        /// Зуммировать объекты
        /// </summary>
        /// <param name="objIds">ObjectId зуммируемых объектов</param>
        static public void ZoomToEntity(ObjectId[] objIds)
        {
            Editor editor = AcApp.DocumentManager.MdiActiveDocument.Editor;

            PromptSelectionResult psr = editor.SelectImplied();
            ObjectId[] selected = null;
            if (psr.Status == PromptStatus.OK)
                selected = psr.Value.GetObjectIds();
            editor.SetImpliedSelection(objIds);

            Autodesk.AutoCAD.Internal.Utils.ZoomObjects(true);

            editor.SetImpliedSelection(selected);
        }
        /// <summary>
        /// Замена разделителя в строке. Сначала заменяются запятые на точки, затем точки на текущий разделитель
        /// </summary>
        /// <param name="str">Искомая строка</param>
        /// <returns></returns>
        public static string ReplaceSeparator(string str)
        {
            return str.Replace(',', '.').Replace('.', Convert.ToChar(MpVars.MpSeparator));
        }

        /// <summary>
        /// Функции вставки/добавления в автокад (в примитивы)
        /// </summary>
        public class InsertToAutoCad
        {
            /// <summary>
            /// Вставить строкове значение в ячейку таблицы автокада.
            /// </summary>
            /// <param name="firstStr">Первая строка. Замена разделителя действует только для неё</param>
            /// <param name="secondString">Вторая строка. Необязательно. Замена разделителя не действует</param>
            /// <param name="useSeparator">Использовать разделитель (для цифровых значений).
            /// Заменяет запятую на точку, а затем на текущий разделитель</param>
            public static void AddStrToAutoCadTableCell(string firstStr, string secondString, bool useSeparator)
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
                                                cell.TextString = firstStr.Replace(',', '.').
                                                                      Replace('.', Convert.ToChar(MpVars.MpSeparator)) +
                                                                  secondString;
                                            else cell.TextString = firstStr + secondString;
                                            end = true;
                                        }
                                    } // try
                                    catch
                                    {
                                        MpMsgWin.Show("Не попали в ячейку!");
                                    }
                                } // while
                                tr.Commit();
                            } //try
                            catch (System.Exception ex)
                            {
                                MpExWin.Show(ex);
                            }
                        } //using tr
                    } //using lock
                } //try
                catch (System.Exception ex)
                {
                    MpExWin.Show(ex);
                }
            }
            /// <summary>
            /// Вставка элемента спецификации в строку таблицы AutoCad
            /// </summary>
            /// <param name="pos">Поз.</param>
            /// <param name="designation">Обозначение</param>
            /// <param name="name">Наименование</param>
            /// <param name="massa">Масса</param>
            /// <param name="note">Примечание</param>
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
                                                //Cell cell = new Cell(table, info.Row, 0) {TextString = designation};
                                                // Обозначение
                                                table.Cells[info.Row, 1].SetValue(designation, ParseOption.SetDefaultFormat);
                                                //cell = new Cell(table, info.Row, 1) {TextString = designation};
                                                // Наименование
                                                table.Cells[info.Row, 2].SetValue(name, ParseOption.SetDefaultFormat);
                                                //cell = new Cell(table, info.Row, 2) {TextString = name};
                                                //Масса
                                                table.Cells[info.Row, table.Columns.Count - 2].SetValue(massa, ParseOption.SetDefaultFormat);
                                                //cell = new Cell(table, info.Row, table.Columns.Count - 2) {TextString = massa};
                                                // Примечание
                                                table.Cells[info.Row, table.Columns.Count - 1].SetValue(note, ParseOption.SetDefaultFormat);
                                                //cell = new Cell(table, info.Row, table.Columns.Count - 1) { TextString = note };
                                                flag = true;
                                            }
                                        }
                                        catch
                                        {
                                            MpMsgWin.Show("Не попали в ячейку!");
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
                                                //cell = new Cell(table, info.Row, 1);
                                                //cell.TextString = name;
                                                //table.SetTextString(info.Row, 1, name);
                                                //Масса
                                                table.Cells[info.Row, table.Columns.Count - 1].SetValue(massa, ParseOption.SetDefaultFormat);
                                                //cell = new Cell(table, info.Row, table.Columns.Count - 1);
                                                //cell.TextString = massa;
                                                //table.SetTextString(info.Row, table.NumColumns - 1, massa);
                                                flag = true;
                                            }
                                        }
                                        catch
                                        {
                                            MpMsgWin.Show("Не попали в ячейку!");
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
                                            MpMsgWin.Show("Не попали в ячейку!");
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
                                            MpMsgWin.Show("Не попали в ячейку!");
                                        }
                                    }
                                }
                                else
                                {
                                    MpMsgWin.Show("Неверное количество столбцов в таблице!");
                                }
                                tr.Commit();
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MpExWin.Show(ex);
                }
            }
            /// <summary>
            /// Вставка однострочного текста
            /// </summary>
            /// <param name="text">Содержимое однострочного текста</param>
            public static void InsertDbText(string text)
            {
                try
                {
                    var doc = AcApp.DocumentManager.MdiActiveDocument;
                    var db = doc.Database;
                    var ed = doc.Editor;
                    //DocumentLock dlock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument();
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
                catch (System.Exception ex)
                {
                    MpExWin.Show(ex);
                }
            }
            /// <summary>
            /// Вставка мультивыноски
            /// </summary>
            /// <param name="txt">Содержимое мультивыноски</param>
            public static void InsertMLeader(string txt)
            {
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                var db = HostApplicationServices.WorkingDatabase;
                using (doc.LockDocument())
                {
                    var ppo = new PromptPointOptions("\nУкажите точку: ") { AllowNone = true };
                    var ppr = ed.GetPoint(ppo);
                    if (ppr.Status != PromptStatus.OK) return;
                    // Создаем текст
                    var jig = new MLeaderJig
                    {
                        FirstPoint = UcsToWcs(ppr.Value),
                        MlText = txt
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
            class DTextJig : EntityJig
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
                    catch (System.Exception)
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
            class MLeaderJig : DrawJig
            {
                private MLeader _mleader;
                public Point3d FirstPoint;
                private Point3d _secondPoint;
                private Point3d _prevPoint;
                public string MlText;

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
                        const string arrowName = "_NONE";
                        ObjectId arrId = GetArrowObjectId(arrowName);

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
                    MpExWin.Show(ex);
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
                                MpExWin.Show(ex);
                                tr.Commit();
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MpExWin.Show(ex);
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
                                MpExWin.Show(ex);
                                tr.Commit();
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MpExWin.Show(ex);
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
                    MpExWin.Show(ex);
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
                    MpExWin.Show(ex);
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
                MpExWin.Show(ex);
            }
        }
        /// <summary>
        /// Получение "реального" имени блока
        /// </summary>
        /// <param name="tr"></param>
        /// <param name="bref"></param>
        /// <returns></returns>
        static public string EffectiveBlockName(Transaction tr, BlockReference bref)
        {
            BlockTableRecord btr;
            if ((bref.IsDynamicBlock) | (bref.Name.StartsWith("*U", StringComparison.InvariantCultureIgnoreCase)))
            {
                btr = tr.GetObject(bref.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            }
            else
            {
                btr = tr.GetObject(bref.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

            }
            return btr?.Name;
        }
        /// <summary>
        /// Класс, содержащий различные константы используемые в функциях ModPlus
        /// </summary>
        public static class MpConstants
        {
            /// <summary>
            /// Русский алфавит. Верхний регистр
            /// </summary>
            public static readonly List<string> RusAlphabetUpper = new List<string> {"А", "Б","В","Г","Д","Е",
            "Ё","Ж","З","И","Й", "К","Л","М","Н","О","П","Р","С","Т","У","Ф","Х","Ц",
            "Ч","Щ","Ъ","Ы","Ь","Э","Ю","Я"};
            /// <summary>
            /// Русский алфавит. Нижний регистр
            /// </summary>
            public static readonly List<string> RusAlphabetLower = new List<string> {"а","б","в","г","д","е","ё","ж",
            "з","и","й", "к", "л", "м", "н", "о", "п" , "р", "с", "т", "у","ф","х","ц",
            "ч","щ","ъ","ы", "ь","э","ю","я"};
            /// <summary>
            /// Английский алфавит. Верхний регистр
            /// </summary>
            public static readonly List<string> EngAlphabetUpper = new List<string> {"A","B","C","D","E","F","G",
            "H","I","J","K","L","M", "N","O","P","Q","R","S","T","U","V","W","X","Y","Z" };
            /// <summary>
            /// Английский алфавит. Нижний регистр
            /// </summary>
            public static readonly List<string> EngAlphabetLower = new List<string> {"a","b","c","d","e","f","g",
            "h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x",",y","z"};
        }
    }

    public static class MpVars
    {
        // Имя пользователя (Логин)
        public static string UserName;
        // Почта пользователя для работы с сервером
        public static string UserEmail;
        // Регистрационный ключ
        public static string RegistryKey;
        // Лента
        public static bool MpRibbon;
        // Палитра
        public static bool MpPalette;
        // Функции в палитре
        public static bool MpPaletteFunctions;
        // Чертежи в палитре
        public static bool MpPaletteDrawings;
        // Плавающее меню
        public static bool MpFloatMenu;
        // Разделитель целой и дробной части
        public static string MpSeparator;
        // Окно "Чертежи"
        public static bool MpChkDrwsOnMnu;
        // Тихая загрузка
        public static bool QuietLoading;
        // Окно Чертежи
        public static bool DrawingsAlone;
        // Сворачивать плавающее меню в
        public static int FloatMenuCollapseTo;
        // Сворачивать Чертежи в
        public static int DrawingsCollapseTo;
        /// <summary>
        /// Чтение значений из файла настроек в глобальные переменные
        /// </summary>
        public static void ReadVarsFromSettingsfile()
        {
            try
            {
                UserName = MpSettings.GetValue("User", "Login");
                // read from regestry
                var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
                using (key)
                    if (key != null)
                        UserEmail = key.GetValue("email", string.Empty).ToString();
                // Так как до обновления был только один вариант получения регистрационного
                // ключа, то это нужно учитывать (хотя если был запущен конфигуратор, то такого быть не должно, но идиотов много)
                // Так как регистрационный ключ нужно получить один раз, то все процедуры делаем тут

                // Регистрационный ключ в зависимости от варианта привязки
                var regVariant = MpSettings.GetValue("User", "RegestryVariant");
                RegistryKey = string.Empty;
                if (!string.IsNullOrEmpty(regVariant))
                {
                    if (regVariant.Equals("0")) // К жесткому диску
                    {
                        var hdSerialNo = MpSettings.GetValue("User", "HDserialNo");
                        if (!string.IsNullOrEmpty(hdSerialNo))
                            RegistryKey = MpReg.GetUserRegKey(UserName, hdSerialNo);
                    }
                    else if (regVariant.Equals("1")) // to google
                    {
                        var gSub = MpSettings.GetValue("User", "gSub");
                        if (!string.IsNullOrEmpty(gSub))
                            RegistryKey = MpReg.GetUserRegKey(UserName, gSub);
                    }
                }
                bool b;
                MpRibbon = !bool.TryParse(MpSettings.GetValue("Settings", "MainSet", "Ribbon"), out b) || b; //true
                MpFloatMenu = bool.TryParse(MpSettings.GetValue("Settings", "MainSet", "FloatMenu"), out b) && b; //false
                MpChkDrwsOnMnu = bool.TryParse(MpSettings.GetValue("Settings", "MainSet", "ChkDrwsOnMnu"), out b) && b; //false
                MpSeparator = MpSettings.GetValue("Settings", "MainSet", "Separator") == "0" ? "." : ",";
                QuietLoading = !bool.TryParse(MpSettings.GetValue("Settings", "MainSet", "ChkQuietLoading"), out b) || b; //true
                DrawingsAlone = bool.TryParse(MpSettings.GetValue("Settings", "MainSet", "DrawingsAlone"), out b) && b; //false
                MpPalette = !bool.TryParse(MpSettings.GetValue("Settings", "MainSet", "Palette"), out b) || b; //true
                MpPaletteDrawings = !bool.TryParse(MpSettings.GetValue("Settings", "MainSet", "PaletteDrawings"), out b) || b; //true
                MpPaletteFunctions = !bool.TryParse(MpSettings.GetValue("Settings", "MainSet", "PaletteFunctions"), out b) || b; //true
                int i;
                FloatMenuCollapseTo = int.TryParse(MpSettings.GetValue("Settings", "MainSet", "FloatMenuCollapseTo"), out i) ? i : 0;
                DrawingsCollapseTo = int.TryParse(MpSettings.GetValue("Settings", "MainSet", "DrawingsCollapseTo"), out i) ? i : 1;
            }
            catch (System.Exception exception)
            {
                MpExWin.Show(exception);
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
                    MpPaletteSet = new PaletteSet("ModPlus", "mpPalette", new Guid("A9C907EF-6281-4FA2-9B6C-E0401E41BB74"));
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
            catch (System.Exception exception) { MpExWin.Show(exception); }
        }

        private static void AddRemovePaletts()
        {
            // functions
            if (MpVars.MpPaletteFunctions)
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
            if (MpVars.MpPaletteDrawings)
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

