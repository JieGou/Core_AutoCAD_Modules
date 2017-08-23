using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Internal;
using ModPlus.App;
using ModPlus.Helpers;
using ModPlusAPI;
using ModPlusAPI.Windows;
// AutoCad
#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif

namespace ModPlus
{
    partial class MpFloatMenu
    {
        // Переменные
        DocumentCollection Docs = AcApp.DocumentManager;
        string GlobalFileName = string.Empty;

        public MpFloatMenu()
        {
            try
            {
                Top = double.Parse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainMenuCoordinates", "top"));
                Left = double.Parse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainMenuCoordinates", "left"));
            }
            catch (Exception)
            {
                Top = 180;
                Left = 60;
            }
            InitializeComponent();
            ModPlusAPI.Windows.Helpers.WindowHelpers.ChangeThemeForResurceDictionary(this.Resources, true);

            MouseEnter += Window_MouseEnter;
            MouseLeave += Window_MouseLeave;
            MouseLeftButtonDown += Window_MouseLeftButtonDown;
            FillFieldsFunction();
            // Заполняем функции
            FillFunctions();

            ////////////////////////////////////////////////////////
            if (ModPlusAPI.Variables.DrawingsInFloatMenu)
            {
                // Подключение обработчиков событий для создания и закрытия чертежей
                AcApp.DocumentManager.DocumentCreated +=
                    DocumentManager_DocumentCreated;
                AcApp.DocumentManager.DocumentDestroyed +=
                    DocumentManager_DocumentDestroyed;
                //////////////////////////////
                try
                {
                    Drawings.Items.Clear();
                    foreach (Document doc in Docs)
                    {
                        var lbi = new ListBoxItem();
                        var filename = Path.GetFileName(doc.Name);
                        lbi.Content = filename;
                        lbi.ToolTip = doc.Name;
                        Drawings.Items.Add(lbi);
                        Drawings.SelectedItem = lbi;
                    }
                }
                catch
                {
                    // ignored
                }
            }
            else
            {
                // Подключение и отключение (чтобы не было ошибки) обработчиков событий для создания и закрытия чертежей
                AcApp.DocumentManager.DocumentCreated +=
                    DocumentManager_DocumentCreated;
                AcApp.DocumentManager.DocumentDestroyed +=
                    DocumentManager_DocumentDestroyed;
                ///////////////////////////
                AcApp.DocumentManager.DocumentCreated -=
                    DocumentManager_DocumentCreated;
                AcApp.DocumentManager.DocumentDestroyed -=
                    DocumentManager_DocumentDestroyed;
            }
            // Обрабатываем событие покидания мышкой окна
            OnMouseLeaving();
        }
        // Заполнение списка функций
        private void FillFunctions()
        {
            try
            {
                // Расположение файла конфигурации
                var confF = UserConfigFile.FullFileName;
                // Грузим
                var configFile = XElement.Load(confF);
                // Проверяем есть ли группа Config
                if (configFile.Element("Config") == null)
                {
                    ModPlusAPI.Windows.MessageBox.Show("Файл конфигурации поврежден! Невозможно заполнить плавающее меню", MessageBoxIcon.Alert);
                    return;
                }
                var element = configFile.Element("Config");
                // Проверяем есть ли подгруппа Cui
                if (element?.Element("CUI") == null)
                {
                    ModPlusAPI.Windows.MessageBox.Show("Файл конфигурации поврежден! Невозможно заполнить плавающее меню", MessageBoxIcon.Alert);
                    return;
                }
                var confCuiXel = element.Element("CUI");
                // Проходим по группам
                if (confCuiXel == null) return;
                foreach (var group in confCuiXel.Elements("Group"))
                {
                    var exp = new Expander
                    {
                        Header = group.Attribute("GroupName")?.Value,
                        IsExpanded = false,
                        Margin = new Thickness(1)
                    };
                    var expStck = new StackPanel { Orientation = Orientation.Vertical };

                    // Проходим по функциям группы
                    foreach (var func in group.Elements("Function"))
                    {
                        var funcNameAttr = func.Attribute("Name")?.Value;
                        if (string.IsNullOrEmpty(funcNameAttr)) continue;

                        var loadedFunction = LoadFunctionsHelper.LoadedFunctions.FirstOrDefault(x => x.Name.Equals(funcNameAttr));
                        if (loadedFunction == null) continue;

                        expStck.Children.Add(
                            WPFMenuesHelper.AddButton(this, loadedFunction.Name, loadedFunction.LName,
                                loadedFunction.BigIconUrl, loadedFunction.Description,
                                loadedFunction.FullDescription, loadedFunction.ToolTipHelpImage,true)
                        );
                        if (loadedFunction.SubFunctionsNames.Any())
                        {
                            for (int i = 0; i < loadedFunction.SubFunctionsNames.Count; i++)
                            {
                                expStck.Children.Add(WPFMenuesHelper.AddButton(this,
                                    loadedFunction.SubFunctionsNames[i],
                                    loadedFunction.SubFunctionsLNames[i], loadedFunction.SubBigIconsUrl[i],
                                    loadedFunction.SubDescriptions[i], loadedFunction.SubFullDescriptions[i],
                                    loadedFunction.SubHelpImages[i], true));
                            }
                        }

                        foreach (var subFunc in func.Elements("SubFunction"))
                        {
                            var subFuncNameAttr = subFunc.Attribute("Name")?.Value;
                            if (string.IsNullOrEmpty(subFuncNameAttr)) continue;
                            var loadedSubFunction = LoadFunctionsHelper.LoadedFunctions.FirstOrDefault(x => x.Name.Equals(subFuncNameAttr));
                            if (loadedSubFunction == null) continue;
                            expStck.Children.Add(
                                WPFMenuesHelper.AddButton(this, loadedSubFunction.Name, loadedSubFunction.LName,
                                loadedSubFunction.BigIconUrl, loadedSubFunction.Description,
                                loadedSubFunction.FullDescription, loadedSubFunction.ToolTipHelpImage, true)
                                );
                        }
                    }
                    exp.Content = expStck;
                    // Добавляем группу, если заполнились функции!
                    if (expStck.Children.Count > 0)
                        FunctionsPanel.Children.Add(exp);
                }
            }
            catch (Exception exception) { ExceptionBox.ShowForConfigurator(exception); }
        }
        
        // Чертеж закрыт
        void DocumentManager_DocumentDestroyed(object sender, DocumentDestroyedEventArgs e)
        {
            try
            {
                foreach (var lbi in Drawings.Items.Cast<ListBoxItem>().Where(
                    lbi => lbi.ToolTip.ToString() == e.FileName))
                {
                    Drawings.Items.Remove(lbi);
                    break;
                }
            }
            catch
            {
                // ignored
            }
        }
        // Документ создан/открыт
        void DocumentManager_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            try
            {
                var lbi = new ListBoxItem();
                var filename = Path.GetFileName(e.Document.Name);
                lbi.Content = filename;
                lbi.ToolTip = e.Document.Name;
                Drawings.Items.Add(lbi);
                Drawings.SelectedItem = lbi;
            }
            catch
            {
                // ignored
            }
        }
        // Наведение мышки на окно
        private void Window_MouseEnter(object sender, MouseEventArgs e)
        {
            if (Docs.Count > 0)
            {
                ExpMpFunctions.Visibility = Visibility.Visible;
                ImgIcon.Visibility = Visibility.Collapsed;
                TbHeader.Visibility = Visibility.Visible;
                BtMpSettings.Visibility = Visibility.Visible;
                if (ModPlusAPI.Variables.DrawingsInFloatMenu)
                {
                    ExpOpenDrawings.Visibility = Visibility.Visible;
                    //////////////////////////////////
                    if (Docs.Count != Drawings.Items.Count)
                    {
                        var names = new string[Docs.Count];
                        var docnames = new string[Docs.Count];
                        var i = 0;
                        foreach (Document doc in Docs)
                        {
                            var filename = Path.GetFileName(doc.Name);
                            names.SetValue(filename, i);
                            docnames.SetValue(doc.Name, i);
                            i++;
                        }
                        foreach (var lbi in Drawings.Items.Cast<ListBoxItem>().Where(
                            lbi => !docnames.Contains(lbi.ToolTip)))
                        {
                            Drawings.Items.Remove(lbi);
                            break;
                        }
                    }
                    try
                    {
                        Drawings.Items.Clear();
                        foreach (Document doc in Docs)
                        {
                            var lbi = new ListBoxItem();
                            var filename = Path.GetFileName(doc.Name);
                            lbi.Content = filename;
                            lbi.ToolTip = doc.Name;
                            Drawings.Items.Add(lbi);
                        }
                    }
                    catch
                    {
                        // ignored
                    }
                    try
                    {
                        foreach (var lbi in Drawings.Items.Cast<ListBoxItem>().Where(
                            lbi => lbi.ToolTip.ToString() == Docs.MdiActiveDocument.Name))
                        {
                            Drawings.SelectedItem = lbi;
                            break;
                        }
                    }
                    catch
                    {
                        // ignored
                    }
                }
                //////////////////////////////////
                Focus();

            }
        }
        // Убирание мышки с окна
        private void Window_MouseLeave(object sender, MouseEventArgs e)
        {
            if (Docs.Count > 0)
                OnMouseLeaving();
        }
        private void OnMouseLeaving()
        {
            if (ModPlusAPI.Variables.FloatMenuCollapseTo.Equals(0)) //icon
            {
                ImgIcon.Visibility = Visibility.Visible;
                TbHeader.Visibility = Visibility.Collapsed;
            }
            else // header
            {
                ImgIcon.Visibility = Visibility.Collapsed;
                TbHeader.Visibility = Visibility.Visible;
                BtMpSettings.Visibility = Visibility.Collapsed;
            }
            ExpMpFunctions.Visibility = Visibility.Collapsed;
            ExpOpenDrawings.Visibility = Visibility.Collapsed;
            Utils.SetFocusToDwgView();
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
        // Выбор чертежа в списке открытых
        private void Drawings_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var lbi = (ListBoxItem)Drawings.SelectedItem;
                foreach (
                    var doc in
                    from Document doc
                        in Docs
                    let filename = Path.GetFileName(doc.Name)
                    where doc.Name == lbi.ToolTip.ToString() & filename == lbi.Content.ToString()
                    select doc)
                {
                    if (Docs.MdiActiveDocument != null && Docs.MdiActiveDocument != doc)
                    {
                        Docs.MdiActiveDocument = doc;
                    }
                    break;
                }
            }
            catch
            {
                // ignored
            }
        }
        // Нажатие кнопки закрытия чертежа
        private void BtCloseDwg_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Drawings.SelectedIndex != -1)
                {
                    var lbi = (ListBoxItem)Drawings.SelectedItem;
                    foreach (var doc in Docs.Cast<Document>().Where(doc => doc.Name == lbi.ToolTip.ToString()))
                    {
                        if (Docs.MdiActiveDocument == doc)
                        {
                            AcApp.DocumentManager.
                                MdiActiveDocument.SendStringToExecute("_CLOSE ", true, false, false);
                            if (Drawings.Items.Count == 1)
                                OnMouseLeaving();
                        }
                        break;
                    }
                }
            }
            catch
            {
                // ignored
            }
        }

        // Вызов окна настроек ModPlus
        private void BtMpSettings_OnClick(object sender, RoutedEventArgs e)
        {
            var win = new MpMainSettings();
            win.ShowDialog();
        }
        // start fields fucntion
        private void BtFields_OnClick(object sender, RoutedEventArgs e)
        {
            if (AcApp.DocumentManager.Count > 0)
                AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_MPSTAMPFIELDS ", false, false, false);
        }
        private void FillFieldsFunction()
        {
            BtFields.Visibility = LoadFunctionsHelper.HasmpStampsFunction() ? Visibility.Visible : Visibility.Collapsed;
        }
    }
    public static class MpMenuFunction
    {
        public static MpFloatMenu MpMainMenuWin;
        /// <summary>
        /// Загрузка основного меню в зависимости от настроек
        /// </summary>
        public static void LoadMainMenu()
        {
            if (ModPlusAPI.Variables.FloatMenu)
            {
                if (MpMainMenuWin == null)
                {
                    MpMainMenuWin = new MpFloatMenu();
                    MpMainMenuWin.Closed += mpMainMenuWin_Closed;
                }
                if (MpMainMenuWin.IsLoaded)
                    return;
                AcApp.ShowModelessWindow(
                    AcApp.MainWindow.Handle, MpMainMenuWin);
            }
            else
            {
                MpMainMenuWin?.Close();
            }
        }


        static void mpMainMenuWin_Closed(object sender, EventArgs e)
        {
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "MainMenuCoordinates", "top", MpMainMenuWin.Top.ToString(CultureInfo.InvariantCulture), true);
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "MainMenuCoordinates", "left", MpMainMenuWin.Left.ToString(CultureInfo.InvariantCulture), true);
            MpMainMenuWin = null;
        }
    }
}
