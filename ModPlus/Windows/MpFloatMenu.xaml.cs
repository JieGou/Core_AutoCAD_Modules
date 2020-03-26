namespace ModPlus.Windows
{
    using System;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.Internal;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
    using App;
    using Helpers;

    partial class MpFloatMenu : Window
    {
        // Переменные
        readonly DocumentCollection _docs = Application.DocumentManager;

        // Переменная хранит значение о наличии функции "Поля"
        private bool _hasFieldsFunction;

        private const string LangItem = "AutocadDlls";

        public MpFloatMenu()
        {
            if (double.TryParse(Regestry.GetValue("FloatingMenuTop"), out var top))
                Top = top;
            else
                Top = 180;

            if (double.TryParse(Regestry.GetValue("FloatingMenuLeft"), out var left))
                Left = left;
            else
                Left = 60;

            InitializeComponent();

            Closed += MpFloatMenu_OnClosed;

            ModPlusAPI.Windows.Helpers.WindowHelpers.ChangeStyleForResourceDictionary(Resources);
            ModPlusAPI.Language.SetLanguageProviderForResourceDictionary(Resources);

            MouseEnter += Window_MouseEnter;
            MouseLeave += Window_MouseLeave;
            MouseLeftButtonDown += Window_MouseLeftButtonDown;
            FillFieldsFunction();

            // Заполняем функции
            FillFunctions();

            if (Variables.DrawingsInFloatMenu)
            {
                // Подключение обработчиков событий для создания и закрытия чертежей
                Application.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;
                Application.DocumentManager.DocumentDestroyed += DocumentManager_DocumentDestroyed;
                try
                {
                    Drawings.Items.Clear();
                    foreach (Document doc in _docs)
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
                Application.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;
                Application.DocumentManager.DocumentDestroyed += DocumentManager_DocumentDestroyed;
                ///////////////////////////
                Application.DocumentManager.DocumentCreated -= DocumentManager_DocumentCreated;
                Application.DocumentManager.DocumentDestroyed -= DocumentManager_DocumentDestroyed;
            }

            // Обрабатываем событие покидания мышкой окна
            OnMouseLeaving();
        }

        // Заполнение списка функций
        private void FillFunctions()
        {
            try
            {
                var confCuiXel = ModPlusAPI.RegistryData.Adaptation.GetCuiAsXElement("AutoCAD");

                // Проходим по группам
                if (confCuiXel == null)
                    return;
                foreach (var group in confCuiXel.Elements("Group"))
                {
                    var exp = new Expander
                    {
                        Header = ModPlusAPI.Language.TryGetCuiLocalGroupName(group.Attribute("GroupName")?.Value),
                        IsExpanded = false,
                        Margin = new Thickness(1)
                    };

                    var expStck = new StackPanel { Orientation = Orientation.Vertical };

                    // Проходим по функциям группы
                    foreach (var func in group.Elements("Function"))
                    {
                        var funcNameAttr = func.Attribute("Name")?.Value;
                        if (string.IsNullOrEmpty(funcNameAttr))
                            continue;

                        var loadedFunction = LoadFunctionsHelper.LoadedFunctions.FirstOrDefault(x => x.Name.Equals(funcNameAttr));
                        if (loadedFunction == null)
                            continue;

                        expStck.Children.Add(
                            WPFMenuesHelper.AddButton(
                                this,
                                loadedFunction.Name,
                                ModPlusAPI.Language.GetFunctionLocalName(loadedFunction.Name, loadedFunction.LName),
                                loadedFunction.BigIconUrl,
                                ModPlusAPI.Language.GetFunctionShortDescription(loadedFunction.Name, loadedFunction.Description),
                                ModPlusAPI.Language.GetFunctionFullDescription(loadedFunction.Name, loadedFunction.FullDescription),
                                loadedFunction.ToolTipHelpImage, true));
                        if (loadedFunction.SubFunctionsNames.Any())
                        {
                            for (int i = 0; i < loadedFunction.SubFunctionsNames.Count; i++)
                            {
                                expStck.Children.Add(WPFMenuesHelper.AddButton(
                                    this,
                                    loadedFunction.SubFunctionsNames[i],
                                    ModPlusAPI.Language.GetFunctionLocalName(loadedFunction.Name, loadedFunction.SubFunctionsLNames[i], i + 1),
                                    loadedFunction.SubBigIconsUrl[i],
                                    ModPlusAPI.Language.GetFunctionShortDescription(loadedFunction.Name, loadedFunction.SubDescriptions[i], i + 1),
                                    ModPlusAPI.Language.GetFunctionFullDescription(loadedFunction.Name, loadedFunction.SubFullDescriptions[i], i + 1),
                                    loadedFunction.SubHelpImages[i], true));
                            }
                        }

                        foreach (var subFunc in func.Elements("SubFunction"))
                        {
                            var subFuncNameAttr = subFunc.Attribute("Name")?.Value;
                            if (string.IsNullOrEmpty(subFuncNameAttr))
                                continue;
                            var loadedSubFunction = LoadFunctionsHelper.LoadedFunctions.FirstOrDefault(x => x.Name.Equals(subFuncNameAttr));
                            if (loadedSubFunction == null)
                                continue;
                            expStck.Children.Add(
                                WPFMenuesHelper.AddButton(
                                    this,
                                    loadedSubFunction.Name,
                                    ModPlusAPI.Language.GetFunctionLocalName(loadedSubFunction.Name, loadedSubFunction.LName),
                                    loadedSubFunction.BigIconUrl,
                                    ModPlusAPI.Language.GetFunctionShortDescription(loadedSubFunction.Name, loadedSubFunction.Description),
                                    ModPlusAPI.Language.GetFunctionFullDescription(loadedSubFunction.Name, loadedSubFunction.FullDescription),
                                    loadedSubFunction.ToolTipHelpImage, true));
                        }
                    }

                    exp.Content = expStck;

                    // Добавляем группу, если заполнились функции!
                    if (expStck.Children.Count > 0)
                        FunctionsPanel.Children.Add(exp);
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
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
            if (_docs.Count > 0)
            {
                ExpMpFunctions.Visibility = Visibility.Visible;
                ImgIcon.Visibility = Visibility.Collapsed;
                TbHeader.Visibility = Visibility.Visible;
                BtMpSettings.Visibility = Visibility.Visible;
                if (_hasFieldsFunction)
                    BtFields.Visibility = Visibility.Visible;
                if (Variables.DrawingsInFloatMenu)
                {
                    ExpOpenDrawings.Visibility = Visibility.Visible;

                    if (_docs.Count != Drawings.Items.Count)
                    {
                        var names = new string[_docs.Count];
                        var docnames = new string[_docs.Count];
                        var i = 0;
                        foreach (Document doc in _docs)
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
                        foreach (Document doc in _docs)
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
                            lbi => lbi.ToolTip.ToString() == _docs.MdiActiveDocument.Name))
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

                Focus();
            }
        }

        // Убирание мышки с окна
        private void Window_MouseLeave(object sender, MouseEventArgs e)
        {
            if (_docs.Count > 0)
                OnMouseLeaving();
        }

        private void OnMouseLeaving()
        {
            if (Variables.FloatMenuCollapseTo.Equals(0)) //// icon
            {
                ImgIcon.Visibility = Visibility.Visible;
                TbHeader.Visibility = Visibility.Collapsed;
            }
            else // header
            {
                ImgIcon.Visibility = Visibility.Collapsed;
                TbHeader.Visibility = Visibility.Visible;
                BtMpSettings.Visibility = Visibility.Collapsed;
                BtFields.Visibility = Visibility.Collapsed;
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
                        in _docs
                    let filename = Path.GetFileName(doc.Name)
                    where doc.Name == lbi.ToolTip.ToString() & filename == lbi.Content.ToString()
                    select doc)
                {
                    if (_docs.MdiActiveDocument != null && _docs.MdiActiveDocument != doc)
                    {
                        _docs.MdiActiveDocument = doc;
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
                    foreach (var doc in _docs.Cast<Document>().Where(doc => doc.Name == lbi.ToolTip.ToString()))
                    {
                        if (_docs.MdiActiveDocument == doc)
                        {
                            Application.DocumentManager.
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

        // start fields function
        private void BtFields_OnClick(object sender, RoutedEventArgs e)
        {
            if (Application.DocumentManager.Count > 0)
                Application.DocumentManager.MdiActiveDocument.SendStringToExecute("_MPSTAMPFIELDS ", false, false, false);
        }

        private void FillFieldsFunction()
        {
            _hasFieldsFunction = LoadFunctionsHelper.HasStampsPlugin();
            BtFields.Visibility = _hasFieldsFunction ? Visibility.Visible : Visibility.Collapsed;
        }

        private void MpFloatMenu_OnClosed(object sender, EventArgs e)
        {
            try
            {
                MouseEnter -= Window_MouseEnter;
                MouseLeave -= Window_MouseLeave;
                MouseLeftButtonDown -= Window_MouseLeftButtonDown;
                Application.DocumentManager.DocumentCreated -= DocumentManager_DocumentCreated;
                Application.DocumentManager.DocumentDestroyed -= DocumentManager_DocumentDestroyed;
            }
            catch
            {
                // ignore
            }
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
            try
            {
                if (Variables.FloatMenu)
                {
                    if (MpMainMenuWin == null)
                    {
                        MpMainMenuWin = new MpFloatMenu();
                        MpMainMenuWin.Closed += mpMainMenuWin_Closed;
                    }

                    if (MpMainMenuWin.IsLoaded)
                        return;

                    Application.ShowModelessWindow(Application.MainWindow.Handle, MpMainMenuWin);
                }
                else
                {
                    MpMainMenuWin?.Close();
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        static void mpMainMenuWin_Closed(object sender, EventArgs e)
        {
            Regestry.SetValue("FloatingMenuTop", MpMainMenuWin.Top.ToString(CultureInfo.InvariantCulture));
            Regestry.SetValue("FloatingMenuLeft", MpMainMenuWin.Left.ToString(CultureInfo.InvariantCulture));
            MpMainMenuWin = null;
        }
    }
}
