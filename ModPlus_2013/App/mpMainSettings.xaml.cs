namespace ModPlus.App
{
    using System;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Documents;
    using System.Windows.Media.Imaging;
    using MinFuncWins;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using Windows;
    using ModPlusAPI.LicenseServer;
    using ModPlusStyle.Controls.Dialogs;
    using Utils = Autodesk.AutoCAD.Internal.Utils;

    /// <summary>
    /// Окно настроек ModPlus
    /// </summary>
    partial class MpMainSettings
    {
        private string _curTheme = string.Empty;
        private bool _curFloatMenu;
        private bool _curPalette;
        private bool _curDrawingsOnMenu;
        private bool _curRibbon;
        private bool _curDrawingsAlone;
        private int _curFloatMenuCollapseTo;
        private Language.LangItem _curLangItem;
        private int _curDrawingsCollapseTo = 1;
        private const string LangItem = "AutocadDlls";

        internal MpMainSettings()
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem(LangItem, "h1");
            FillAndSetLanguages();
            SetLanguageValues();
            FillThemesAndColors();
            LoadSettingsFromConfigFileAndRegistry();
            GetDataByVars();
            Closed += MpMainSettings_OnClosed;

            // license server
            if (ClientStarter.IsClientWorking())
            {
                BtStopConnectionToLicenseServer.IsEnabled = true;
                BtRestoreConnectionToLicenseServer.IsEnabled = false;
                TbLocalLicenseServerIpAddress.IsEnabled = false;
                TbLocalLicenseServerPort.IsEnabled = false;
            }
            else
            {
                BtStopConnectionToLicenseServer.IsEnabled = false;
                BtRestoreConnectionToLicenseServer.IsEnabled = true;
                TbLocalLicenseServerIpAddress.IsEnabled = true;
                TbLocalLicenseServerPort.IsEnabled = true;
            }
        }

        private void FillAndSetLanguages()
        {
            var languagesByFiles = ModPlusAPI.Language.GetLanguagesByFiles();
            CbLanguages.ItemsSource = languagesByFiles;
            CbLanguages.SelectedItem = languagesByFiles.FirstOrDefault(li => li.Name == ModPlusAPI.Language.CurrentLanguageName);
            _curLangItem = (Language.LangItem)CbLanguages.SelectedItem;
        }

        private void SetLanguageValues()
        {
            // Так как элементы окна по Серверу лицензий ссылаются на узел ModPlusAPI
            // присваиваю им значения в коде, после установки языка
            var li = "ModPlusAPI";
            GroupBoxLicenseServer.Header = ModPlusAPI.Language.GetItem(li, "h16");
            TbLocalLicenseServerIpAddressHeader.Text = ModPlusAPI.Language.GetItem(li, "h17");
            TbLocalLicenseServerPortHeader.Text = ModPlusAPI.Language.GetItem(li, "h18");
            BtCheckLocalLicenseServerConnection.Content = ModPlusAPI.Language.GetItem(li, "h19");
            BtStopConnectionToLicenseServer.Content = ModPlusAPI.Language.GetItem(li, "h23");
            BtRestoreConnectionToLicenseServer.Content = ModPlusAPI.Language.GetItem(li, "h24");
            ChkDisableConnectionWithLicenseServer.Content = ModPlusAPI.Language.GetItem(li, "h25");
        }

        private void FillThemesAndColors()
        {
            MiTheme.ItemsSource = ModPlusStyle.ThemeManager.Themes;
            var pluginStyle = ModPlusStyle.ThemeManager.Themes.First();
            var savedPluginStyleName = Regestry.GetValue("PluginStyle");
            if (!string.IsNullOrEmpty(savedPluginStyleName))
            {
                var theme = ModPlusStyle.ThemeManager.Themes.Single(t => t.Name == savedPluginStyleName);
                if (theme != null)
                    pluginStyle = theme;
            }

            _curTheme = pluginStyle.Name;
            MiTheme.SelectedItem = pluginStyle;
        }

        /// <summary>Загрузка данных из файла конфигурации которые требуется отобразить в окне</summary>
        private void LoadSettingsFromConfigFileAndRegistry()
        {
            // Separator
            var separator = Regestry.GetValue("Separator");
            CbSeparatorSettings.SelectedIndex = string.IsNullOrEmpty(separator) ? 0 : int.Parse(separator);
            // mini functions
            ChkEntByBlock.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "EntByBlockOCM"), out var b) || b; //true
            ChkNestedEntLayer.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "NestedEntLayerOCM"), out b) || b; //true
            ChkFastBlocks.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "FastBlocksCM"), out b) || b; //true
            ChkVPtoMS.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "VPtoMS"), out b) || b; //true
            ChkWipeoutEditOCM.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "WipeoutEditOCM"), out b) || b; //true
            ChkDisableConnectionWithLicenseServer.IsChecked =
                bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "DisableConnectionWithLicenseServerInAutoCAD"), out b) && b; // false
            TbLocalLicenseServerIpAddress.Text = Regestry.GetValue("LocalLicenseServerIpAddress");
            TbLocalLicenseServerPort.Value = int.TryParse(Regestry.GetValue("LocalLicenseServerPort"), out var i) ? i : 0;
        }

        /// <summary>Получение значений из глобальных переменных плагина</summary>
        private void GetDataByVars()
        {
            try
            {
                // Адаптация
                ChkMpFloatMenu.IsChecked = _curFloatMenu = Variables.FloatMenu;
                ChkMpPalette.IsChecked = _curPalette = Variables.Palette;
                // palette by visibility
                if (Variables.Palette && !MpPalette.MpPaletteSet.Visible)
                {
                    ChkMpPalette.IsChecked = _curPalette = false;
                    Variables.Palette = false; //
                }
                ChkMpPaletteFunctions.IsChecked = Variables.FunctionsInPalette;
                ChkMpPaletteDrawings.IsChecked = Variables.DrawingsInPalette;
                ChkMpRibbon.IsChecked = _curRibbon = Variables.Ribbon;
                ChkMpChkDrwsOnMnu.IsChecked = _curDrawingsOnMenu = Variables.DrawingsInFloatMenu;
                ChkMpDrawingsAlone.IsChecked = _curDrawingsAlone = Variables.DrawingsFloatMenu;
                // Выбор в выпадающих списках (сворачивать в)
                CbFloatMenuCollapseTo.SelectedIndex = _curFloatMenuCollapseTo = Variables.FloatMenuCollapseTo;
                CbDrawingsCollapseTo.SelectedIndex = _curDrawingsCollapseTo = Variables.DrawingsFloatMenuCollapseTo;
                // Видимость в зависимости от галочек
                ChkMpChkDrwsOnMnu.Visibility = CbFloatMenuCollapseTo.Visibility =
                        TbFloatMenuCollapseTo.Visibility = _curFloatMenu ? Visibility.Visible : Visibility.Collapsed;
                CbDrawingsCollapseTo.Visibility = TbDrawingsCollapseTo.Visibility = _curDrawingsAlone ? Visibility.Visible : Visibility.Collapsed;
                ChkMpPaletteDrawings.Visibility = ChkMpPaletteFunctions.Visibility = _curPalette ? Visibility.Visible : Visibility.Collapsed;
                // Тихая загрузка
                ChkQuietLoading.IsChecked = Variables.QuietLoading;
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        // Выбор разделителя целой и дробной части для чисел
        private void CbSeparatorSettings_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Regestry.SetValue("Separator", ((ComboBox)sender).SelectedIndex.ToString(CultureInfo.InvariantCulture));
        }

        // Выбор темы
        private void MiTheme_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var theme = (ModPlusStyle.Theme)e.AddedItems[0];
            Regestry.SetValue("PluginStyle", theme.Name);
            ModPlusStyle.ThemeManager.ChangeTheme(this, theme);
        }

        [SuppressMessage("ReSharper", "PossibleInvalidOperationException")]
        private void MpMainSettings_OnClosed(object sender, EventArgs e)
        {
            try
            {
                var isDifferentLanguage = IsDifferentLanguage();

                // Если отключили плавающее меню
                if (!ChkMpFloatMenu.IsChecked.Value)
                {
                    // Закрываем плавающее меню
                    if (MpMenuFunction.MpMainMenuWin != null)
                        MpMenuFunction.MpMainMenuWin.Close();
                }
                else // Если включили плавающее меню
                {
                    // Если плавающее меню было включено
                    if (MpMenuFunction.MpMainMenuWin != null)
                    {
                        // Перегружаем плавающее меню, если изменилась тема, вкл/выкл открытые чертежи, сворачивать в
                        if (!Regestry.GetValue("PluginStyle").Equals(_curTheme) ||
                            !Regestry.GetValue("FloatMenuCollapseTo").Equals(_curFloatMenuCollapseTo.ToString()) ||
                            !ChkMpChkDrwsOnMnu.IsChecked.Value.Equals(_curDrawingsOnMenu) ||
                            isDifferentLanguage)
                        {
                            MpMenuFunction.MpMainMenuWin.Close();
                            MpMenuFunction.LoadMainMenu();
                        }
                    }
                    else MpMenuFunction.LoadMainMenu();
                }

                // если отключили палитру
                if (!ChkMpPalette.IsChecked.Value)
                {
                    if (MpPalette.MpPaletteSet != null)
                        MpPalette.MpPaletteSet.Visible = false;
                }
                else // если включили палитру
                {
                    MpPalette.CreatePalette();
                }

                // Если отключили плавающее меню Чертежи
                if (!ChkMpDrawingsAlone.IsChecked.Value)
                {
                    if (MpDrawingsFunction.MpDrawingsWin != null)
                        MpDrawingsFunction.MpDrawingsWin.Close();
                }
                else
                {
                    if (MpDrawingsFunction.MpDrawingsWin != null)
                    {
                        // Перегружаем плавающее меню, если изменилась тема, вкл/выкл открытые чертежи, границы, сворачивать в
                        if (!Regestry.GetValue("PluginStyle").Equals(_curTheme) ||
                            !Regestry.GetValue("DrawingsCollapseTo").Equals(_curDrawingsCollapseTo.ToString()) ||
                            !ChkMpDrawingsAlone.IsChecked.Value.Equals(_curDrawingsAlone) ||
                            isDifferentLanguage)
                        {
                            MpDrawingsFunction.MpDrawingsWin.Close();
                            MpDrawingsFunction.LoadMainMenu();
                        }
                    }
                    else MpDrawingsFunction.LoadMainMenu();
                }

                // Ribbon
                // Если включили и была выключена
                if (ChkMpRibbon.IsChecked.Value && !_curRibbon)
                    RibbonBuilder.BuildRibbon();

                // Если включили и была включена, но сменился язык
                if (ChkMpRibbon.IsChecked.Value && _curRibbon && isDifferentLanguage)
                {
                    RibbonBuilder.RemoveRibbon();
                    RibbonBuilder.BuildRibbon(true);
                }

                // Если выключили и была включена
                if (!ChkMpRibbon.IsChecked.Value && _curRibbon)
                    RibbonBuilder.RemoveRibbon();

                // context menu
                // если сменился язык, то все выгружаю
                if (isDifferentLanguage)
                    MiniFunctions.UnloadAll();

                MiniFunctions.LoadUnloadContextMenu();

                // License server
                UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "DisableConnectionWithLicenseServerInAutoCAD",
                    ChkDisableConnectionWithLicenseServer.IsChecked.Value.ToString(), true);
                Regestry.SetValue("LocalLicenseServerIpAddress", TbLocalLicenseServerIpAddress.Text);
                Regestry.SetValue("LocalLicenseServerPort", TbLocalLicenseServerPort.Value.ToString());

                if (_restartClientOnClose)
                {
                    // reload server
                    ClientStarter.StopConnection();
                    ClientStarter.StartConnection(ProductLicenseType.AutoCAD);
                }

                // перевод фокуса на автокад
                Utils.SetFocusToDwgView();
            }
            catch (Exception ex)
            {
                ExceptionBox.Show(ex);
            }

        }

        private bool IsDifferentLanguage()
        {
            if (((Language.LangItem)CbLanguages.SelectedItem).Name != _curLangItem.Name)
                return true;

            return false;
        }

        /// <summary> Сохранение в файл конфигурации значений вкл/выкл для меню
        ///  Имена должны начинаться с ChkMp!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!</summary>
        private void Menues_OnChecked_Unchecked(object sender, RoutedEventArgs e)
        {
            if (!(sender is CheckBox chkBox)) return;
            var name = chkBox.Name;
            Regestry.SetValue(name.Substring(5), chkBox.IsChecked?.ToString());
            if (name.Equals("ChkMpFloatMenu"))
            {
                ChkMpChkDrwsOnMnu.Visibility = TbFloatMenuCollapseTo.Visibility = CbFloatMenuCollapseTo.Visibility =
                    chkBox.IsChecked != null && chkBox.IsChecked.Value ? Visibility.Visible : Visibility.Collapsed;
            }
            if (name.Equals("ChkMpDrawingsAlone"))
            {
                TbDrawingsCollapseTo.Visibility = CbDrawingsCollapseTo.Visibility =
                    chkBox.IsChecked != null && chkBox.IsChecked.Value ? Visibility.Visible : Visibility.Collapsed;
            }
            if (name.Equals("ChkMpPalette"))
            {
                ChkMpPaletteDrawings.Visibility = ChkMpPaletteFunctions.Visibility =
                    chkBox.IsChecked != null && chkBox.IsChecked.Value ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        // Тихая загрузка
        private void ChkQuietLoading_OnChecked_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Variables.QuietLoading = ChkQuietLoading.IsChecked != null && ChkQuietLoading.IsChecked.Value;
        }

        // Сворачивать в - для плавающего меню
        private void CbFloatMenuCollapseTo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb)
            {
                Variables.FloatMenuCollapseTo = cb.SelectedIndex;
            }
        }

        private void CbDrawingsCollapseTo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb)
            {
                Variables.DrawingsFloatMenuCollapseTo = cb.SelectedIndex;
            }
        }

        #region Контекстные меню

        // Задать вхождения ПоБлоку
        private void ChkEntByBlock_OnChecked(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox chk)
            {
                UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "EntByBlockOCM", (chk.IsChecked != null && chk.IsChecked.Value).ToString(), true);
                if (chk.IsChecked != null && chk.IsChecked.Value)
                    MiniFunctions.MiniFunctionsContextMenuExtensions.EntByBlockObjectContextMenu.Attach();
                else MiniFunctions.MiniFunctionsContextMenuExtensions.EntByBlockObjectContextMenu.Detach();
            }
        }

        private void ChkNestedEntLayer_OnChecked(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox chk)
            {
                UserConfigFile.SetValue("NestedEntLayerOCM", (chk.IsChecked != null && chk.IsChecked.Value).ToString(), true);
                if (chk.IsChecked != null && chk.IsChecked.Value)
                    MiniFunctions.MiniFunctionsContextMenuExtensions.NestedEntLayerObjectContextMenu.Attach();
                else MiniFunctions.MiniFunctionsContextMenuExtensions.NestedEntLayerObjectContextMenu.Detach();
            }
        }

        // Частоиспользуемые блоки
        private void ChkFastBlocks_OnChecked(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox chk)
            {
                UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "FastBlocksCM", (chk.IsChecked != null && chk.IsChecked.Value).ToString(), true);
                if (chk.IsChecked != null && chk.IsChecked.Value)
                    MiniFunctions.MiniFunctionsContextMenuExtensions.FastBlockContextMenu.Attach();
                else MiniFunctions.MiniFunctionsContextMenuExtensions.FastBlockContextMenu.Detach();
            }
        }
        private void BtFastBlocksSettings_OnClick(object sender, RoutedEventArgs e)
        {
            var win = new FastBlocksSettings();
            win.ShowDialog();
        }
        // Границы ВЭ в модель
        private void ChkVPtoMS_OnChecked(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox chk)
            {
                UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "VPtoMS", (chk.IsChecked != null && chk.IsChecked.Value).ToString(), true);
                if (chk.IsChecked != null && chk.IsChecked.Value)
                    MiniFunctions.MiniFunctionsContextMenuExtensions.VPtoMSObjectContextMenu.Attach();
                else MiniFunctions.MiniFunctionsContextMenuExtensions.VPtoMSObjectContextMenu.Detach();
            }
        }
        // wipeout edit
        private void ChkWipeoutEditOCM_OnChecked(object sender, RoutedEventArgs e)
        {
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "WipeoutEditOCM", true.ToString(), true);
            MiniFunctions.MiniFunctionsContextMenuExtensions.WipeoutEditObjectContextMenu.Attach();
        }
        private void ChkWipeoutEditOCM_OnUnchecked(object sender, RoutedEventArgs e)
        {
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "WipeoutEditOCM", false.ToString(), true);
            MiniFunctions.MiniFunctionsContextMenuExtensions.WipeoutEditObjectContextMenu.Detach();
        }
        #endregion

        private async void BtCheckLocalLicenseServerConnection_OnClick(object sender, RoutedEventArgs e)
        {
            // ReSharper disable once AsyncConverter.AsyncAwaitMayBeElidedHighlighting
            await this.ShowMessageAsync(
                ClientStarter.IsLicenseServerAvailable()
                    ? ModPlusAPI.Language.GetItem("ModPlusAPI", "h21")
                    : ModPlusAPI.Language.GetItem("ModPlusAPI", "h20"),
                ModPlusAPI.Language.GetItem("ModPlusAPI", "h22") + " " +
                TbLocalLicenseServerIpAddress.Text + ":" + TbLocalLicenseServerPort.Value).ConfigureAwait(true);
        }

        private bool _restartClientOnClose = true;

        private void BtStopConnectionToLicenseServer_OnClick(object sender, RoutedEventArgs e)
        {
            ClientStarter.StopConnection();
            BtRestoreConnectionToLicenseServer.IsEnabled = true;
            BtStopConnectionToLicenseServer.IsEnabled = false;
            TbLocalLicenseServerIpAddress.IsEnabled = true;
            TbLocalLicenseServerPort.IsEnabled = true;
            _restartClientOnClose = false;
        }

        private void BtRestoreConnectionToLicenseServer_OnClick(object sender, RoutedEventArgs e)
        {
            ClientStarter.StartConnection(ProductLicenseType.AutoCAD);
            BtRestoreConnectionToLicenseServer.IsEnabled = false;
            BtStopConnectionToLicenseServer.IsEnabled = true;
            TbLocalLicenseServerIpAddress.IsEnabled = false;
            TbLocalLicenseServerPort.IsEnabled = false;
            _restartClientOnClose = true;
        }

        private void MiniFunctionsHelpHyperlink_OnClick(object sender, RoutedEventArgs e)
        {
            if (sender is Hyperlink hyperlink)
            {
                var s = ModPlusAPI.Language.RusWebLanguages.Contains(ModPlusAPI.Language.CurrentLanguageName) ? "ru" : "en";
                Process.Start($"https://modplus.org/{s}/help/mini-plugins/{hyperlink.NavigateUri}");
            }
        }

        private void CbLanguages_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // fill image
            if (e.AddedItems[0] is Language.LangItem li)
            {
                ModPlusAPI.Language.SetCurrentLanguage(li.Name);
                this.SetLanguageProviderForModPlusWindow();
                SetLanguageValues();
                if (TbMessageAboutLanguage != null && _curLangItem != null)
                {
                    TbMessageAboutLanguage.Visibility = li.Name == _curLangItem.Name
                          ? Visibility.Collapsed
                          : Visibility.Visible;
                }
                try
                {
                    BitmapImage bi = new BitmapImage();
                    bi.BeginInit();
                    bi.UriSource = new Uri($"pack://application:,,,/ModPlus_{MpVersionData.CurCadVers};component/Resources/Flags/{li.Name}.png");
                    bi.EndInit();
                    LanguageImage.Source = bi;
                }
                catch
                {
                    LanguageImage.Source = null;
                }
            }
        }
    }
}
