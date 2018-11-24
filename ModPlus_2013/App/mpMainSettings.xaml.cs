namespace ModPlus.App
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Media;
    using Autodesk.AutoCAD.Internal;
    using Autodesk.AutoCAD.Runtime;
    using MinFuncWins;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using Windows;

    partial class MpMainSettings
    {
        private string _curTheme = string.Empty;
        private bool _curFloatMenu;
        private bool _curPalette;
        private bool _curDrawingsOnMenu;
        private bool _curRibbon;
        private bool _curDrawingsAlone;
        private int _curFloatMenuCollapseTo = 0;
        private int _curDrawingsCollapseTo = 1;
        private const string LangItem = "AutocadDlls";

        internal MpMainSettings()
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem(LangItem, "h1");
            FillThemesAndColors();
            SetAppRegistryKeyForCurrentUser();
            GetDataFromConfigFile();
            GetDataByVars();
            Closing += MpMainSettings_Closing;
            Closed += MpMainSettings_OnClosed;
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
        // Заполнение поля Ключ продукта
        private void SetAppRegistryKeyForCurrentUser()
        {
            // Ключ берем из глобальных настроек
            var key = ModPlusAPI.Variables.RegistryKey;
            if (string.IsNullOrEmpty(key))
            {
                TbAboutRegKey.Visibility = Visibility.Collapsed;
                TbRegistryKey.Text = string.Empty;
            }
            else
            {
                TbRegistryKey.Text = key;
                var regVariant = Regestry.GetValue("RegestryVariant");
                if (!string.IsNullOrEmpty(regVariant))
                {
                    TbAboutRegKey.Visibility = Visibility.Visible;
                    if (regVariant.Equals("0"))
                        TbAboutRegKey.Text = ModPlusAPI.Language.GetItem(LangItem, "h10") + " " +
                                             Regestry.GetValue("HDmodel");
                    else if (regVariant.Equals("1"))
                        TbAboutRegKey.Text = ModPlusAPI.Language.GetItem(LangItem, "h11") + " " +
                                             Regestry.GetValue("gName");
                }
            }
        }

        /// <summary>Загрузка данных из файла конфигурации которые требуется отобразить в окне</summary>
        private void GetDataFromConfigFile()
        {
            // Separator
            var separator = Regestry.GetValue("Separator");
            CbSeparatorSettings.SelectedIndex = string.IsNullOrEmpty(separator) ? 0 : int.Parse(separator);
            // mini functions
            ChkEntByBlock.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "EntByBlockOCM"), out var b) || b; //true
            ChkFastBlocks.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "FastBlocksCM"), out b) || b; //true
            ChkVPtoMS.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "VPtoMS"), out b) || b; //true
            ChkWipeoutEditOCM.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "WipeoutEditOCM"), out b) || b; //true
        }
        /// <summary>Получение значений из глобальных переменных плагина</summary>
        private void GetDataByVars()
        {
            try
            {
                // Адаптация
                ChkMpFloatMenu.IsChecked = _curFloatMenu = ModPlusAPI.Variables.FloatMenu;
                ChkMpPalette.IsChecked = _curPalette = ModPlusAPI.Variables.Palette;
                // palette by visibility
                if (ModPlusAPI.Variables.Palette && !MpPalette.MpPaletteSet.Visible)
                {
                    ChkMpPalette.IsChecked = _curPalette = false;
                    ModPlusAPI.Variables.Palette = false; //
                }
                ChkMpPaletteFunctions.IsChecked = ModPlusAPI.Variables.FunctionsInPalette;
                ChkMpPaletteDrawings.IsChecked = ModPlusAPI.Variables.DrawingsInPalette;
                ChkMpRibbon.IsChecked = _curRibbon = ModPlusAPI.Variables.Ribbon;
                ChkMpChkDrwsOnMnu.IsChecked = _curDrawingsOnMenu = ModPlusAPI.Variables.DrawingsInFloatMenu;
                ChkMpDrawingsAlone.IsChecked = _curDrawingsAlone = ModPlusAPI.Variables.DrawingsFloatMenu;
                // Выбор в выпадающих списках (сворачивать в)
                CbFloatMenuCollapseTo.SelectedIndex = _curFloatMenuCollapseTo = ModPlusAPI.Variables.FloatMenuCollapseTo;
                CbDrawingsCollapseTo.SelectedIndex = _curDrawingsCollapseTo = ModPlusAPI.Variables.DrawingsFloatMenuCollapseTo;
                // Видимость в зависимости от галочек
                ChkMpChkDrwsOnMnu.Visibility = CbFloatMenuCollapseTo.Visibility =
                        TbFloatMenuCollapseTo.Visibility = _curFloatMenu ? Visibility.Visible : Visibility.Collapsed;
                CbDrawingsCollapseTo.Visibility = TbDrawingsCollapseTo.Visibility = _curDrawingsAlone ? Visibility.Visible : Visibility.Collapsed;
                ChkMpPaletteDrawings.Visibility = ChkMpPaletteFunctions.Visibility = _curPalette ? Visibility.Visible : Visibility.Collapsed;
                // Тихая загрузка
                ChkQuietLoading.IsChecked = ModPlusAPI.Variables.QuietLoading;
                // email
                TbEmailAdress.Text = ModPlusAPI.Variables.UserEmail;
            }
            catch (System.Exception exception)
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

        private void MpMainSettings_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!string.IsNullOrEmpty(TbEmailAdress.Text))
            {
                if (IsValidEmail(TbEmailAdress.Text))
                    TbEmailAdress.BorderBrush = FindResource("BlackBrush") as Brush;
                else
                {
                    TbEmailAdress.BorderBrush = Brushes.Red;
                    ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "tt4"));
                    TbEmailAdress.Focus();
                    e.Cancel = true;
                }
            }
        }

        [SuppressMessage("ReSharper", "PossibleInvalidOperationException")]
        private void MpMainSettings_OnClosed(object sender, EventArgs e)
        {
            try
            {
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
                        if (!string.IsNullOrEmpty(_curTheme) &
                            !string.IsNullOrEmpty(_curFloatMenuCollapseTo.ToString()))
                        {
                            if (!Regestry.GetValue("PluginStyle").Equals(_curTheme) |
                                !Regestry.GetValue("FloatMenuCollapseTo").Equals(_curFloatMenuCollapseTo.ToString()) |
                                !ChkMpChkDrwsOnMnu.IsChecked.Value.Equals(_curDrawingsOnMenu))
                            {
                                MpMenuFunction.MpMainMenuWin.Close();
                                MpMenuFunction.LoadMainMenu();
                            }
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
                        if (!string.IsNullOrEmpty(_curTheme) &
                            !string.IsNullOrEmpty(_curDrawingsCollapseTo.ToString()))
                        {
                            if (!Regestry.GetValue("PluginStyle").Equals(_curTheme) |
                                !Regestry.GetValue("DrawingsCollapseTo").Equals(_curDrawingsCollapseTo.ToString()) |
                                !ChkMpDrawingsAlone.IsChecked.Value.Equals(_curDrawingsAlone))
                            {
                                MpDrawingsFunction.MpDrawingsWin.Close();
                                MpDrawingsFunction.LoadMainMenu();
                            }
                        }
                    }
                    else MpDrawingsFunction.LoadMainMenu();
                }

                // Ribbon
                if(ChkMpRibbon.IsChecked.Value && !_curRibbon)
                    RibbonBuilder.BuildRibbon();

                if(!ChkMpRibbon.IsChecked.Value && _curRibbon)
                    RibbonBuilder.RemoveRibbon();

                // context menues
                MiniFunctions.LoadUnloadContextMenu();

                // перевод фокуса на автокад
                Utils.SetFocusToDwgView();
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
            }

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
            ModPlusAPI.Variables.QuietLoading = ChkQuietLoading.IsChecked != null && ChkQuietLoading.IsChecked.Value;
        }

        // Сворачивать в - для плавающего меню
        private void CbFloatMenuCollapseTo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb)
            {
                ModPlusAPI.Variables.FloatMenuCollapseTo = cb.SelectedIndex;
            }
        }

        private void CbDrawingsCollapseTo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb)
            {
                ModPlusAPI.Variables.DrawingsFloatMenuCollapseTo = cb.SelectedIndex;
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

        private void TbEmailAdress_OnLostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb)
            {
                if (IsValidEmail(tb.Text))
                    tb.BorderBrush = FindResource("BlackBrush") as Brush;
                else tb.BorderBrush = Brushes.Red;
            }
        }

        private static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
    }

    public class MpMainSettingsFunction
    {
        [CommandMethod("ModPlus", "mpSettings", CommandFlags.Modal)]
        public void Main()
        {
            try
            {
                var win = new MpMainSettings();
                win.ShowDialog();
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }
    }
}
