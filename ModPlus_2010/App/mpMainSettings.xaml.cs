using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Runtime;
using MahApps.Metro;
using ModPlus.MinFuncWins;
using ModPlusAPI;
using ModPlusAPI.Windows;
using ModPlusAPI.Windows.Helpers;

namespace ModPlus.App
{
    partial class MpMainSettings
    {
        private string _curTheme = string.Empty;
        private string _curColor = string.Empty;
        private bool _curFloatMenu;
        private bool _curPalette;
        private bool _curDrwsOnMnu;
        private bool _curRibbon;
        private bool _curDrawingsAlone;
        private int _curFloatMenuCollapseTo = 0;
        private int _curDrawingsCollapseTo = 1;
        private string _curBordersType = string.Empty;
        private readonly string _curLang;
        public List<AccentColorMenuData> AccentColors { get; set; }
        public List<AppThemeMenuData> AppThemes { get; set; }
        private const string LangItem = "AutocadDlls";

        internal MpMainSettings()
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem(LangItem, "h1");
            LoadIcon();
            FillThemesAndColors();
            SetAppRegistryKeyForCurrentUser();
            GetDataFromConfigFile();
            GetDataByVars();
            Closing += MpMainSettings_Closing;
            Closed += MpMainSettings_OnClosed;
            // fill languages
            CbLanguages.ItemsSource = ModPlusAPI.Language.GetLanguagesByFiles();
            CbLanguages.SelectedItem = ((List<Language.LangItem>)CbLanguages.ItemsSource)
                .FirstOrDefault(x => x.Name.Equals(ModPlusAPI.Language.CurrentLanguageName));
            _curLang = ((Language.LangItem)CbLanguages.SelectedItem)?.Name;
            CbLanguages.SelectionChanged += CbLanguages_SelectionChanged;
        }
        // Change language
        private void CbLanguages_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb && cb.SelectedItem is Language.LangItem langItem)
            {
                ModPlusAPI.Language.SetCurrentLanguage(langItem);
                ModPlusAPI.Language.SetLanguageProviderForWindow(this);
            }
        }

        private void FillThemesAndColors()
        {
            // create accent color menu items for the demo
            AccentColors = ThemeManager.Accents
                                            .Select(a => new AccentColorMenuData() { Name = a.Name, ColorBrush = a.Resources["AccentColorBrush"] as Brush })
                                            .ToList();

            // create metro theme color menu items for the demo
            AppThemes = ThemeManager.AppThemes
                                           .Select(a => new AppThemeMenuData() { Name = a.Name, BorderColorBrush = a.Resources["BlackColorBrush"] as Brush, ColorBrush = a.Resources["WhiteColorBrush"] as Brush })
                                           .ToList();

            MiColor.ItemsSource = AccentColors;
            MiTheme.ItemsSource = AppThemes;

            // Устанавливаем текущие. На всякий случай "без ошибок"
            try
            {
                _curTheme = UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "Theme");
                foreach (var item in MiTheme.Items.Cast<AppThemeMenuData>().Where(item => item.Name.Equals(_curTheme)))
                {
                    MiTheme.SelectedIndex = MiTheme.Items.IndexOf(item);
                }

                _curColor = UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "AccentColor");
                foreach (
                    var item in MiColor.Items.Cast<AccentColorMenuData>().Where(item => item.Name.Equals(_curColor)))
                {
                    MiColor.SelectedIndex = MiColor.Items.IndexOf(item);
                }
            }
            catch
            {
                //ignored
            }
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
                var regVariant = UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.User, "RegestryVariant");
                if (!string.IsNullOrEmpty(regVariant))
                {
                    TbAboutRegKey.Visibility = Visibility.Visible;
                    if (regVariant.Equals("0"))
                        TbAboutRegKey.Text = ModPlusAPI.Language.GetItem(LangItem, "h10") + " " +
                                             UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.User, "HDmodel");
                    else if (regVariant.Equals("1"))
                        TbAboutRegKey.Text = ModPlusAPI.Language.GetItem(LangItem, "h11") + " " +
                                             UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.User, "gName");
                }
            }
        }

        private void ChangeWindowTheme()
        {
            //Theme
            try
            {
                ThemeManager.ChangeAppStyle(this.Resources,
                    ThemeManager.Accents.First(
                        x => x.Name.Equals(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "AccentColor"))
                        ),
                    ThemeManager.AppThemes.First(
                        x => x.Name.Equals(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "Theme")))
                    );
                ChangeTitleBrush();
            }
            catch
            {
                //ignored
            }
        }
        /// <summary>Загрузка данных из файла конфигурации которые требуется отобразить в окне</summary>
        private void GetDataFromConfigFile()
        {
            // Separator
            var separator = UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "Separator");
            CbSeparatorSettings.SelectedIndex = string.IsNullOrEmpty(separator) ? 0 : int.Parse(separator);
            // Check updates and new
            ChkEntByBlock.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "EntByBlockOCM"), out bool b) || b; //true
            ChkFastBlocks.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "FastBlocksCM"), out b) || b; //true
            ChkVPtoMS.IsChecked = !bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "VPtoMS"), out b) || b; //true

            // Виды границ окна
            var border = UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "BordersType");
            foreach (ComboBoxItem item in CbWindowsBorders.Items)
            {
                if (item.Tag.Equals(border))
                {
                    CbWindowsBorders.SelectedItem = item; break;
                }
            }
            if (CbWindowsBorders.SelectedIndex == -1) CbWindowsBorders.SelectedIndex = 3;
            _curBordersType = ((ComboBoxItem)CbWindowsBorders.SelectedItem).Tag.ToString();
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
                ChkMpChkDrwsOnMnu.IsChecked = _curDrwsOnMnu = ModPlusAPI.Variables.DrawingsInFloatMenu;
                ChkMpDrawingsAlone.IsChecked = _curDrawingsAlone = ModPlusAPI.Variables.DrawingsFloatMenu;
                // Выбор в выпадающих списках (сворачивать в)
                CbFloatMenuCollapseTo.SelectedIndex = _curFloatMenuCollapseTo = ModPlusAPI.Variables.FloatMenuCollapseTo;
                CbDrawingsCollapseTo.SelectedIndex = _curDrawingsCollapseTo = ModPlusAPI.Variables.DrawingsFloatMenuCollapseTo;
                // Видимость в зависимости от галочек
                ChkMpChkDrwsOnMnu.Visibility =
                    CbFloatMenuCollapseTo.Visibility =
                        TbFloatMenuCollapseTo.Visibility = _curFloatMenu ? Visibility.Visible : Visibility.Collapsed;
                CbDrawingsCollapseTo.Visibility =
                    TbDrawingsCollapseTo.Visibility = _curDrawingsAlone ? Visibility.Visible : Visibility.Collapsed;
                ChkMpPaletteDrawings.Visibility =
                    ChkMpPaletteFunctions.Visibility = _curPalette ? Visibility.Visible : Visibility.Collapsed;
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
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "Separator",
                ((ComboBox)sender).SelectedIndex.ToString(CultureInfo.InvariantCulture), true);
        }
        // Выбор темы
        private void MiTheme_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "Theme", ((AppThemeMenuData)e.AddedItems[0]).Name, true);
            ChangeWindowTheme();
        }
        // Выбор цвета
        private void MiColor_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "AccentColor", ((AccentColorMenuData)e.AddedItems[0]).Name, true);
            ChangeWindowTheme();
        }
        // windows borders select
        private void CbWindowsBorders_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cb = sender as ComboBox;
            if (!(cb?.SelectedItem is ComboBoxItem cbi)) return;
            this.ChangeWindowBordes(cbi.Tag.ToString());
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "BordersType", cbi.Tag.ToString(), true);
        }


        private void MpMainSettings_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!string.IsNullOrEmpty(TbEmailAdress.Text))
            {
                if (IsValidEmail(TbEmailAdress.Text))
                    TbEmailAdress.BorderBrush = FindResource("TextBoxBorderBrush") as Brush;
                else
                {
                    TbEmailAdress.BorderBrush = Brushes.Red;
                    ModPlusAPI.Windows.MessageBox.Show(
                        ModPlusAPI.Language.GetItem(LangItem, "tt4"));
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
                bool needRestartByLang = !((Language.LangItem)CbLanguages.SelectedItem).Name.Equals(_curLang);
                if (needRestartByLang)
                    ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "tt17"));
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
                        // Перегружаем плавающее меню, если изменилась тема, вкл/выкл открытые чертежи, границы, сворачивать в
                        if (!string.IsNullOrEmpty(_curColor) &
                            !string.IsNullOrEmpty(_curTheme) &
                            !string.IsNullOrEmpty(_curBordersType) &
                            !string.IsNullOrEmpty(_curFloatMenuCollapseTo.ToString()) ||
                            needRestartByLang)
                        {
                            if (!UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "Theme").Equals(_curTheme) |
                                !UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "AccentColor").Equals(_curColor) |
                                !UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "BordersType").Equals(_curBordersType) |
                                !UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "FloatMenuCollapseTo").Equals(_curFloatMenuCollapseTo.ToString()) |
                                !ChkMpChkDrwsOnMnu.IsChecked.Value.Equals(_curDrwsOnMnu))
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
                        if (!string.IsNullOrEmpty(_curColor) &
                            !string.IsNullOrEmpty(_curTheme) &
                            !string.IsNullOrEmpty(_curBordersType) &
                            !string.IsNullOrEmpty(_curDrawingsCollapseTo.ToString()) ||
                            needRestartByLang)
                        {
                            if (!UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "Theme").Equals(_curTheme) |
                                !UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "AccentColor").Equals(_curColor) |
                                !UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "BordersType").Equals(_curBordersType) |
                                !UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "DrawingsCollapseTo").Equals(_curDrawingsCollapseTo.ToString()) |
                                !ChkMpDrawingsAlone.IsChecked.Value.Equals(_curDrawingsAlone))
                            {
                                MpDrawingsFunction.MpDrawingsWin.Close();
                                MpDrawingsFunction.LoadMainMenu();
                            }
                        }
                    }
                    else MpDrawingsFunction.LoadMainMenu();
                }
                if (needRestartByLang)
                {
                    RibbonBuilder.RemoveRibbon();
                    RibbonBuilder.BuildRibbon();
                }
                else
                {
                    // Если выключили/включили ленту
                    if (!ChkMpRibbon.IsChecked.Value.Equals(_curRibbon))
                    {
                        if (ChkMpRibbon.IsChecked.Value) RibbonBuilder.BuildRibbon();
                        else RibbonBuilder.RemoveRibbon();
                    }
                }

                // context menues
                MiniFunctions.LoadUnloadContextMenues();
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
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet",
                name.Substring(5),
                chkBox.IsChecked?.ToString(),
                true
                );
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
            UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "ChkQuietLoading",
                (ChkQuietLoading.IsChecked != null && ChkQuietLoading.IsChecked.Value).ToString(), true);
            ModPlusAPI.Variables.QuietLoading = (ChkQuietLoading.IsChecked != null && ChkQuietLoading.IsChecked.Value);
        }
        // Сворачивать в - для плавающего меню
        private void CbFloatMenuCollapseTo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb)
            {
                UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "FloatMenuCollapseTo",
                    cb.SelectedIndex.ToString(CultureInfo.InvariantCulture), true);
                ModPlusAPI.Variables.FloatMenuCollapseTo = cb.SelectedIndex;
            }
        }

        private void CbDrawingsCollapseTo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb)
            {
                UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "MainSet", "DrawingsCollapseTo",
                    cb.SelectedIndex.ToString(CultureInfo.InvariantCulture), true);
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
                    MiniFunctions.ContextMenues.EntByBlockObjectContextMenu.Attach();
                else MiniFunctions.ContextMenues.EntByBlockObjectContextMenu.Detach();
            }
        }
        // Частоиспользуемые блоки
        private void ChkFastBlocks_OnChecked(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox chk)
            {
                UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "FastBlocksCM", (chk.IsChecked != null && chk.IsChecked.Value).ToString(), true);
                if (chk.IsChecked != null && chk.IsChecked.Value)
                    MiniFunctions.ContextMenues.FastBlockContextMenu.Attach();
                else MiniFunctions.ContextMenues.FastBlockContextMenu.Detach();
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
                    MiniFunctions.ContextMenues.VPtoMSobjectContextMenu.Attach();
                else MiniFunctions.ContextMenues.VPtoMSobjectContextMenu.Detach();
            }
        }
        #endregion
        private void TbEmailAdress_OnLostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb)
            {
                if (IsValidEmail(tb.Text))
                    tb.BorderBrush = FindResource("TextBoxBorderBrush") as Brush;
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

        private void TitleWin_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
    }

    public class AccentColorMenuData
    {
        public string Name { get; set; }
        public Brush BorderColorBrush { get; set; }
        public Brush ColorBrush { get; set; }

    }
    public class AppThemeMenuData : AccentColorMenuData
    {
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
