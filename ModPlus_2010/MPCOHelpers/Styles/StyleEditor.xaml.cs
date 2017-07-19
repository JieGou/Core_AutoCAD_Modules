using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using mpMsg;
using mpSettings;
using Microsoft.Win32;

namespace ModPlus.MPCOHelpers.Styles
{
    /// <summary>
    /// Логика взаимодействия для StyleEditor.xaml
    /// </summary>
    public partial class StyleEditor
    {
        public StyleEditor()
        {
            InitializeComponent();
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
            );
            Loaded += StyleEditor_OnLoaded;
            MouseLeftButtonDown += StyleEditor_OnMouseLeftButtonDown;
            PreviewKeyDown += StyleEditor_OnPreviewKeyDown;
            ContentRendered += StyleEditor_ContentRendered;
        }

        #region Window work
        private void StyleEditor_OnLoaded(object sender, RoutedEventArgs e)
        {
            SizeToContent = SizeToContent.Manual;
        }

        private void StyleEditor_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void StyleEditor_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) Close();
        }
        #endregion

        private List<StyleToBind> Styles;
        private void StyleEditor_ContentRendered(object sender, EventArgs e)
        {
            try
            {
                Styles = new List<StyleToBind>();
                GetSylesFromFiles();
                TvStyles.ItemsSource = Styles;
            }
            catch (Exception exception)
            {
                MpExWin.Show(exception);
            }
        }

        private void GetSylesFromFiles()
        {
            // Получаем системные стили
            var systemStyles = Helpers.GetSystemStyle();
            if (systemStyles.Any())
                foreach (MPCOstyle systemStyle in systemStyles)
                {
                    StyleToBind styleToBind = new StyleToBind();
                    styleToBind.StyleName = systemStyle.Name;
                    styleToBind.Styles.Add(systemStyle);
                    Styles.Add(styleToBind);
                }
        }

        private void TvStyles_OnSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var item = e.NewValue;
            if (item is MPCOstyle)
            {
                PropertyGrid.SelectedObject = item;
            }
        }
    }
    
    class StyleToBind
    {
        public StyleToBind()
        {
            Styles = new List<MPCOstyle>();
        }
        public string StyleName { get; set; }
        public List<MPCOstyle> Styles { get; set; }
    }
}
