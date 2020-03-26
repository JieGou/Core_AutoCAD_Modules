namespace ModPlus
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
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

    partial class MpDrawings
    {
        // Переменные
        readonly DocumentCollection _docs = AcApp.DocumentManager;

        internal MpDrawings()
        {
            if (double.TryParse(Regestry.GetValue("DrawingsWinTop"), out var top))
                Top = top;
            else
                Top = 180;

            if (double.TryParse(Regestry.GetValue("DrawingsWinLeft"), out var left))
                Left = left;
            else
                Left = 60;

            InitializeComponent();
            ModPlusAPI.Windows.Helpers.WindowHelpers.ChangeStyleForResourceDictionary(Resources);
            ModPlusAPI.Language.SetLanguageProviderForResourceDictionary(Resources);

            MouseEnter += Window_MouseEnter;
            MouseLeave += Window_MouseLeave;
            MouseLeftButtonDown += Window_MouseLeftButtonDown;

            // Подключение обработчиков событий для создания и закрытия чертежей
            AcApp.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;
            AcApp.DocumentManager.DocumentDestroyed += DocumentManager_DocumentDestroyed;
            AcApp.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;

            GetDocuments();

            // Обрабатываем событие покидания мышкой окна
            OnMouseLeaving();
        }

        private void DocumentManager_DocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            GetDocuments();
            CheckUnused();
        }

        private void GetDocuments()
        {
            Drawings.SelectionChanged -= Drawings_SelectionChanged;
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
            finally
            {
                Drawings.SelectionChanged += Drawings_SelectionChanged;
            }
        }

        private void CheckUnused()
        {
            try
            {
                foreach (ListBoxItem item in Drawings.Items)
                {
                    if (item.Content.Equals(Path.GetFileName(_docs.MdiActiveDocument.Name)) &
                        item.ToolTip.Equals(_docs.MdiActiveDocument.Name))
                        Drawings.SelectedItem = item;
                }
            }
            catch
            {
                // ignored
            }
        }

        // Чертеж закрыт
        private void DocumentManager_DocumentDestroyed(object sender, DocumentDestroyedEventArgs e)
        {
            GetDocuments();
            CheckUnused();
        }

        // Документ создан/открыт
        private void DocumentManager_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            GetDocuments();
            CheckUnused();
        }

        // Наведение мышки на окно
        private void Window_MouseEnter(object sender, MouseEventArgs e)
        {
            if (_docs.Count > 0)
            {
                ImgIcon.Visibility = Visibility.Collapsed;
                TbHeader.Visibility = Visibility.Visible;
                ExpOpenDrawings.Visibility = Visibility.Visible;

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
            // icon
            if (Variables.DrawingsFloatMenuCollapseTo.Equals(0))
            {
                ImgIcon.Visibility = Visibility.Visible;
                TbHeader.Visibility = Visibility.Collapsed;
            }

            // header
            else
            {
                ImgIcon.Visibility = Visibility.Collapsed;
                TbHeader.Visibility = Visibility.Visible;
            }

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
    }

    internal static class MpDrawingsFunction
    {
        public static MpDrawings MpDrawingsWin;

        /// <summary>
        /// Загрузка основного меню в зависимости от настроек
        /// </summary>
        public static void LoadMainMenu()
        {
            if (Variables.DrawingsFloatMenu)
            {
                if (MpDrawingsWin == null)
                {
                    MpDrawingsWin = new MpDrawings();
                    MpDrawingsWin.Closed += MpDrawingsWinClosed;
                }

                if (MpDrawingsWin.IsLoaded)
                    return;
                AcApp.ShowModelessWindow(
                    AcApp.MainWindow.Handle, MpDrawingsWin);
            }
            else
            {
                MpDrawingsWin?.Close();
            }
        }

        private static void MpDrawingsWinClosed(object sender, EventArgs e)
        {
            Regestry.SetValue("DrawingsWinTop", MpDrawingsWin.Top.ToString(CultureInfo.InvariantCulture));
            Regestry.SetValue("DrawingsWinLeft", MpDrawingsWin.Left.ToString(CultureInfo.InvariantCulture));
            MpDrawingsWin = null;
        }
    }
}
