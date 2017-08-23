#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.AutoCAD.ApplicationServices;
using Path = System.IO.Path;

namespace ModPlus.Windows
{
    internal partial class mpPaletteDrawings 
    {
        // Переменные
        private readonly DocumentCollection Docs = AcApp.DocumentManager;
        
        internal mpPaletteDrawings()
        {
            InitializeComponent();
            ModPlusAPI.Windows.Helpers.WindowHelpers.ChangeThemeForResurceDictionary(this.Resources, true);
            Loaded += MpPaletteDrawings_Loaded;
        }

        private void MpPaletteDrawings_Loaded(object sender, RoutedEventArgs e)
        {
            // Подключение обработчиков событий для создания и закрытия чертежей
            AcApp.DocumentManager.DocumentCreated +=
                DocumentManager_DocumentCreated;
            AcApp.DocumentManager.DocumentDestroyed +=
                DocumentManager_DocumentDestroyed;
            AcApp.DocumentManager.DocumentActivated +=
                DocumentManager_DocumentActivated;
            //////////////////////////////
            GetDocuments();
        }
        // документ стал активным
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
                    if (item.Content.Equals(Path.GetFileName(Docs.MdiActiveDocument.Name)) &
                        item.ToolTip.Equals(Docs.MdiActiveDocument.Name))
                        Drawings.SelectedItem = item;
                }
            }
            catch
            {
                //ignored
            }
        }
        // Чертеж закрыт
        void DocumentManager_DocumentDestroyed(object sender, DocumentDestroyedEventArgs e)
        {
            GetDocuments();
            CheckUnused();
        }
        // Документ создан/открыт
        void DocumentManager_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            GetDocuments();
            CheckUnused();
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

        private void BtAddDrawing_OnClick(object sender, RoutedEventArgs e)
        {
            if(Docs.Count > 0)
                AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_NEW ", true, false, false);
        }

        private void BtOpenDrawing_OnClick(object sender, RoutedEventArgs e)
        {
            if (Docs.Count > 0)
                AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_OPEN ", true, false, false);
        }
    }
}
