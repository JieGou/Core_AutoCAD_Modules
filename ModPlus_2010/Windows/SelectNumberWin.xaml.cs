using System.Windows;
using System.Windows.Controls;
using mpSettings;

namespace ModPlus.Windows
{
    /// <summary>
    /// Логика взаимодействия для SelectNumberWin.xaml
    /// </summary>
    public partial class SelectNumberWin
    {
        // Для доступа
        public ListBox LbNumbersPub
        {
            get { return LbNumbers; }
            set { LbNumbers = value; }
        }

        public SelectNumberWin()
        {
            InitializeComponent();
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
                );
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
