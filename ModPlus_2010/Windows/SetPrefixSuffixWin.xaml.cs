using System.Windows;
using System.Windows.Controls;
using mpSettings;

namespace ModPlus.Windows
{
    /// <summary>
    /// Логика взаимодействия для SetPrefixSuffixWin.xaml
    /// </summary>
    public partial class SetPrefixSuffixWin
    {
        // Для доступа из других сборок
        public TextBox TbPrefixPub
        {
            get { return TbPrefix; }
            set { TbPrefix = value; }
        }

        public TextBox TbSuffixPub
        {
            get { return TbSuffix; }
            set { TbSuffix = value; }
        }

        public SetPrefixSuffixWin()
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
