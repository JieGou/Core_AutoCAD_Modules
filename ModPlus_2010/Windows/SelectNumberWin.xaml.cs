using System.Windows;
using System.Windows.Controls;
using ModPlusAPI.Windows;

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
            this.OnWindowStartUp();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
