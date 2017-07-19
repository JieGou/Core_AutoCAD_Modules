using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using mpSettings;

namespace ModPlus.Windows
{
    /// <summary>
    /// Логика взаимодействия для Reclame.xaml
    /// </summary>
    public partial class Reclame
    {
        // Для доступа
        public TextBlock TbPricePub
        {
            get { return TbPrice; }
            set { TbPrice = value; }
        }
        // Timer
        readonly DispatcherTimer _timer = new DispatcherTimer();
        private int _count;
        public Reclame()
        {
            InitializeComponent();
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
                );
            // Timer
            _timer.Tick += OnTimer;
            _timer.Interval = TimeSpan.FromMilliseconds(1000);
            this.BtClose.IsEnabled = false;
            _count = 3;
            _timer.Start();
            // Images
            Background = new ImageBrush(
                    new BitmapImage(
                        new Uri("pack://application:,,,/ModPlus_" + MpVersionData.CurCadVers +
                                ";component/Resources/Reclame.jpg")))
            { Stretch = Stretch.None };
            var logo = new BitmapImage();
            logo.BeginInit();
            logo.UriSource = new Uri("pack://application:,,,/ModPlus_" + MpVersionData.CurCadVers + ";component/Resources/Logo.png");
            logo.EndInit();
            ImgLogo.Source = logo;
        }
        protected virtual void OnTimer(object source, EventArgs e)
        {
            _count--;

            this.BtClose.Content = "Закрыть через: " + _count;

            if (_count == 0)
            {
                BtClose.Content = "Закрыть";
                BtClose.IsEnabled = true;
                _timer.IsEnabled = false;
            }
        }

        private void BtClose_OnClick(object sender, RoutedEventArgs e)
        {
                this.Close();
        }
    }
}
