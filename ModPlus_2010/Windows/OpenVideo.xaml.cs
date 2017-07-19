using System;
using mpSettings;

namespace ModPlus.Windows
{
    /// <summary>
    /// Логика взаимодействия для OpenVideo.xaml
    /// </summary>
    public partial class OpenVideo
    {
        public OpenVideo(string url)
        {
            InitializeComponent();
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
                );
            Closed += OpenVideo_Closed;
            WebBrowser.Navigate(new Uri(url));
        }

        private void OpenVideo_Closed(object sender, System.EventArgs e)
        {
            WebBrowser.Dispose();
        }
    }
}
