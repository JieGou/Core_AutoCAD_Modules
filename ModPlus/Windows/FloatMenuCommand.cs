namespace ModPlus.Windows
{
    using System;
    using System.Globalization;
    using Autodesk.AutoCAD.ApplicationServices.Core;
    using ModPlusAPI;
    using ModPlusAPI.Windows;

    public static class FloatMenuCommand
    {
        public static FloatMenu MainMenuWin;

        /// <summary>
        /// Загрузка основного меню в зависимости от настроек
        /// </summary>
        public static void LoadMainMenu()
        {
            try
            {
                if (Variables.FloatMenu)
                {
                    if (MainMenuWin == null)
                    {
                        MainMenuWin = new FloatMenu();
                        MainMenuWin.Closed += MainMenuWinClosed;
                    }

                    if (MainMenuWin.IsLoaded)
                        return;

                    Application.ShowModelessWindow(Application.MainWindow.Handle, MainMenuWin);
                }
                else
                {
                    MainMenuWin?.Close();
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        static void MainMenuWinClosed(object sender, EventArgs e)
        {
            RegistryUtils.SetValue("FloatingMenuTop", MainMenuWin.Top.ToString(CultureInfo.InvariantCulture));
            RegistryUtils.SetValue("FloatingMenuLeft", MainMenuWin.Left.ToString(CultureInfo.InvariantCulture));
            MainMenuWin = null;
        }
    }
}