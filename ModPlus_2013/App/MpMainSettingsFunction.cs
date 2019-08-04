namespace ModPlus.App
{
    using Autodesk.AutoCAD.Runtime;
    using ModPlusAPI.Windows;

    /// <summary>
    /// Команда запуска окна настроек
    /// </summary>
    public class MpMainSettingsFunction
    {
        /// <summary>
        /// Запуск окна настроек
        /// </summary>
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