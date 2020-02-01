namespace ModPlus.App
{
    using Autodesk.AutoCAD.Runtime;
    using ModPlusAPI.Windows;

    /// <summary>
    /// Данные пользователя
    /// </summary>
    public class UserInfoCommand
    {
        /// <summary>
        /// Запустить окно "Личный кабинет"
        /// </summary>
        [CommandMethod("ModPlus", "mpUserInfo", CommandFlags.Modal)]
        public void ShowUserInfo()
        {
            try
            {
                ModPlusAPI.UserInfo.UserInfoService.ShowUserInfo();
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }
    }
}
