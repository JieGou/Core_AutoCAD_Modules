using System.IO;
using System.Reflection;
using System.Xml.Linq;

namespace ModPlus.MPCOHelpers
{
    public static class MPCOInitialize
    {
        /// <summary>
        /// Имя регистрируемого приложения, помещаемое в XData
        /// Для всех СПДС примитивов одинаково. Примитивы будут различаться 
        /// по коду 1000 - строка, в которую будет записываться название примитива
        /// </summary>
        //public const string AppName = "ModPlus.SPDS";

        public static string StylesPath;

        public static string SystemStylesFile;
        /// <summary>
        /// Инициализация 
        /// </summary>
        public static void StartUpInitialize()
        {
            // ReSharper disable once AssignNullToNotNullAttribute
            var curDir = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
            // ReSharper disable once AssignNullToNotNullAttribute
            var mpcoPath = Path.Combine(curDir, "UserData");
            if (!Directory.Exists(mpcoPath))
                Directory.CreateDirectory(mpcoPath);
            var mpcoStylesPath = Path.Combine(mpcoPath, "Styles");
            if (!Directory.Exists(mpcoStylesPath))
                Directory.CreateDirectory(mpcoStylesPath);
            // Создаем специальный файл хранения системных стилей
            var systemStylesFile = Path.Combine(mpcoStylesPath, "mpStyles.mpsf");
            if (!File.Exists(systemStylesFile))
            {
                var sysXel = new XElement("SystemStyles");
                sysXel.Save(systemStylesFile);
            }
            else
            {
                try
                {
                    XElement.Load(systemStylesFile);
                }
                catch
                {
                    var sysXel = new XElement("SystemStyles");
                    sysXel.Save(systemStylesFile);
                }
            }
            // set public parameter
            StylesPath = mpcoStylesPath;
            SystemStylesFile = systemStylesFile;
        }
    }
}
