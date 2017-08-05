#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Globalization;
using System.Linq;
using Autodesk.Windows;
using System.Windows.Controls;

namespace ModPlus.Helpers
{
    internal static class RibbonHelpers
    {
        /// <summary>
        /// Создание маленькой кнопки
        /// </summary>
        /// <param name="fName">Навание функции (= параметр запуска функции)</param>
        /// <param name="lName">Локальное название функции</param>
        /// <param name="img">Иконка</param>
        /// <param name="description">Описание функции</param>
        /// <param name="fullDescription"></param>
        /// <param name="helpImage"></param>
        /// <returns></returns>
        public static RibbonButton AddSmallButton(string fName, string lName, string img, string description, string fullDescription, string helpImage)
        {
            try
            {
                var tt = new RibbonToolTip
                {
                    IsHelpEnabled = false,
                    Content = description,
                    Command = fName,
                    IsProgressive = true
                };
                if (!string.IsNullOrEmpty(fullDescription))
                    tt.ExpandedContent = fullDescription;
                try
                {
                    if(!string.IsNullOrEmpty(helpImage))
                        tt.ExpandedImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(helpImage, UriKind.RelativeOrAbsolute));
                }
                catch
                {
                    // ignored
                }
                var ribBtn = new RibbonButton
                {
                    CommandParameter = tt.Command = fName,
                    Name = tt.Title = lName,
                    CommandHandler = new RibbonCommandHandler(),
                    Orientation = Orientation.Horizontal,
                    Size = RibbonItemSize.Standard,
                    ShowImage = true,
                    ShowText = false,
                    ToolTip = tt
                };
                try
                {
                    if (!string.IsNullOrEmpty(img))
                        ribBtn.Image = new System.Windows.Media.Imaging.BitmapImage(new Uri(img, UriKind.RelativeOrAbsolute));
                }
                catch
                {
                    // ignored
                }
                return ribBtn;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Создание большой кнопки
        /// </summary>
        /// <param name="fName">Навание функции (= параметр запуска функции)</param>
        /// <param name="lName">Локальное название функции</param>
        /// <param name="img">Иконка</param>
        /// <param name="description">Описание функции</param>
        /// <param name="orientation">Ориентация кнопки: горизонтальная или вертикальная</param>
        /// <param name="fullDescription"></param>
        /// <param name="helpImage"></param>
        /// <returns></returns>
        public static RibbonButton AddBigButton(string fName, string lName, string img, string description, Orientation orientation, string fullDescription, string helpImage)
        {
            try
            {
                var tt = new RibbonToolTip
                {
                    IsHelpEnabled = false,
                    Content = description,
                    Command = fName,
                    IsProgressive = true
                };
                if (!string.IsNullOrEmpty(fullDescription))
                    tt.ExpandedContent = fullDescription;
                try
                {
                    if (!string.IsNullOrEmpty(helpImage))
                        tt.ExpandedImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(helpImage, UriKind.RelativeOrAbsolute));
                }
                catch 
                {
                    // ignored
                }
                var ribBtn = new RibbonButton
                {
                    CommandParameter = tt.Command = fName,
                    Name = tt.Title = lName,
                    Text = ConvertLName(lName),
                    CommandHandler = new RibbonCommandHandler(),
                    Orientation = orientation,
                    Size = RibbonItemSize.Large,
                    ShowImage = true,
                    ShowText = true,
                    ToolTip = tt
                };
                try
                {
                    if (!string.IsNullOrEmpty(img))
                        ribBtn.LargeImage =
                            new System.Windows.Media.Imaging.BitmapImage(new Uri(img, UriKind.RelativeOrAbsolute));
                }
                catch
                {
                    // ignored
                }
                return ribBtn;
            }
            catch (Exception)
            {
                return null;
            }
        }
        public static RibbonButton AddButton(string fName, string lName, string img16, string img32, string description, Orientation orientation, string fullDescription, string helpImage)
        {
            try
            {
                var tt = new RibbonToolTip
                {
                    IsHelpEnabled = false,
                    Content = description,
                    Command = fName,
                    IsProgressive = true
                };
                if (!string.IsNullOrEmpty(fullDescription))
                    tt.ExpandedContent = fullDescription;
                try
                {
                    if (!string.IsNullOrEmpty(helpImage))
                        tt.ExpandedImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(helpImage, UriKind.RelativeOrAbsolute));
                }
                catch
                {
                    // ignored
                }
                var ribBtn = new RibbonButton
                {
                    CommandParameter = tt.Command = fName,
                    Name = tt.Title = lName,
                    Text = ConvertLName(lName),
                    CommandHandler = new RibbonCommandHandler(),
                    Orientation = orientation,
                    Size = RibbonItemSize.Large,
                    ShowImage = true,
                    ShowText = true,
                    ToolTip = tt
                };
                try
                {
                    if (!string.IsNullOrEmpty(img16))
                        ribBtn.Image = new System.Windows.Media.Imaging.BitmapImage(new Uri(img16, UriKind.RelativeOrAbsolute));
                    if (!string.IsNullOrEmpty(img32))
                        ribBtn.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(img32, UriKind.RelativeOrAbsolute));
                }
                catch
                {
                    // ignored
                }
                return ribBtn;
            }
            catch (Exception)
            {
                return null;
            }
        }
        public class RibbonCommandHandler : System.Windows.Input.ICommand
        {
            public bool CanExecute(object parameter)
            {
                return true;
            }
#pragma warning disable 67
            public event EventHandler CanExecuteChanged;
#pragma warning restore 67

            public void Execute(object parameter)
            {
                // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
                if (parameter is RibbonButton)
                {
                    var button = (RibbonButton)parameter;
                    AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute(
                        button.CommandParameter + " ", true, false, true);
                }
            }
        }

        private static string ConvertLName(string lName)
        {
            if (!lName.Contains(" ")) return lName;
            if (lName.Length <= 8) return lName;
            if (lName.Count(x => x == ' ') == 1)
            {
                return lName.Split(' ')[0] + Environment.NewLine + lName.Split(' ')[1];
            }
            else
            {
                var lst = lName.Split(' ');
                var newLName = string.Empty;
                int center = (int)(lst.Length / 2.0);
                for (int i = 0; i < lst.Length; i++)
                {
                    var s = lst[i];
                    if (i != center) newLName = newLName + s;
                    else newLName = newLName + s + Environment.NewLine;
                }
                return newLName;
            }
        }
    }
}
