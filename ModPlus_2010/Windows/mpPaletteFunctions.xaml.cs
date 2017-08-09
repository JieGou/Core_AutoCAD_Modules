#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using ModPlus.Helpers;
using ModPlusAPI;
using ModPlusAPI.Windows;

namespace ModPlus.Windows
{
    internal partial class mpPaletteFunctions
    {
        internal mpPaletteFunctions()
        {
            InitializeComponent();
            ModPlusAPI.Windows.Helpers.WindowHelpers.ChangeThemeForResutceDictionary(this.Resources, true);
            Loaded += MpPaletteFunctions_Loaded;
        }

        private void MpPaletteFunctions_Loaded(object sender, RoutedEventArgs e)
        {
            FillFieldsFunction();
            FillFunctions();

        }
        // Заполнение списка функций
        private void FillFunctions()
        {
            try
            {
                // Расположение файла конфигурации
                var confF = UserConfigFile.FullFileName;
                // Грузим
                var configFile = XElement.Load(confF);
                // Проверяем есть ли группа Config
                if (configFile.Element("Config") == null)
                {
                    ModPlusAPI.Windows.MessageBox.Show("Файл конфигурации поврежден! Невозможно заполнить плавающее меню");
                    return;
                }
                var element = configFile.Element("Config");
                // Проверяем есть ли подгруппа Cui
                if (element?.Element("CUI") == null)
                {
                    ModPlusAPI.Windows.MessageBox.Show("Файл конфигурации поврежден! Невозможно заполнить плавающее меню");
                    return;
                }
                var confCuiXel = element.Element("CUI");
                // Проходим по группам
                if (confCuiXel == null) return;

                foreach (var group in confCuiXel.Elements("Group"))
                {
                    var exp = new Expander
                    {
                        Header = group.Attribute("GroupName")?.Value,
                        IsExpanded = false,
                        Margin = new Thickness(1)
                    };
                    var grid = new Grid(){HorizontalAlignment = HorizontalAlignment.Stretch};
                    var index = 0;
                    // Проходим по функциям группы
                    foreach (var func in group.Elements("Function"))
                    {
                        var funcNameAttr = func.Attribute("Name")?.Value;
                        if (string.IsNullOrEmpty(funcNameAttr)) continue;

                        var loadedFunction =
                            LoadFunctionsHelper.LoadedFunctions.FirstOrDefault(x => x.Name.Equals(funcNameAttr));
                        if (loadedFunction == null) continue;
                        grid.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
                        var btn = WPFMenuesHelper.AddButton(this, loadedFunction.Name, loadedFunction.LName,
                            loadedFunction.BigIconUrl, loadedFunction.Description,
                            loadedFunction.FullDescription, loadedFunction.ToolTipHelpImage, false);
                        
                        btn.SetValue(Grid.RowProperty, index);
                        grid.Children.Add(btn);

                        index++;
                        
                        if (loadedFunction.SubFunctionsNames.Any())
                        {
                            for (int i = 0; i < loadedFunction.SubFunctionsNames.Count; i++)
                            {
                                btn = WPFMenuesHelper.AddButton(this,
                                    loadedFunction.SubFunctionsNames[i],
                                    loadedFunction.SubFunctionsLNames[i], loadedFunction.SubBigIconsUrl[i],
                                    loadedFunction.SubDescriptions[i], loadedFunction.SubFullDescriptions[i],
                                    loadedFunction.SubHelpImages[i], false);
                                grid.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
                                btn.SetValue(Grid.RowProperty, index);
                                grid.Children.Add(btn);
                                index++;
                            }
                        }

                        foreach (var subFunc in func.Elements("SubFunction"))
                        {
                            var subFuncNameAttr = subFunc.Attribute("Name")?.Value;
                            if (string.IsNullOrEmpty(subFuncNameAttr)) continue;
                            var loadedSubFunction =
                                LoadFunctionsHelper.LoadedFunctions.FirstOrDefault(x => x.Name.Equals(subFuncNameAttr));
                            if (loadedSubFunction == null) continue;
                            btn = WPFMenuesHelper.AddButton(this, loadedSubFunction.Name, loadedSubFunction.LName,
                                loadedSubFunction.BigIconUrl, loadedSubFunction.Description,
                                loadedSubFunction.FullDescription, loadedSubFunction.ToolTipHelpImage, false);
                            grid.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
                            btn.SetValue(Grid.RowProperty, index);
                            grid.Children.Add(btn);
                            index++;
                        }
                    }
                    exp.Content = grid;
                    // Добавляем группу, если заполнились функции!
                    if (grid.Children.Count > 0)
                        FunctionsPanel.Children.Add(exp);
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.ShowForConfigurator(exception);
            }
        }

        private void FillFieldsFunction()
        {
            if (LoadFunctionsHelper.HasmpStampsFunction(out string icon))
            {
                BtFields.Visibility = Visibility.Visible;
                try
                {
                    var bitmapImage = new BitmapImage(new Uri(icon, UriKind.RelativeOrAbsolute));
                    var img = new Image
                    {
                        Source = bitmapImage,
                        Stretch = Stretch.Uniform,
                        Width = 16,
                        Height = 16,
                        SnapsToDevicePixels = true
                    };
                    BtFields.Content = img;
                }
                catch
                {
                    // ignored
                }
            }
            else
            {
                BtFields.Visibility = Visibility.Collapsed;
            }
        }

        // Заполнение списка функций
        //private void FillFunctions()
        //{
        //    try
        //    {
        //        // Расположение файла конфигурации
        //        var confF = MpSettings.FullFileName;
        //        // Грузим
        //        var configFile = XElement.Load(confF);
        //        // Проверяем есть ли группа Config
        //        if (configFile.Element("Config") == null)
        //        {
        //            MpMsgWin.Show("Файл конфигурации поврежден! Невозможно заполнить плавающее меню");
        //            return;
        //        }
        //        var element = configFile.Element("Config");
        //        // Проверяем есть ли подгруппа Functions
        //        if (element?.Element("Functions") == null)
        //        {
        //            MpMsgWin.Show("Файл конфигурации поврежден! Невозможно заполнить плавающее меню");
        //            return;
        //        }
        //        var confFuncXel = element.Element("Functions");
        //        // Проверяем есть ли подгруппа Cui
        //        if (element.Element("CUI") == null)
        //        {
        //            MpMsgWin.Show("Файл конфигурации поврежден! Невозможно заполнить плавающее меню");
        //            return;
        //        }
        //        var confCuiXel = element.Element("CUI");
        //        // Проходим по группам
        //        if (confCuiXel == null) return;
        //        foreach (var group in confCuiXel.Elements("Group"))
        //        {
        //            var exp = new Expander
        //            {
        //                Header = group.Attribute("GroupName")?.Value,
        //                IsExpanded = false,
        //                Margin = new Thickness(1, 0, 1, 1)
        //            };
        //            var expStck = new StackPanel { Orientation = Orientation.Vertical };

        //            // Проходим по функциям группы
        //            foreach (var func in group.Elements("Function"))
        //            {
        //                if (LoadFunctionsHelper.HasFunction(func, confFuncXel, out XElement f, out string fileVersion))
        //                {
        //                    expStck.Children.Add(
        //                        AddButton(
        //                            func.Attribute("Name")?.Value,
        //                            func.Attribute("LName")?.Value,
        //                            LoadFunctionsHelper.GetBigIconUriString(func.Attribute("Name")?.Value,
        //                                LoadFunctionsHelper.GetFunctionVersion(fileVersion, f)),
        //                            func.Attribute("Description")?.Value
        //                        ));
        //                }
        //                foreach (var subFunc in func.Elements("SubFunction"))
        //                {
        //                    if (LoadFunctionsHelper.HasFunction(subFunc, confFuncXel, out f, out fileVersion))
        //                    {
        //                        expStck.Children.Add(
        //                            AddButton(
        //                                subFunc.Attribute("Name")?.Value,
        //                                subFunc.Attribute("LName")?.Value,
        //                                LoadFunctionsHelper.GetBigIconUriString(subFunc.Attribute("Name")?.Value,
        //                                    LoadFunctionsHelper.GetFunctionVersion(fileVersion, f)),
        //                                subFunc.Attribute("Description")?.Value
        //                            ));
        //                    }
        //                }
        //            }
        //            exp.Content = expStck;
        //            // Добавляем группу, если заполнились функции!
        //            if (expStck.Children.Count > 0)
        //                FunctionsPanel.Children.Add(exp);
        //        }
        //    }
        //    catch (Exception exception) { ExceptionBox.ShowForConfigurator(exception); }
        //}

        //private Button AddButton(string name, string lname, string img32, string description)
        //{
        //    var brd = new Border
        //    {
        //        Padding = new Thickness(1),
        //        Margin = new Thickness(1),
        //        Background = Brushes.White
        //    };
        //    try
        //    {
        //        var bitmapImage = new BitmapImage(new Uri(img32, UriKind.RelativeOrAbsolute));
        //        var img = new Image
        //        {
        //            Source = bitmapImage,
        //            Stretch = Stretch.Uniform,
        //            Width = 32,
        //            Height = 32
        //        };
        //        brd.Child = img;
        //    }
        //    catch
        //    {
        //        // ignored
        //    }
        //    var txt = new TextBlock
        //    {
        //        VerticalAlignment = VerticalAlignment.Center,
        //        TextWrapping = TextWrapping.Wrap,
        //        Width = 150,
        //        Text = lname,
        //        Margin = new Thickness(3, 0, 1, 0)
        //    };
        //    var stck = new StackPanel
        //    {
        //        Orientation = Orientation.Horizontal,
        //        HorizontalAlignment = HorizontalAlignment.Left
        //    };
        //    stck.Children.Add(brd);
        //    stck.Children.Add(txt);
        //    var btn = new Button
        //    {
        //        Name = name,
        //        Content = stck,
        //        ToolTip = description,
        //        Margin = new Thickness(1),
        //        Padding = new Thickness(1),
        //        Style = FindResource("BtStyle") as Style
        //    };
        //    btn.Click += CommandButtonClick;

        //    return btn;
        //}
        // Обработка запуска функций        
        //private static void CommandButtonClick(object sender, RoutedEventArgs e)
        //{
        //    if (AcApp.DocumentManager.Count > 0)
        //        AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_" + ((Button)sender).Name + " ", false, false, false);
        //}
        // настройки modplus
        private void BtSettings_OnClick(object sender, RoutedEventArgs e)
        {
            if (AcApp.DocumentManager.Count > 0)
                AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_MPSETTINGS ", false, false, false);
        }
        // start fields function
        private void BtFields_OnClick(object sender, RoutedEventArgs e)
        {
            if (AcApp.DocumentManager.Count > 0)
                AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_MPSTAMPFIELDS ", false, false, false);
        }
    }
}
