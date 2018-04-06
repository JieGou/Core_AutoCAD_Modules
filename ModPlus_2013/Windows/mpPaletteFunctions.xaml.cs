#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;
using ModPlus.Helpers;
using ModPlusAPI;
using ModPlusAPI.Windows;

namespace ModPlus.Windows
{
    // ReSharper disable once InconsistentNaming
    internal partial class mpPaletteFunctions
    {
        private const string LangItem = "AutocadDlls";

        internal mpPaletteFunctions()
        {
            InitializeComponent();
            ModPlusAPI.Windows.Helpers.WindowHelpers.ChangeThemeForResurceDictionary(Resources, true);
            ModPlusAPI.Language.SetLanguageProviderForWindow(Resources);
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
                XElement configFile;
                using (FileStream fs = new FileStream(confF, FileMode.Open, FileAccess.Read, FileShare.None))
                    configFile = XElement.Load(fs);
                // Проверяем есть ли группа Config
                if (configFile.Element("Config") == null)
                {
                    ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err7"));
                    return;
                }
                var element = configFile.Element("Config");
                // Проверяем есть ли подгруппа Cui
                if (element?.Element("CUI") == null)
                {
                    ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err7"));
                    return;
                }
                var confCuiXel = element.Element("CUI");
                // Проходим по группам
                if (confCuiXel == null) return;

                foreach (var group in confCuiXel.Elements("Group"))
                {
                    var exp = new Expander
                    {
                        Header = ModPlusAPI.Language.TryGetCuiLocalGroupName(group.Attribute("GroupName")?.Value),
                        IsExpanded = false,
                        Margin = new Thickness(1)
                    };
                    var grid = new Grid { HorizontalAlignment = HorizontalAlignment.Stretch };
                    var index = 0;
                    // Проходим по функциям группы
                    foreach (var func in group.Elements("Function"))
                    {
                        var funcNameAttr = func.Attribute("Name")?.Value;
                        if (string.IsNullOrEmpty(funcNameAttr)) continue;

                        var loadedFunction =
                            LoadFunctionsHelper.LoadedFunctions.FirstOrDefault(x => x.Name.Equals(funcNameAttr));
                        if (loadedFunction == null) continue;
                        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                        var btn = WPFMenuesHelper.AddButton(this,
                            loadedFunction.Name,
                            ModPlusAPI.Language.GetFunctionLocalName(loadedFunction.Name, loadedFunction.LName),
                            loadedFunction.BigIconUrl,
                            ModPlusAPI.Language.GetFunctionShortDescrition(loadedFunction.Name, loadedFunction.Description),
                            ModPlusAPI.Language.GetFunctionFullDescription(loadedFunction.Name, loadedFunction.FullDescription),
                            loadedFunction.ToolTipHelpImage, false);

                        btn.SetValue(Grid.RowProperty, index);
                        grid.Children.Add(btn);

                        index++;

                        if (loadedFunction.SubFunctionsNames.Any())
                        {
                            for (int i = 0; i < loadedFunction.SubFunctionsNames.Count; i++)
                            {
                                btn = WPFMenuesHelper.AddButton(this,
                                    loadedFunction.SubFunctionsNames[i],
                                    ModPlusAPI.Language.GetFunctionLocalName(loadedFunction.Name, loadedFunction.SubFunctionsLNames[i], i + 1),
                                    loadedFunction.SubBigIconsUrl[i],
                                    ModPlusAPI.Language.GetFunctionShortDescrition(loadedFunction.Name, loadedFunction.SubDescriptions[i], i + 1),
                                    ModPlusAPI.Language.GetFunctionFullDescription(loadedFunction.Name, loadedFunction.SubFullDescriptions[i], i + 1),
                                    loadedFunction.SubHelpImages[i], false);
                                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
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
                            btn = WPFMenuesHelper.AddButton(this,
                                loadedSubFunction.Name,
                                ModPlusAPI.Language.GetFunctionLocalName(loadedSubFunction.Name, loadedSubFunction.LName),
                                loadedSubFunction.BigIconUrl,
                                ModPlusAPI.Language.GetFunctionShortDescrition(loadedSubFunction.Name, loadedSubFunction.Description),
                                ModPlusAPI.Language.GetFunctionFullDescription(loadedSubFunction.Name, loadedSubFunction.FullDescription),
                                loadedSubFunction.ToolTipHelpImage, false);
                            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
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
                ExceptionBox.Show(exception);
            }
        }

        private void FillFieldsFunction()
        {
            BtFields.Visibility = LoadFunctionsHelper.HasmpStampsFunction(out string _) ? Visibility.Visible : Visibility.Collapsed;
        }

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

        private void BtHideProductIcon_OnClick(object sender, RoutedEventArgs e)
        {
            if (AcApp.DocumentManager.Count > 0)
                AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_MPHIDEPRODUCTICONS ", false, false, false);
        }

        private void BtShowProductIcon_OnClick(object sender, RoutedEventArgs e)
        {
            if (AcApp.DocumentManager.Count > 0)
                AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_MPSHOWPRODUCTICONS ", false, false, false);
        }
    }
}
