namespace ModPlus.Windows
{
    using System;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using Helpers;
    using ModPlusAPI.Windows;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

    // ReSharper disable once InconsistentNaming
    internal partial class mpPaletteFunctions
    {
        internal mpPaletteFunctions()
        {
            InitializeComponent();
            ModPlusStyle.ThemeManager.ChangeTheme(Resources, "LightBlue");
            ModPlusAPI.Language.SetLanguageProviderForResourceDictionary(Resources);
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
                var confCuiXel = ModPlusAPI.RegistryData.Adaptation.GetCuiAsXElement("AutoCAD");

                // Проходим по группам
                if (confCuiXel == null)
                    return;

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
                        if (string.IsNullOrEmpty(funcNameAttr))
                            continue;

                        var loadedFunction =
                            LoadFunctionsHelper.LoadedFunctions.FirstOrDefault(x => x.Name.Equals(funcNameAttr));
                        if (loadedFunction == null)
                            continue;
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
                                btn = WPFMenuesHelper.AddButton(
                                    this,
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
                            if (string.IsNullOrEmpty(subFuncNameAttr))
                                continue;
                            var loadedSubFunction =
                                LoadFunctionsHelper.LoadedFunctions.FirstOrDefault(x => x.Name.Equals(subFuncNameAttr));
                            if (loadedSubFunction == null)
                                continue;
                            btn = WPFMenuesHelper.AddButton(
                                this,
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
            BtFields.Visibility = LoadFunctionsHelper.HasStampsPlugin(1, out _) ? Visibility.Visible : Visibility.Collapsed;
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

        private void BtUserInfo_OnClick(object sender, RoutedEventArgs e)
        {
            if (AcApp.DocumentManager.Count > 0)
                AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_MPUSERINFO ", false, false, false);
        }
    }
}
