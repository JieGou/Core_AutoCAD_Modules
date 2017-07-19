#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;
using mpMsg;
using RibbonPanelSource = Autodesk.Windows.RibbonPanelSource;
using RibbonRowPanel = Autodesk.Windows.RibbonRowPanel;
using System.Windows.Controls;
using ModPlus.Helpers;

namespace ModPlus.MPCOHelpers
{
    public class MPCORibbonBuilder
    {
        public static void BuildRibbon()
        {
            if (!IsLoaded())
            {
                CreateRibbon();
                AcApp.SystemVariableChanged += acadApp_SystemVariableChanged;
            }
        }
        private static bool IsLoaded()
        {
            var loaded = false;
            var ribCntrl = ComponentManager.Ribbon;
            foreach (var tab in ribCntrl.Tabs)
            {
                if (tab.Id.Equals("ModPlus_ESKD") & tab.Title.Equals("ModPlus ЕСКД"))
                    loaded = true;
                else loaded = false;
            }
            return loaded;
        }
        public static void RemoveRibbon()
        {
            try
            {
                if (IsLoaded())
                {
                    var ribCntrl = ComponentManager.Ribbon;
                    foreach (var tab in ribCntrl.Tabs.Where(
                        tab => tab.Id.Equals("ModPlus_ESKD") & tab.Title.Equals("ModPlus ЕСКД")))
                    {
                        ribCntrl.Tabs.Remove(tab);
                        AcApp.SystemVariableChanged -= acadApp_SystemVariableChanged;
                        break;
                    }
                }
            }
            catch (Exception exception)
            {
                MpExWin.Show(exception);
            }
        }
        static void acadApp_SystemVariableChanged(object sender, SystemVariableChangedEventArgs e)
        {
            if (e.Name.Equals("WSCURRENT")) BuildRibbon();
        }
        private static void CreateRibbon()
        {
            try
            {
                var ribCntrl = ComponentManager.Ribbon;
                // add the tab
                var ribTab = new RibbonTab { Title = "ModPlus ЕСКД", Id = "ModPlus_ESKD" };
                ribCntrl.Tabs.Add(ribTab);
                // add content
                AddPanels(ribTab);
                // add settings panel
                AddSettingsPanel(ribTab);
                ////////////////////////
                ribCntrl.UpdateLayout();
            }
            catch (Exception exception)
            {
                MpExWin.Show(exception);
            }
        }

        private static void AddPanels(RibbonTab ribTab)
        {
            // Линии
            // create the panel source
            var ribSourcePanel = new RibbonPanelSource { Title = "Линии" };
            // now the panel
            var ribPanel = new RibbonPanel { Source = ribSourcePanel };
            ribTab.Panels.Add(ribPanel);

            var ribRowPanel = new RibbonRowPanel();
            // mpBreakLine
            if (LoadFunctionsHelper.LoadedFunctions.Any(x => x.Name.Equals("mpBreakLine")))
            {
                var lf = LoadFunctionsHelper.LoadedFunctions.First(x => x.Name.Equals("mpBreakLine"));
                // Создаем SplitButton
                var risSplitBtn = new RibbonSplitButton
                {
                    Text = "RibbonSplitButton",
                    Orientation = Orientation.Vertical,
                    Size = RibbonItemSize.Large,
                    ShowImage = true,
                    ShowText = true,
                    ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton,
                    ResizeStyle = RibbonItemResizeStyles.NoResize,
                    ListStyle = RibbonSplitButtonListStyle.List
                };
                // Добавляем в него первую функцию, которую делаем основной
                var ribBtn = RibbonHelpers.AddBigButton(lf.Name, lf.LName, lf.BigIconUrl, lf.Description, Orientation.Vertical, lf.FullDescription, lf.ToolTipHelpImage);
                risSplitBtn.Items.Add(ribBtn);
                risSplitBtn.Current = ribBtn;
                // Затем добавляем подфункции
                for (int i = 0; i < lf.SubFunctionsNames.Count; i++)
                {
                    risSplitBtn.Items.Add(RibbonHelpers.AddBigButton(
                        lf.SubFunctionsNames[i], lf.SubFunctionsLNames[i], lf.SubBigIconsUrl[i], lf.SubDescriptions[i], Orientation.Vertical, lf.SubFullDescriptions[i], lf.SubHelpImages[i]
                        ));
                }
                ribRowPanel.Items.Add(risSplitBtn);
            }
            if (ribRowPanel.Items.Any())
            {
                ribSourcePanel.Items.Add(ribRowPanel);
            }
        }
        private static void AddSettingsPanel(RibbonTab ribTab)
        {
            // create the panel source
            var ribSourcePanel = new RibbonPanelSource
            {
                Title = "Настройки"
            };
            // now the panel
            var ribPanel = new RibbonPanel
            {
                Source = ribSourcePanel
            };
            ribTab.Panels.Add(ribPanel);

            var ribRowPanel = new RibbonRowPanel();

            ribRowPanel.Items.Add(
                RibbonHelpers.AddBigButton(
                    "mpSettings",
                    "Настройки",
                    "pack://application:,,,/Modplus_" + MpVersionData.CurCadVers + ";component/Resources/HelpBt.png",
                    "Настройки ModPlus", Orientation.Vertical, "", ""
                ));
            ribSourcePanel.Items.Add(ribRowPanel);
        }
    }
}
