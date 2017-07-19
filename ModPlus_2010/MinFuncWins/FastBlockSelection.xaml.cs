using System.Collections.Generic;
using System.Windows;
using mpMsg;
using mpSettings;

namespace ModPlus.MinFuncWins
{
    /// <summary>
    /// Логика взаимодействия для FastBlockSelection.xaml
    /// </summary>
    public partial class FastBlockSelection
    {
        public List<string> ValidateNames;

        public FastBlockSelection()
        {
            InitializeComponent();
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
                );
        }

        private void BtOk_OnClick(object sender, RoutedEventArgs e)
        {
            if (ValidateNames.Contains(TbBlockName.Text))
            {
                MpMsgWin.Show("Такое имя уже используется! Придумайте другое");
                return;
            }
            if (string.IsNullOrEmpty(TbBlockName.Text))
            {
                MpMsgWin.Show("Нужно указать отображаемое имя блока!");
                return;
            }
            if (LbBlocks.SelectedIndex == -1)
            {
                MpMsgWin.Show("Нужно выбрать блок в списке!");
                return;
            }
            DialogResult = true;
        }

        private void BtCancel_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
