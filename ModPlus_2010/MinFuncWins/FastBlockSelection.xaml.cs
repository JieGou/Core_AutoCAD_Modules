using System.Collections.Generic;
using System.Windows;
using ModPlusAPI.Windows;

namespace ModPlus.MinFuncWins
{
    partial class FastBlockSelection
    {
        private readonly List<string> ValidateNames;

        internal FastBlockSelection(List<string> validateNames)
        {
            ValidateNames = validateNames;
            InitializeComponent();
            this.OnWindowStartUp();
        }

        private void BtOk_OnClick(object sender, RoutedEventArgs e)
        {
            if (ValidateNames.Contains(TbBlockName.Text))
            {
                ModPlusAPI.Windows.MessageBox.Show("Такое имя уже используется! Придумайте другое");
                return;
            }
            if (string.IsNullOrEmpty(TbBlockName.Text))
            {
                ModPlusAPI.Windows.MessageBox.Show("Нужно указать отображаемое имя блока!");
                return;
            }
            if (LbBlocks.SelectedIndex == -1)
            {
                ModPlusAPI.Windows.MessageBox.Show("Нужно выбрать блок в списке!");
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
