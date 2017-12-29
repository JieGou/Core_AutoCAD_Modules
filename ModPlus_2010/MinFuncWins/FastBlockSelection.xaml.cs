using System.Collections.Generic;
using System.Windows;
using ModPlusAPI.Windows.Helpers;

namespace ModPlus.MinFuncWins
{
    internal partial class FastBlockSelection
    {
        private readonly List<string> _validateNames;
        private static string _langItem = "AutocadDlls";

        internal FastBlockSelection(List<string> validateNames)
        {
            _validateNames = validateNames;
            InitializeComponent();
            this.OnWindowStartUp();
        }

        private void BtOk_OnClick(object sender, RoutedEventArgs e)
        {
            if (_validateNames.Contains(TbBlockName.Text))
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(_langItem, "err1"));
                return;
            }
            if (string.IsNullOrEmpty(TbBlockName.Text))
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(_langItem, "err2"));
                return;
            }
            if (LbBlocks.SelectedIndex == -1)
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(_langItem, "err3"));
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
