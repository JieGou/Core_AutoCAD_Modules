namespace ModPlus.MinFuncWins
{
    using System.Collections.Generic;
    using System.Windows;

    internal partial class FastBlockSelection
    {
        private readonly List<string> _validateNames;
        private const string LangItem = "AutocadDlls";

        internal FastBlockSelection(List<string> validateNames)
        {
            _validateNames = validateNames;
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem(LangItem, "h27");
        }

        private void BtOk_OnClick(object sender, RoutedEventArgs e)
        {
            if (_validateNames.Contains(TbBlockName.Text))
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err1"));
                return;
            }

            if (string.IsNullOrEmpty(TbBlockName.Text))
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err2"));
                return;
            }

            if (LbBlocks.SelectedIndex == -1)
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err3"));
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
