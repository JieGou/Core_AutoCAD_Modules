using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Size = System.Drawing.Size;
using UserControl = System.Windows.Controls.UserControl;

namespace ModPlus.MPCOHelpers.DynProperties
{
    /// <summary>
    /// Логика взаимодействия для PropertiesPalette.xaml
    /// </summary>
    public partial class PropertiesPalette : UserControl
    {
        public PropertiesPalette()
        {
            InitializeComponent();
            AcadHelpers.Document.ImpliedSelectionChanged += Document_ImpliedSelectionChanged;
        }

        private void Document_ImpliedSelectionChanged(object sender, EventArgs e)
        {
            PromptSelectionResult psr = AcadHelpers.Editor.SelectImplied();
            if(psr.Value == null) return;
            var selected = psr.Value[0];
            if (selected is BlockReference)
            {
            }
        }
    }

    public class PropertiesFunction
    {
        private PaletteSet _paletteSet;
        [CommandMethod("ModPlus","mpPropertiesPalette", CommandFlags.Modal)]
        public void Start()
        {
            _paletteSet = new PaletteSet("MP: Свойства примитивов ModPlus", new Guid("756d08d2-fa22-4496-aae6-73d5e98ab722"));
            _paletteSet.Load += _paletteSet_Load;
            _paletteSet.Save += _paletteSet_Save;
            PropertiesPalette propertiesPalette = new PropertiesPalette();
            ElementHost elementHost = new ElementHost()
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Child = propertiesPalette
            };
            _paletteSet.Add("Свойства примитивов ModPlus", elementHost);
            _paletteSet.Style = PaletteSetStyles.ShowCloseButton | PaletteSetStyles.ShowPropertiesMenu | PaletteSetStyles.ShowAutoHideButton;
            _paletteSet.MinimumSize = new Size(100, 300);
            _paletteSet.DockEnabled = DockSides.Right | DockSides.Left;
            _paletteSet.Visible = true;
        }
        private void _paletteSet_Load(object sender, PalettePersistEventArgs e)
        {
            double num = (double)e.ConfigurationSection.ReadProperty("mpPropertiesPalette", 22.3);
        }

        private void _paletteSet_Save(object sender, PalettePersistEventArgs e)
        {
            e.ConfigurationSection.WriteProperty("mpPropertiesPalette", 32.3);
        }
    }
}
