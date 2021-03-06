﻿namespace ModPlus.Windows.MiniPlugins
{
    using System.Collections.Generic;
    using System.Windows;
    using System.Windows.Controls;
    using Autodesk.AutoCAD.DatabaseServices;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

    /// <summary>
    /// Select layer window
    /// </summary>
    public partial class SelectLayer
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SelectLayer"/> class.
        /// </summary>
        public SelectLayer()
        {
            InitializeComponent();
            ContentRendered += (sender, args) => FillLayers();
        }

        private void FillLayers()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            var layers = new List<SelLayer>();
            using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
            {
                var lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                if (lt != null)
                {
                    foreach (var layerId in lt)
                    {
                        var layer = tr.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                        if (layer != null)
                        {
                            layers.Add(new SelLayer(layer.Name, layerId));
                        }
                    }
                }
            }

            LbLayers.ItemsSource = layers;
        }

        private void BtOk_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void LbLayers_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ListBox lb && lb.SelectedIndex != -1)
                BtOk.IsEnabled = true;
        }

        internal class SelLayer
        {
            public SelLayer(string name, ObjectId layerId)
            {
                Name = name;
                LayerId = layerId;
            }

            public string Name { get; }

            public ObjectId LayerId { get; }
        }
    }
}
