using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ConsumoVeicoli.Models;

namespace ConsumoVeicoli
{
    public partial class PeriodiWindow : Window
    {
        public List<PeriodoPerVeicolo> Risultato { get; private set; }

        public PeriodiWindow(List<PeriodoPerVeicolo> lista)
        {
            InitializeComponent();
            dgPeriodi.ItemsSource = lista;
            dgPeriodi.CellEditEnding += DgPeriodi_CellEditEnding;
            
            Risultato = lista;
        }


        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in Risultato)
            {
                // Debug
                /*MessageBox.Show(
                    $"Targa {item.Targa}\nDataDa={item.DataDa:dd/MM/yyyy}\nDataA={item.DataA:dd/MM/yyyy}",
                    "Debug date"
                );*/

                if (item.DataDa > item.DataA)
                {
                    MessageBox.Show($"Targa {item.Targa}: Data Da > Data A non valido.");
                    return;
                }
            }
            DialogResult = true;
        }
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
        private void DgPeriodi_CellEditEnding(object? sender, DataGridCellEditEndingEventArgs e)

        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    dgPeriodi.CommitEdit(DataGridEditingUnit.Row, true);
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
        }
    
    }
}
