using System;
using System.Windows;

namespace ConsumoVeicoli
{
    public partial class FinestraDate : Window
    {
        public DateTime DataDa { get; private set; }
        public DateTime DataA { get; private set; }

        public FinestraDate()
        {
            InitializeComponent();
            dpDa.SelectedDate = DateTime.Today;
            dpA.SelectedDate = DateTime.Today;
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            DataDa = dpDa.SelectedDate ?? DateTime.Today;
            DataA = dpA.SelectedDate ?? DateTime.Today;
            DialogResult = true;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
