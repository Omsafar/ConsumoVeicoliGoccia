using System.Configuration;
using System.Data;
using System.Windows;

namespace ConsumoVeicoli;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Forza la cultura italiana (gg/MM/aaaa)
        System.Threading.Thread.CurrentThread.CurrentCulture =
            new System.Globalization.CultureInfo("it-IT");
        System.Threading.Thread.CurrentThread.CurrentUICulture =
            new System.Globalization.CultureInfo("it-IT");

     
    }
  

}

