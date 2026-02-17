using Microsoft.Extensions.Configuration;
using System.Configuration;
using System.Data;
using System.Windows;

namespace CreadorDeFormatos
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var config = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            string lang = config["Language"] ?? "en";
            Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo(lang);

        }
    }

}
