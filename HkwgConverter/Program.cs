using HkwgConverter.Core;
using HkwgConverter.Model;
using NLog;
using System;
using System.IO;
using System.Linq;

namespace HkwgConverter
{
    class Program
    {
        private static LogWrapper log = LogWrapper.GetLogger(LogManager.GetCurrentClassLogger());
       
        private static bool ValidateConfiguration()
        {
            log.Info("Validiere Konfigurationseinstellungen.");
            bool isValid = true;

            var configuredFolders = Settings.Default.GetType().GetProperties().Where(p => p.Name.EndsWith("Folder", StringComparison.InvariantCultureIgnoreCase));

            foreach (var item in configuredFolders)
            {
                var value = item.GetValue(Settings.Default, null).ToString();

                if (!Directory.Exists(value))
                {
                   log.Error("Der konfigurierte Pfad '{1}' im Setting '{0}' kann nicht gefunden werden.", item.Name, value);
                    isValid = false;
                }
            }

            if(!isValid)
            {
                log.Error("Es wurden Konfigurationsfehler gefunden. Bitte überprüfen Sie die Ihre Einstellungen in der Datei 'HkwgConverter.config'.");
            }

            return isValid;
        }

        static void Main(string[] args)
        {
            log.Info("---------------------------------------------------------------------------------------------------");
            log.Info("Starte HKWG Converter");
            log.Info("---------------------------------------------------------------------------------------------------");

            bool isValid = ValidateConfiguration();

            if (isValid)
            {
                var appDataAccessor = new WorkflowStore(Settings.Default.WorkflowStoreFolder);

                var test = System.Configuration.ConfigurationManager.GetSection("businessSettings");

                var config = (BusinessConfigurationSection)System.Configuration.ConfigurationManager.GetSection("businessSettings");

                var inboundConversion = new InboundConverter(appDataAccessor, Settings.Default, config);
                inboundConversion.Run();

                var outboundConversion = new OutboundConverter(appDataAccessor, Settings.Default);
                outboundConversion.Run();
            }

            log.Info("---------------------------------------------------------------------------------------------------");
            log.Info("Anwendung wird beendet");
            log.Info("---------------------------------------------------------------------------------------------------");

#if DEBUG
            Console.ReadLine();

#endif
        }
    }
}
