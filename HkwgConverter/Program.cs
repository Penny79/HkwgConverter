using HkwgConverter.Core;
using log4net.Config;
using System;
using System.IO;
using System.Linq;

namespace HkwgConverter
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
                

        private static bool ValidateConfiguration()
        {
            log.Info("Validiere Konfigurationseinstellungen.");
            bool isValid = true;

            var configuredFolders = Settings.Default.GetType().GetProperties().Where(p => p.Name.EndsWith("Folder"));

            foreach (var item in configuredFolders)
            {
                var value = item.GetValue(Settings.Default).ToString();

                if (!Directory.Exists(value))
                {
                    log.ErrorFormat("Der konfigurierte Pfad '{1}' im Setting '{0}' kann nicht gefunden werden.", item.Name, value);
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
            XmlConfigurator.Configure();

            log.Info("---------------------------------------------------------------------------------------------------");
            log.Info("Starte HKWG Converter");
            log.Info("---------------------------------------------------------------------------------------------------");

            bool isValid = ValidateConfiguration();

            if (isValid)
            {
                var appDataAccessor = new WorkflowStore(Settings.Default.WorkflowStoreFolder);

                var inboundConversion = new InboundConverter(appDataAccessor, Settings.Default);
                inboundConversion.Run();

                var outboundConversion = new OutboundConverter(appDataAccessor, Settings.Default);
                outboundConversion.Run();
            }

            log.Info("---------------------------------------------------------------------------------------------------");
            log.Info("Anwendung wird beendet");
            log.Info("---------------------------------------------------------------------------------------------------");
        }
    }
}
