using HkwgConverter.Core;
using log4net.Config;
using System;
using System.IO;

namespace HkwgConverter
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
                

        private static bool ValidateConfiguration()
        {
            log.Info("Validiere Konfigurationseinstellungen.");
            bool isValid = true;

            if(!Directory.Exists(Settings.Default.InputWatchFolder))
            {
                log.ErrorFormat("Setting: 'InputWatchFolder' Wert='{0}'. Dies ist das Verzeichnis in dem die CSV-Dateien von Cottbus ankommen. ", Settings.Default.InputWatchFolder);
                isValid = false;
            }

            if (!Directory.Exists(Settings.Default.InputSuccessFolder))
            {
                log.ErrorFormat("Setting: 'InputSuccessFolder' - Wert='{0}'. Hierhin werden erfolgreich verarbeitete CSV-Dateien verschoben.", Settings.Default.InputSuccessFolder);
                isValid = false;
            }

            if (!Directory.Exists(Settings.Default.InputErrorFolder))
            {
                log.ErrorFormat("Setting: 'InputErrorFolder' - Wert='{0}'. Hierhin werden nicht verarbeitbare CSV-Dateien verschoben.", Settings.Default.InputErrorFolder);
                isValid = false;
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
                var inputConverter = new InputConverter();
                inputConverter.Run();               
            }

            log.Info("---------------------------------------------------------------------------------------------------");
            log.Info("Anwendung wird beendet");
            log.Info("---------------------------------------------------------------------------------------------------");
        }
    }
}
