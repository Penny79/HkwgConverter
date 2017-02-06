using HkwgConverter.Core;
using log4net.Config;
using System;
using System.IO;

namespace HkwgConverter
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static void InitInputWatcher()
        {
            var inputWatcher = new FileSystemWatcher()
            {
                Path = Settings.Default.InputWatchFolder,
                NotifyFilter = NotifyFilters.LastWrite |NotifyFilters.LastAccess | NotifyFilters.LastWrite| NotifyFilters.FileName | NotifyFilters.DirectoryName,
                Filter = "*.csv"
            };
                      
            inputWatcher.Created += new FileSystemEventHandler(OnChanged);

            // Begin watching.
            inputWatcher.EnableRaisingEvents = true;

            log.Info("Überwachung des Input Verzeichnis gestartet.");
        }
        

        private static void OnChanged(object source, FileSystemEventArgs e)
        {
            try
            {
                log.InfoFormat("Neue CSV -Datei im Inputordner gefunden. '{0}'", e.Name);
                var inputConverter = new InputConverter(e.FullPath);
                inputConverter.Execute();
                log.ErrorFormat("Datei '{0}' wurde erfolgreich verarbeitet.");
            }
            catch (Exception ex)
            {
                log.Error(ex);
                File.Move(e.FullPath, Path.Combine(Settings.Default.InputErrorFolder, e.Name));
                log.ErrorFormat("Bei der Verarbeitung der Datei '{0}' ist ein Fehler aufgetreten.");
            }
            
        }

        private static bool ValidateConfiguration()
        {
            log.Info(".......Validiere Konfigurationseinstellungen.");
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

            log.Info("......................Starte HKWG Converter.................");

            bool isValid = ValidateConfiguration();

            if (isValid)
            {
                InitInputWatcher();

                Console.WriteLine("Press \'q\' to quit the sample.");
                while (Console.Read() != 'q') ;
            }

            log.Info("......................Anwendung wird beendet................");
        }
    }
}
