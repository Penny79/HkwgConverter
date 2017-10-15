using ClosedXML.Excel;
using HkwgConverter.Model;
using NLog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace HkwgConverter.Core
{
    /// <summary>
    /// Implements the export conversion logic
    /// </summary>
    public class OutboundConverter
    {
        #region fields


        private static ILogger log = LogManager.GetCurrentClassLogger();
        private TransactionRepository transactionRepository;
        private Settings configData;
        private const string outputCsvFilePrefix = "Cottbus_ConfirmedDeal_";
        private BusinessConfigurationSection businessSettings;

        #endregion

        #region ctor

        public OutboundConverter(TransactionRepository store, Settings config, BusinessConfigurationSection businessSettings)
        {
            this.transactionRepository = store;
            this.configData = config;
            this.businessSettings = businessSettings;
        }


        #endregion

        #region private methods

        /// <summary>
        /// Performs the conversion process from HKWG CSV into the KISS-Excel Format
        /// </summary>
        private void ProcessFile(FileInfo excelFile)
        {
            var kissReader = new ConfirmedDealReader(this.businessSettings);
            var timeSliceData = kissReader.Read(excelFile.FullName);

            var deliveryDay = DateTime.Parse(timeSliceData.Values.FirstOrDefault().Time).Date;

            var currentTransaction = this.transactionRepository.GetLatest(deliveryDay);

            if (currentTransaction == null)
            {
                log.Error("Die Datei {0} kann nicht verarbeitet werden weil es für den Liefertag keinen offenen Prozess gibt.");
                return;
            }

            currentTransaction.ConfirmedDealFile = excelFile.Name;
            currentTransaction.UpdateDate = DateTime.Now;
            this.transactionRepository.SaveChanges();

            this.WriteCsvFile(currentTransaction, timeSliceData.Values.ToList());
        }       

        private void WriteCsvFile(Transaction transacion, List<CsvLineItem> newData)
        {
            
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Zeitstempel;FlexPos_Change;FlexNeg_Change");
            sb.AppendLine("[dd.MM.yyyy hh:mm:ss];[mw.kw];[mw.kw]");

            for (int i = 0; i < newData.Count; i++)
            {
               
                string time = DateTime.Parse(newData[i].Time).ToString("dd.MM.yyyy HH:mm:ss");
                
                sb.Append(time);
                sb.Append(";");
                sb.Append(newData[i].FlexPosDemand.ToString("0.00", CultureInfo.InvariantCulture));
                sb.Append(";");
                sb.AppendLine(newData[i].FlexNegDemand.ToString("0.00", CultureInfo.InvariantCulture));
            }
            
            var targetFile = Path.Combine(this.configData.OutboundDropFolder, outputCsvFilePrefix + transacion.CsvFile.Replace("Cottbus_FlexDayAhead_", ""));

            File.WriteAllText(targetFile, sb.ToString());    
                     
        }

        #endregion

        #region interface

        /// <summary>
        /// Starts the input conversion process
        /// </summary>
        public void Run()
        {
            log.Info("Suche nach neuen Dateien.");
            var filesToProcess = Directory.GetFiles(Settings.Default.OutboundWatchFolder, "*.xlsx");

            if(filesToProcess.Count() > 0)
            {
                log.Info("Es wurden {0} neue Dateien im Outbound-Ordner gefunden. Starte Konvertierung...", filesToProcess.Count());
            }
            else
            {
                log.Info("Es wurden keine neuen Dateien im Outbound-Ordner gefunden.");
            }            

            foreach (var item in filesToProcess)
            {
                FileInfo file = new FileInfo(item);
                string newFileName = string.Empty;
                try
                {
                    this.ProcessFile(file);
                    newFileName = Path.Combine(Settings.Default.OutboundSuccessFolder, file.Name);
                    File.Move(file.FullName, newFileName);
                    log.Info("Datei '{0}' wurde erfolgreich verarbeitet.", file.Name);
                }
                catch (Exception ex)
                {
                    log.Error(ex.ToString());
                    newFileName = file.Name.Replace(".csv", "_" + DateTime.Now.Ticks + ".csv");
                    newFileName = Path.Combine(Settings.Default.OutboundErrorFolder, newFileName);

                    File.Move(file.FullName, newFileName);

                    log.Error("Bei der Verarbeitung der Datei '{0}' ist ein Fehler aufgetreten.", file.Name);
                }
                finally
                {
                    File.SetLastWriteTime(newFileName, DateTime.Now);
                }
            }
        }

        #endregion
    }
}
