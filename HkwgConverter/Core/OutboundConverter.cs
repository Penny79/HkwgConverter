using ClosedXML.Excel;
using HkwgConverter.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HkwgConverter.Core
{
    /// <summary>
    /// Implements the export conversion logic
    /// </summary>
    public class OutboundConverter
    {
        #region fields

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private AppDataAccessor appDataAccessor;

        #endregion

        #region ctor

        public OutboundConverter()
        {            
            appDataAccessor = new AppDataAccessor(Settings.Default.AppDataFolder);                                          
        }

        #endregion

        #region private methods


        private List<HkwgInputItem> ReadFile(string fileName)
        {
            XLWorkbook workBook = new XLWorkbook(fileName);
            var worksheet = workBook.Worksheets.Skip(1).FirstOrDefault();
            
            var lastrow = worksheet.Rows()
                            .FirstOrDefault(x => x.FirstCell().Value.ToString().Contains("23:45"));

            var lines = new List<HkwgInputItem>();

            for (int i = 18; i <= lastrow.RowNumber(); i++)
            {
                var currentRow = worksheet.Row(i);

                var entry = new HkwgInputItem()
                {
                    Time = currentRow.Cell(1).GetString(),
                    FlexNeg = currentRow.Cell(3).GetValue<decimal>(),
                    FlexPos = currentRow.Cell(4).GetValue<decimal>()
                };

                lines.Add(entry);
            }
            
            return lines.ToList();
        }



        /// <summary>
        /// Performs the conversion process from HKWG CSV into the KISS-Excel Format
        /// </summary>
        private void ProcessFile(FileInfo csvFile)
        {                      
            var content = this.ReadFile(csvFile.FullName);

                      
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
                log.InfoFormat("Es wurden {0} neue Dateien im Outbound-Ordner gefunden. Starte Konvertierung...", filesToProcess.Count());
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
                    log.InfoFormat("Datei '{0}' wurde erfolgreich verarbeitet.", file.Name);
                }
                catch (Exception ex)
                {
                    log.Error(ex);
                    newFileName = file.Name.Replace(".csv", "_" + DateTime.Now.Ticks + ".csv");
                    newFileName = Path.Combine(Settings.Default.OutboundErrorFolder, newFileName);

                    File.Move(file.FullName, newFileName);

                    log.ErrorFormat("Bei der Verarbeitung der Datei '{0}' ist ein Fehler aufgetreten.", file.Name);
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
