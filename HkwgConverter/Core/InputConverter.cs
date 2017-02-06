using ClosedXML.Excel;
using HkwgConverter.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HkwgConverter.Core
{
    /// <summary>
    /// Implements the input conversion logic
    /// </summary>
    public class InputConverter
    {
        #region fields

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private AppDataAccessor appDataAccessor;

        #endregion

        #region ctor

        public InputConverter()
        {            
            appDataAccessor = new AppDataAccessor(Settings.Default.AppDataFolder);                                          
        }

        #endregion

        #region private methods

        private List<HkwgInputItem> ReadCsvFile(string fileName)
        {
            var lines = File.ReadAllLines(fileName)
                .Skip(2)
                .Select(x => x.Split(';'))
                .Select(x => new Model.HkwgInputItem()
                {
                    Time = x[0],
                    FPLast = decimal.Parse(x[1]),
                    FlexPos = decimal.Parse(x[2]),
                    FlexNeg = decimal.Parse(x[3]),
                    MarginalCost = decimal.Parse(x[4]),
                });

            return lines.ToList();
        }

        /// <summary>
        /// Makes sure that all 4 values of an hour have the same values
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        private List<HkwgInputItem> Transform(List<Model.HkwgInputItem> content)
        {
            for (int i = 0; i < content.Count; i+=4)
            {
                var quarterhours = content.Skip(i).Take(4);

                var avgflexPos = quarterhours.Average(x => x.FlexPos);
                var avgflexNeg = quarterhours.Average(x => x.FlexNeg);
                var avgCost = quarterhours.Average(x => x.MarginalCost);

                for (int k = i; k <= i+3; k++)
                {
                    content[k].FlexPos = avgflexPos;
                    content[k].FlexNeg = avgflexNeg;
                    content[k].MarginalCost = avgCost;
                }           
            }

            return content;
        }

        private string GenerateKissFile(FileInfo csvFile, List<HkwgInputItem> data, DateTime deliveryDay, int nextVersion, bool isPurchase)
        {
            MemoryStream ms = new MemoryStream(Resource.Template_Kiss_Input);

            XLWorkbook workBook = new XLWorkbook(ms);

            var worksheet = workBook.Worksheets.Skip(1).FirstOrDefault();
          
            WriteHeaderCells(worksheet, deliveryDay.ToString(), isPurchase);

            decimal totalFlexPos = 0.0m;

            for (int i = 0; i < data.Count(); i++)
            {
                worksheet.Cell(18 + i, 1).Value = DateTime.Parse(data[i].Time);
                worksheet.Cell(18 + i, 3).SetValue(isPurchase ? data[i].FlexPos : data[i].FlexNeg);
                worksheet.Cell(18 + i, 4).SetValue(data[i].MarginalCost);
                totalFlexPos+= data[i].FlexPos;
            }

            worksheet.Cell("C15").SetValue(totalFlexPos / 4);
            worksheet.Cell("C114").SetValue(totalFlexPos / 4);



            string outputFileName = String.Format("{0}_{1}_HKWG_{2}.xlsx",
                deliveryDay.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                isPurchase ? "FLEXPOS" : "FLEXNEG",
                nextVersion.ToString("D2"));


            workBook.SaveAs(Path.Combine(csvFile.Directory.FullName, outputFileName), false);

            return outputFileName;
        }
        

        private static void WriteHeaderCells(IXLWorksheet worksheet, string deliveryDay, bool isPurchase)
        {
            worksheet.Cell("C1").Value = deliveryDay;
            worksheet.Cell("D1").Value = deliveryDay;

            if(isPurchase)
            {
                worksheet.Cell("C4").Value = Constants.CottbusSettlementArea;
                worksheet.Cell("D4").Value = Constants.CottbusSettlementArea;

                worksheet.Cell("C5").Value = Constants.EnviamSettlementArea;
                worksheet.Cell("D5").Value = Constants.EnviamSettlementArea;

                worksheet.Cell("C7").Value = Constants.EnviamSettlementArea;
                worksheet.Cell("D7").Value = Constants.EnviamSettlementArea;

                worksheet.Cell("C9").Value = Constants.CottbusBusinessPartnerName;
                worksheet.Cell("D9").Value = Constants.CottbusBusinessPartnerName;

                worksheet.Cell("C10").Value = Constants.CottbusContactPerson;
                worksheet.Cell("D10").Value = Constants.CottbusContactPerson;
            }
            else
            {
                worksheet.Cell("C4").Value = Constants.EnviamSettlementArea;
                worksheet.Cell("D4").Value = Constants.EnviamSettlementArea;

                worksheet.Cell("C5").Value = Constants.CottbusSettlementArea;
                worksheet.Cell("D5").Value = Constants.CottbusSettlementArea;

                worksheet.Cell("C7").Value = Constants.CottbusSettlementArea;
                worksheet.Cell("D7").Value = Constants.CottbusSettlementArea;                
            }            
        }
             

        /// <summary>
        /// Performs the conversion process from HKWG CSV into the KISS-Excel Format
        /// </summary>
        private void ProcessFile(FileInfo csvFile)
        {            
            string flexPosFile = null;
            string flexNegFile = null;

            var content = this.ReadCsvFile(csvFile.FullName);

            content = this.Transform(content);

            var deliveryDay = DateTime.Parse(content.FirstOrDefault().Time).Date;
            var version = appDataAccessor.GetNextInputVersionNumer(deliveryDay);

            // Only generate the file if there any non zero demand values
            if (content.Any(x => x.FlexPos != 0.0m))
            {
                flexPosFile = this.GenerateKissFile(csvFile, content, deliveryDay, version, true);
            }

            // Only generate the file if there any non zero demand values
            if (content.Any(x => x.FlexNeg != 0.0m))
            {
                flexNegFile = this.GenerateKissFile(csvFile, content, deliveryDay, version, false);
            }

            var appDataItem = new AppDataInputItem()
            {
                CsvFile = csvFile.Name,
                DeliveryDay = deliveryDay,
                FlexPosFile = flexPosFile,
                FlexNegFile = flexNegFile,
                Timestamp = DateTime.Now,
                Version = version,
            };

            this.appDataAccessor.AppendInputLogItem(appDataItem);
        }

        #endregion

        #region interface

        /// <summary>
        /// Starts the input conversion process
        /// </summary>
        public void Run()
        {
            log.Info("Suche nach neuen Dateien.");
            var filesToProcess = Directory.GetFiles(Settings.Default.InputWatchFolder, "*.csv");

            if(filesToProcess.Count() > 0)
            {
                log.InfoFormat("Es wurden {0} neue Dateien im Input-Ordner gefunden. Starte Konvertierung...", filesToProcess.Count());
            }
            else
            {
                log.InfoFormat("Es wurden keine neuen Dateien im Input-Ordner gefunden.");
            }            

            foreach (var item in filesToProcess)
            {
                FileInfo file = new FileInfo(item);
                try
                {
                    this.ProcessFile(file);

                    File.Move(file.FullName, Path.Combine(Settings.Default.InputSuccessFolder, file.Name));
                    log.InfoFormat("Datei '{0}' wurde erfolgreich verarbeitet.", file.Name);
                }
                catch (Exception ex)
                {
                    log.Error(ex);
                    var newFilename = file.Name.Replace(".csv", "_" + DateTime.Now.Ticks + ".csv");
                    File.Move(file.FullName, Path.Combine(Settings.Default.InputErrorFolder, newFilename));
                    log.ErrorFormat("Bei der Verarbeitung der Datei '{0}' ist ein Fehler aufgetreten.");
                }
            }
        }

        #endregion
    }
}
