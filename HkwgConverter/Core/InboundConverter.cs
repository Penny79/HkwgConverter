using ClosedXML.Excel;
using HkwgConverter.Model;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HkwgConverter.Core
{
    /// <summary>
    /// Implements the input conversion logic
    /// </summary>
    public class InboundConverter
    {
        #region fields

        private static LogWrapper log = LogWrapper.GetLogger(LogManager.GetCurrentClassLogger());

        private WorkflowStore workflowStore;
        private Settings configData;

        #endregion

        #region ctor

        public InboundConverter(WorkflowStore store, Settings config)
        {
            this.workflowStore = store;
            this.configData = config;
        }

        #endregion

        #region private methods

        private List<HkwgInputItem> ReadCsvFile(string fileName)
        {
            var lines = File.ReadAllLines(fileName)
                .Skip(2)
                .Select(x => x.Split(';'))
                .Select(x => new HkwgInputItem()
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
        private List<HkwgInputItem> Transform(List<HkwgInputItem> content)
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

            decimal totalEnergy = 0.0m;           

            int rowOffset = 18;
            for (int i = 0; i < data.Count(); i++)
            {
                var toTime = DateTime.Parse(data[i].Time).AddMinutes(15).ToString("HH:mm");

                int currentRow = rowOffset + i;
                worksheet.Cell(currentRow, 1).SetValue(data[i].Time);
                worksheet.Cell(currentRow, 2).SetValue(toTime);
                worksheet.Cell(currentRow, 3).SetValue(isPurchase ? data[i].FlexPos : data[i].FlexNeg);
                worksheet.Cell(currentRow, 4).SetValue(data[i].MarginalCost);

                if ((i + 1) % 4 == 0)
                {
                    worksheet.Range("A" + currentRow.ToString(), "D" + currentRow.ToString()).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                }

                totalEnergy += isPurchase ? data[i].FlexPos : data[i].FlexNeg;
            }

            //Some styling
            int lastRow = rowOffset + data.Count() - 1;

            //write footer values
            worksheet.Cell("C15").SetValue(totalEnergy / 4);
            worksheet.Cell(lastRow + 1, 2).SetValue(" Arbeit[MWh]:");
            worksheet.Cell(lastRow + 1, 3).SetValue(totalEnergy / 4);

            ApplyStyling(worksheet, lastRow);

            string outputFileName = String.Format("{0}_{1}_HKWG_{2}.xlsx",
                deliveryDay.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                isPurchase ? "FLEXPOS" : "FLEXNEG",
                nextVersion.ToString("D2"));


            workBook.SaveAs(Path.Combine(configData.InboundDropFolder, outputFileName), false);

            return outputFileName;
        }

        private static void ApplyStyling(IXLWorksheet worksheet, int lastRow)
        {
            var color1 = XLColor.FromArgb(192, 192, 192);
            var color2 = XLColor.FromArgb(204, 255, 204);
            var color3 = XLColor.FromArgb(255, 255, 153);

            worksheet.Range("A18", "B" + lastRow.ToString()).Style.Fill.SetBackgroundColor(color1);

            worksheet.Range("C18", "C" + lastRow.ToString()).Style.Fill.SetBackgroundColor(color2);
            worksheet.Range("C18", "C" + (lastRow + 1).ToString()).Style.NumberFormat.SetFormat("0.000");

            worksheet.Range("D18", "D" + lastRow.ToString()).Style.Fill.SetBackgroundColor(color3);
            worksheet.Range("D18", "D" + lastRow.ToString()).Style.NumberFormat.SetFormat("0.00");

            worksheet.Range("A18", "D" + (lastRow + 1).ToString()).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            worksheet.Range("A18", "D" + lastRow.ToString()).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            worksheet.Range("A18", "B" + lastRow.ToString()).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            worksheet.Range("C18", "C" + lastRow.ToString()).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            worksheet.Range("A18", "A" + lastRow.ToString()).Style.Border.RightBorder = XLBorderStyleValues.Thin;

            worksheet.Range("A" + (lastRow + 1).ToString(), "D" + (lastRow + 1).ToString()).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            worksheet.Range("A" + (lastRow + 1).ToString(), "D" + (lastRow + 1).ToString()).Style.Font.Bold = true;
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
            var version = workflowStore.GetNextInputVersionNumer(deliveryDay);

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

            var appDataItem = new Workflow()
            {
                CsvFile = csvFile.Name,
                DeliveryDay = deliveryDay,
                FlexPosFile = flexPosFile,
                FlexNegFile = flexNegFile,
                Timestamp = DateTime.Now,
                Version = version,
            };

            this.workflowStore.Add(appDataItem);
        }

        #endregion

        #region interface

        /// <summary>
        /// Starts the input conversion process
        /// </summary>
        public void Run()
        {
           
            log.Info("Suche nach neuen Dateien.");
            var filesToProcess = Directory.GetFiles(this.configData.InboundWatchFolder, "*.csv");

            if(filesToProcess.Count() > 0)
            {
                log.Info("Es wurden {0} neue Dateien im Input-Ordner gefunden. Starte Konvertierung...", filesToProcess.Count());
            }
            else
            {
                log.Info("Es wurden keine neuen Dateien im Input-Ordner gefunden.");
            }            

            foreach (var item in filesToProcess)
            {
                FileInfo file = new FileInfo(item);
                string newFileName = string.Empty;
                try
                {
                    this.ProcessFile(file);
                    newFileName = Path.Combine(this.configData.InboundSuccessFolder, file.Name);
                    File.Move(file.FullName, newFileName);
                    log.Info("Datei '{0}' wurde erfolgreich verarbeitet.", file.Name);
                }
                catch (Exception ex)
                {
                    log.Error(ex.ToString());
                    newFileName = file.Name.Replace(".csv", "_" + DateTime.Now.Ticks + ".csv");
                    newFileName = Path.Combine(this.configData.InboundErrorFolder, newFileName);

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
