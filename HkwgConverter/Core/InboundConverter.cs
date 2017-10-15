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

        private static ILogger log = LogManager.GetCurrentClassLogger();

        private TransactionRepository transactionRepository;
        private Settings configData;
        private BusinessConfigurationSection businessSettings;

        #endregion

        #region ctor

        public InboundConverter(TransactionRepository store, Settings config, BusinessConfigurationSection businessSettings)
        {
            this.transactionRepository = store;
            this.configData = config;
            this.businessSettings = businessSettings;
        }

        #endregion

        #region private methods

        private List<CsvLineItem> ReadCsvFile(string fileName)
        {
            var lines = File.ReadAllLines(fileName)
                .Skip(2)
                .Select(x => x.Split(';'))
                .Select(x => new CsvLineItem()
                {
                    Time = x[0],
                    FlexPosDemand = decimal.Parse(x[1], System.Globalization.CultureInfo.InvariantCulture),
                    FlexPosMarginalCost = decimal.Parse(x[2], System.Globalization.CultureInfo.InvariantCulture),
                    FlexNegDemand = decimal.Parse(x[3], System.Globalization.CultureInfo.InvariantCulture),
                    FlexNegMarginalCost = decimal.Parse(x[4], System.Globalization.CultureInfo.InvariantCulture)
                });

            return lines.ToList();
        }

        /// <summary>
        /// Makes sure that all 4 values of an hour have the same values
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        private List<CsvLineItem> Transform(List<CsvLineItem> content)
        {
            for (int i = 0; i < content.Count; i+=4)
            {
                
                    var quarterhours = content.Skip(i).Take(4);

                    var avgflexPosTotal = quarterhours.Average(x => x.FlexPosDemand);
                    var avgflexPosCosts = quarterhours.Average(x => x.FlexPosMarginalCost);
                    var avgflexNegTotal = quarterhours.Average(x => x.FlexNegDemand);
                    var avgflexNegCosts = quarterhours.Average(x => x.FlexNegMarginalCost);

                    for (int k = i; k <= i + 3; k++)
                    {
                        content[k].FlexPosDemand = avgflexPosTotal;
                        content[k].FlexPosMarginalCost = avgflexPosCosts;
                        content[k].FlexNegDemand = avgflexNegTotal;
                        content[k].FlexNegMarginalCost = avgflexNegCosts;
                    }           
            }

            return content;
        }

        private string GenerateKissFile(List<CsvLineItem> data, DateTime deliveryDay, int nextVersion, bool isPurchase)
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
                worksheet.Cell(currentRow, 3).SetValue(isPurchase ? data[i].FlexPosDemand : data[i].FlexNegDemand);
                worksheet.Cell(currentRow, 4).SetValue(isPurchase ? data[i].FlexPosMarginalCost : data[i].FlexNegMarginalCost);

                if ((i + 1) % 4 == 0)
                {
                    worksheet.Range("A" + currentRow.ToString(), "D" + currentRow.ToString()).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                }

                totalEnergy += isPurchase ? data[i].FlexPosDemand : data[i].FlexNegDemand;
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

        private void WriteHeaderCells(IXLWorksheet worksheet, string deliveryDay, bool isPurchase)
        {
            var fromBusinessPartner = isPurchase ? this.businessSettings.PartnerCottbus : this.businessSettings.PartnerEnviaM;
            var toBusinessPartner = isPurchase ? this.businessSettings.PartnerEnviaM : this.businessSettings.PartnerCottbus;

            worksheet.Cell("C1").Value = deliveryDay;
            worksheet.Cell("D1").Value = deliveryDay;

            worksheet.Cell("C4").Value = fromBusinessPartner.SettlementArea;
            worksheet.Cell("D4").Value = fromBusinessPartner.SettlementArea;

            worksheet.Cell("C5").Value = toBusinessPartner.SettlementArea;
            worksheet.Cell("D5").Value = toBusinessPartner.SettlementArea;

            worksheet.Cell("C7").Value = fromBusinessPartner.SettlementArea;
            worksheet.Cell("D7").Value = fromBusinessPartner.SettlementArea;

            worksheet.Cell("C9").Value = fromBusinessPartner.BusinessPartnerName;
            worksheet.Cell("D9").Value = fromBusinessPartner.BusinessPartnerName;

            worksheet.Cell("C10").Value = fromBusinessPartner.ContactPerson;
            worksheet.Cell("D10").Value = fromBusinessPartner.ContactPerson;

            worksheet.Cell("C11").Value = toBusinessPartner.BusinessPartnerName;
            worksheet.Cell("D11").Value = toBusinessPartner.BusinessPartnerName;

            worksheet.Cell("C12").Value = toBusinessPartner.ContactPerson;
            worksheet.Cell("D12").Value = toBusinessPartner.ContactPerson;
        }
             

        /// <summary>
        /// Performs the conversion process from HKWG CSV into the KISS-Excel Format
        /// </summary>
        private void ProcessFile(FileInfo csvFile)
        {            
            string flexPosFile = null;
            string flexNegFile = null;

            var flexibilities = this.ReadCsvFile(csvFile.FullName);

            if (businessSettings.FlattenDemandPerHour)
            {
                flexibilities = this.Transform(flexibilities);
            }

            var deliveryDay = DateTime.Parse(flexibilities.FirstOrDefault().Time).Date;

            var lastTransactionforDeliveryDay = transactionRepository.GetLatest(deliveryDay);

            if(lastTransactionforDeliveryDay != null)
            {
                // subtract the already booked numbers from the new ones
                flexibilities = this.SubtractAmountsofPreviousTransaction(flexibilities, lastTransactionforDeliveryDay);
            }

            var version = transactionRepository.GetNextInputVersionNumer(deliveryDay);

            // Only generate the file if there any non zero demand values
            if (flexibilities.Any(x => x.FlexPosDemand != 0.0m))
            {
                flexPosFile = this.GenerateKissFile(flexibilities, deliveryDay, version, true);
            }

            // Only generate the file if there any non zero demand values
            if (flexibilities.Any(x => x.FlexNegDemand != 0.0m))
            {
                flexNegFile = this.GenerateKissFile(flexibilities, deliveryDay, version, false);
            }

            var transaction = new Transaction()
            {
                CsvFile = csvFile.Name,
                DeliveryDate = deliveryDay,
                FlexPosFile = flexPosFile,
                FlexNegFile = flexNegFile,
                CreateDate = DateTime.Now,
                UpdateDate = DateTime.Now,
                Version = version,
            };

            //var newRow = this.transactionRepository.Set<Transaction>();
            //newRow.Add(transaction);

            this.transactionRepository.Transactions.Add(transaction);
            this.transactionRepository.SaveChanges();

        }

        private List<CsvLineItem> SubtractAmountsofPreviousTransaction(List<CsvLineItem> flexibilities, Transaction lastTransactionforDeliveryDay)
        {
            var kissReader = new ConfirmedDealReader(this.businessSettings);

            var excelFilePath = Path.Combine(this.configData.OutboundSuccessFolder, lastTransactionforDeliveryDay.ConfirmedDealFile);
            var timeSliceData = kissReader.Read(excelFilePath);

            foreach (var item in flexibilities)
            {
                if(timeSliceData.ContainsKey(item.Time))
                {
                    var previousItem = timeSliceData[item.Time];
                    item.FlexNegDemand -= previousItem.FlexNegDemand;
                    item.FlexPosDemand -= previousItem.FlexPosDemand;
                }
            }

            return flexibilities;
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
