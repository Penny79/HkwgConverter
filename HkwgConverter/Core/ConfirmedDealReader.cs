using ClosedXML.Excel;
using HkwgConverter.Model;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;

namespace HkwgConverter.Core
{
    public class ConfirmedDealReader
    {
        #region fields


        private static ILogger log = LogManager.GetCurrentClassLogger();
        private Settings configData;
        private BusinessConfigurationSection businessSettings;

        #endregion

        #region ctor

        public ConfirmedDealReader(BusinessConfigurationSection businessSettings)
        {
            this.businessSettings = businessSettings;
        }


        #endregion

        #region private methods

        private int DetermineColumnFromSettlementAreas(IXLWorksheet workssheet, bool isFlexPos)
        {
            string fromSettlementArea = isFlexPos ? this.businessSettings.PartnerCottbus.SettlementArea : this.businessSettings.PartnerEnviaM.SettlementArea;
            string toSettlementArea = isFlexPos ? this.businessSettings.PartnerEnviaM.SettlementArea : this.businessSettings.PartnerCottbus.SettlementArea;

            for (char k = 'C'; k <= 'D'; k++)
            {
                if (fromSettlementArea == workssheet.Cell(k + "4").Value.ToString() &&
                toSettlementArea == workssheet.Cell(k + "5").Value.ToString())
                    return (k - 'A') + 1;
            }

            return -1;
        }

        #endregion

        #region interface

        public Dictionary<string,CsvLineItem> Read(string fileName)
        {
            XLWorkbook workBook = new XLWorkbook(fileName);
            var worksheet = workBook.Worksheets.Skip(1).FirstOrDefault();

            if (worksheet.Cell("D17").Value.ToString() != "MW")
            {
                string msg = "Die abgelegte Exceldatei hat nicht in Zelle D17 den Wert 'MW' stehen.";
                log.Error(msg);
                throw new InvalidOperationException(msg);
            }

            var lastrow = worksheet.Rows()
                            .FirstOrDefault(x => x.FirstCell().Value.ToString().Contains("23:45"));

            var lines = new Dictionary<string,CsvLineItem>();

            int flexPosCol = DetermineColumnFromSettlementAreas(worksheet, true);
            int flexNegCol = DetermineColumnFromSettlementAreas(worksheet, false);

            if (flexPosCol == -1 || flexNegCol == -1)
            {
                string msg = "Die Spalte mit den Leistungswerten für FlexPos oder FlexNeg konnte nicht identifiziert werden.";
                log.Error(msg);
                throw new InvalidOperationException(msg);
            }

            for (int i = 18; i <= lastrow.RowNumber(); i++)
            {
                var currentRow = worksheet.Row(i);

                var entry = new CsvLineItem()
                {
                    Time = currentRow.Cell(1).GetString(),
                    FlexNegDemand = currentRow.Cell(flexNegCol).GetValue<decimal>(),
                    FlexPosDemand = currentRow.Cell(flexPosCol).GetValue<decimal>()
                };
                
                lines.Add(DateTime.Parse(entry.Time).ToString("dd.MM.yyyy HH:mm"), entry);
            }

            return lines;
        }

        #endregion
    }
}
