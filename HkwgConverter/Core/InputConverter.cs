using ClosedXML.Excel;
using HkwgConverter.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HkwgConverter.Core
{
    public class InputConverter
    {
        private FileInfo csvFile;      

        public InputConverter(string fileName)
        {
            if (!File.Exists(fileName))
                throw new ArgumentException(String.Format("The given file '{0}' does not exist", fileName));

            csvFile = new FileInfo(fileName);                                   
        }

        #region private methods

        private List<Model.HkwgInputItem> ParseInput()
        {
            var lines = File.ReadAllLines(this.csvFile.FullName)
                .Skip(2)
                .Select(x => x.Split(';'))
                .Select(x => new Model.HkwgInputItem()
                       {
                           Time = x[0],
                           FPLast = Decimal.Parse(x[1]),
                           FlexPos = Decimal.Parse(x[2]),
                           FlexNeg = Decimal.Parse(x[3]),
                           MarginalCost = Decimal.Parse(x[4]),
                       });

            return lines.ToList();
        }

        /// <summary>
        /// Makes sure that all 4 values of an hour have the same values
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        private List<Model.HkwgInputItem> Transform(List<Model.HkwgInputItem> content)
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

        private void GenerateKissFile(List<HkwgInputItem> data, bool isPurchase)
        {
            MemoryStream ms = new MemoryStream(Resource.Template_Kiss_Input);

            XLWorkbook workBook = new XLWorkbook(ms);
            var worksheet = workBook.Worksheets.Skip(1).FirstOrDefault();

            var deliveryDay = DateTime.Parse(data.FirstOrDefault().Time).Date;
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
                "01");


            workBook.SaveAs(Path.Combine(csvFile.Directory.FullName, outputFileName), false);
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

        #endregion

        #region interface

        /// <summary>
        /// Performs the conversion process from HKWG CSV into the KISS-Excel Format
        /// </summary>
        public void Execute()
        {
            var content = this.ParseInput();

            content = this.Transform(content);

            this.GenerateKissFile(content, true);

            this.GenerateKissFile(content, false);
        }

        #endregion
    }
}
