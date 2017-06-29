namespace HkwgConverter.Model
{
    /// <summary>
    /// abstraction for the items in the Cottbus CSV file
    /// </summary>
    public class CsvLineItem
    {
        public string Time { get; set; }

        public decimal FlexPosDemand { get; set; }

        public decimal FlexPosMarginalCost { get; set; }

        public decimal FlexNegDemand { get; set; }

        public decimal FlexNegMarginalCost { get; set; }

    }
}
