namespace HkwgConverter.Model
{
    /// <summary>
    /// abstraction for the items in the Cottbus CSV file
    /// </summary>
    public class HkwgInputItem
    {
        public string Time { get; set; }

        public decimal FPLast { get; set; }

        public decimal FlexPos { get; set; }

        public decimal FlexNeg { get; set; }

        public decimal MarginalCost { get; set; }

    }
}
