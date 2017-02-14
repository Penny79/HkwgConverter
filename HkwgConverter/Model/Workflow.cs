using System;
using System.Linq;

namespace HkwgConverter.Model
{
    /// <summary>
    /// Abstraction for the items in the app data input log
    /// </summary>
    public class Workflow
    {
        public DateTime Timestamp { get; set; }

        public DateTime DeliveryDay { get; set; }

        public string CsvFile { get; set; }

        public string FlexPosFile { get; set; }

        public string FlexNegFile { get; set; }

        public int Version { get; set; }

        public WorkflowState Status { get; set; }

        /// <summary>
        /// Generates a header lines for the CSV files via reflection
        /// </summary>
        /// <returns></returns>
        public static string GetHeaderLine()
        {
            return string.Join(";", typeof(Workflow).GetProperties().Select(x => x.Name).ToArray()) + Environment.NewLine;
        }

        /// <summary>
        /// Generates a csv line for the values of the item
        /// </summary>
        /// <returns></returns>
        public string GetValuesLine()
        {
            return string.Join(";", typeof(Workflow).GetProperties().Select(x => x.GetValue(this)).ToArray()) + Environment.NewLine;
        }
    }
}
