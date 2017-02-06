using HkwgConverter.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HkwgConverter.Core
{
    /// <summary>
    /// Encapsulates the work with some metadata about the processing itself. This data is kept in a simple csv file.
    /// </summary>
    internal class AppDataAccessor
    {
        #region fields

        DirectoryInfo appDataDirectory;
        private string inputLogFile;

        #endregion

        #region private methods

        private void EnsureFilesExist()
        {
            if (!File.Exists(inputLogFile))
            {
                File.WriteAllText(inputLogFile, AppDataInputItem.GetHeaderLine());
            }
        }

        private List<AppDataInputItem> ReadInputFile()
        {
            var lines = File.ReadAllLines(inputLogFile)
                .Skip(1)
                .Select(x => x.Split(';'))
                .Select(x => new Model.AppDataInputItem()
                {
                    Timestamp = DateTime.Parse(x[0]),
                    DeliveryDay = DateTime.Parse(x[1]),
                    CsvFile = x[2],
                    FlexPosFile = x[3],
                    FlexNegFile = x[4],
                    Version = int.Parse(x[5])
                });

            return lines.ToList();
        }

        #endregion

        #region ctor

        internal AppDataAccessor(string appDataPath)
        {
            appDataDirectory = new DirectoryInfo(appDataPath);
            
            this.inputLogFile = Path.Combine(appDataDirectory.FullName, "inputlog.csv");

            this.EnsureFilesExist();            
        }

        #endregion

        #region interface
        

        /// <summary>
        /// Retrieves the last version for the given delivery day from the input logfile
        /// </summary>
        /// <param name="deliveryday"></param>
        /// <returns></returns>
        internal int GetNextInputVersionNumer(DateTime deliveryday)
        {
            var previousFilesForDay = this.ReadInputFile().Where(x => x.DeliveryDay == deliveryday);
            var nextVersionNumber = 1;

            if(previousFilesForDay.Count() > 0)
            {
                nextVersionNumber = previousFilesForDay.Max(x => x.Version) + 1;
            }

            return nextVersionNumber;                
        }

        internal void AppendInputLogItem(AppDataInputItem item)
        {
            File.AppendAllText(this.inputLogFile, item.GetValuesLine());
        }

        #endregion
    }
}
