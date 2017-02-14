using HkwgConverter.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace HkwgConverter.Core
{
    /// <summary>
    /// Encapsulates the work with some metadata about the processing itself. This data is kept in a simple csv file.
    /// </summary>
    public class WorkflowStore
    {
        #region fields

        DirectoryInfo appDataDirectory;
        private string dataFile;

        #endregion

        #region private methods

        private void EnsureFilesExist()
        {
            if (!File.Exists(dataFile))
            {
                File.WriteAllText(dataFile, Workflow.GetHeaderLine());
            }
        }

        private List<Workflow> ReadInputFile()
        {
            var lines = File.ReadAllLines(dataFile)
                .Skip(1)
                .Select(x => x.Split(';'))
                .Select(x => new Model.Workflow()
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

        public WorkflowStore(string appDataPath)
        {
            appDataDirectory = new DirectoryInfo(appDataPath);
            
            this.dataFile = Path.Combine(appDataDirectory.FullName, "workflowDb.dat");

            this.EnsureFilesExist();            
        }

        #endregion

        #region interface
        

        /// <summary>
        /// Retrieves the last version for the given delivery day from the input logfile
        /// </summary>
        /// <param name="deliveryday"></param>
        /// <returns></returns>
        public int GetNextInputVersionNumer(DateTime deliveryday)
        {
            var previousFilesForDay = this.ReadInputFile().Where(x => x.DeliveryDay == deliveryday);
            var nextVersionNumber = 1;

            if(previousFilesForDay.Count() > 0)
            {
                nextVersionNumber = previousFilesForDay.Max(x => x.Version) + 1;
            }

            return nextVersionNumber;                
        }

        public void Add(Workflow item)
        {
            File.AppendAllText(this.dataFile, item.GetValuesLine());
        }

        /// <summary>
        /// Returns the latest Workflow in the system
        /// </summary>
        /// <param name="deliveryDay"></param>
        /// <returns></returns>
        public Workflow GetLatest(DateTime deliveryDay)
        {
            var latestWorkFlow = this.ReadInputFile().Where(x => x.DeliveryDay == deliveryDay && x.Status == WorkflowState.Open)                                                        
                                        .LastOrDefault();

            return latestWorkFlow;
        }
        
        #endregion
    }
}
