using ClosedXML.Excel;
using HkwgConverter.Core;
using System.Configuration;
using System.IO;
using System.Linq;

namespace HkwgConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            var files = Directory.GetFiles(Settings.Default.InputFolder, "*.csv");
            var firstFile = files.FirstOrDefault();
            var inputConverter = new InputConverter(firstFile);

            inputConverter.Execute();
               
        }
    }
}
