using System;
using System.IO;
using System.Linq;

using Aspose.Cells;
using Workbook = Aspose.Cells.Workbook;
using Worksheet = Aspose.Cells.Worksheet;

namespace filterData
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create an empty Excel workbook
            //Workbook workbook = new Workbook(@"C:\Users\kaizhen.goh\source\report\CrystalReportViewer-From11Nov2021.xlsx");
            Workbook workbook = new Workbook(@"C:\Users\kaizhen.goh\source\report\CrystalReportViewer1 (6).xlsx");


            //Get the worksheet at first indexed position in the workbook - default worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            AutoFilter filter = worksheet.AutoFilter; // not necessary

            //Set the range to which the specified autofilters would be applied
            filter.Range = "A1:P15044";
            //Now add your desired filter to first column to select your desired data
            filter.AddFilter(2, "YN3267H");
            filter.Refresh();


            int rowCount = worksheet.Cells.Rows.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                if (worksheet.Cells.IsRowHidden(i))
                {
                    worksheet.Cells.DeleteRow(i);
                    i--;
                }
            }

            //Save Excel XLSX file
            workbook.Save(@"C:\Users\kaizhen.goh\source\report\DeleteCrystalReportViewerAsposeYN3267H.xlsx");
        }
    }
}
