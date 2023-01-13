using System;
using System.IO;
using System.Linq;

using Aspose.Cells; // aspose fast
using Workbook = Aspose.Cells.Workbook;
using Worksheet = Aspose.Cells.Worksheet;

//using Spire.Xls.Collections; // spire slow
//using Workbook = Spire.Xls.Workbook;
//using Worksheet = Spire.Xls.Worksheet;

namespace filterData
{
    class Program
    {
        static void Main(string[] args)
        {
            //Aspose better, faster
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
            //workbook.Save(@"C:\Users\kaizhen.goh\source\report\FilterCrystalReportViewerAspose.xlsx");


            ////spire too slow
            //// create workbook, load existing file, get first worksheet
            //Workbook workbook = new Workbook();
            ////workbook.LoadFromFile(@"C:\Users\kaizhen.goh\source\report\CrystalReportViewer-From11Nov2021.xlsx");
            //workbook.LoadFromFile(@"C:\Users\kaizhen.goh\source\report\CrystalReportViewer1 (6) - Copy.xlsx");
            //Worksheet worksheet = workbook.Worksheets[0];

            //// filter
            //AutoFiltersCollection filters = worksheet.AutoFilters;
            //filters.Range = worksheet.Range["A1:P15044"];
            //filters.AddFilter(2, "YN3210X");
            //filters.Filter();

            //int rowCount = worksheet.Rows.Count();

            //for (int i = 1; i <= rowCount; i++)
            //{
            //    if (worksheet.GetRowIsHide(i))
            //    {
            //        worksheet.DeleteRow(i);
            //        i--;
            //    }
            //}

            //// save file
            //workbook.SaveToFile(@"C:\Users\kaizhen.goh\source\report\DeleteCrystalReportViewerSpireYN3210X.xlsx");
            ////workbook.SaveToFile(@"C:\Users\kaizhen.goh\source\report\FilterCrystalReportViewerSpire.xlsx");
        }
    }
}
