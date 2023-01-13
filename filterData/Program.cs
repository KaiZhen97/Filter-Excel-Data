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
            Workbook workbook = new Workbook("excelFile.xlsx");


            //Get the worksheet at first indexed position in the workbook - default worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            AutoFilter filter = worksheet.AutoFilter;
            
            //Create class name = setFilter
            file setFilter = new file();

            //Set the range to which the specified autofilters would be applied
            filter.Range = "firstCell:lastCell";
            //Now add your desired filter to first column to select your desired data
            filter.AddFilter(setFilter.columnNo, setFilter.keyWord);
            filter.Refresh();

            int rowCount = worksheet.Cells.Rows.Count;

            //Delete hidden rows
            for (int i = 1; i <= rowCount; i++)
            {
                if (worksheet.Cells.IsRowHidden(i))
                {
                    worksheet.Cells.DeleteRow(i);
                    i--;
                }
            }

            //Save Excel XLSX file
            workbook.Save("fileName.xlsx");
        }
    }
}
