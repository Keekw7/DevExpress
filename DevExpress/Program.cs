using DevExpress.Export.Xl;
using System.Diagnostics;
using System.Drawing;
using DevExpress.Spreadsheet;
using Worksheet = DevExpress.Spreadsheet.Worksheet;
using CellRange = DevExpress.Spreadsheet.CellRange;
using HeaderFooterSection = DevExpress.Spreadsheet.HeaderFooterSection;

namespace XLExportExamples
{
    public class Record
    {
        public string ProductNumber { get; set; }
        public string ProductDescription { get; set; }
        public string[] VdwNumbers { get; set; }
        public string[] Description { get; set; }
        public int[] Quantity { get; set; }
        public string Currency { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            PopulateExcel();
        }

        static void PopulateExcel()
        {
            string filePath = @"Document.xlsx";
            using (var workbook = new Workbook())
            {
                // Add a new worksheet to the document.
                Worksheet sheet = workbook.Worksheets[0];
                // Specify the worksheet name.
                sheet.Name = "Sales report";

                var imagePathLeft = "C:\\Users\\KHAKEE\\source\\repos\\DevExpress\\DevExpress\\Assets\\vandewiele-logo.png";
                var imagePathRight = "C:\\Users\\KHAKEE\\source\\repos\\DevExpress\\DevExpress\\Assets\\roj-logo.jpg";

                // Specify Header and Footer
                WorksheetHeaderFooterOptions options = sheet.HeaderFooterOptions;
                options.DifferentFirst = true;
                HeaderFooterPicture leftHeaderPicture = options.FirstHeader.AddPicture(imagePathLeft, HeaderFooterSection.Left);
                leftHeaderPicture.Height = 500;
                leftHeaderPicture.Width = 500;

                HeaderFooterPicture rightHeaderPicture = options.FirstHeader.AddPicture(imagePathRight, HeaderFooterSection.Right);
                rightHeaderPicture.Height = 300;
                rightHeaderPicture.Width = 300;

                // Specify cell font attributes.
                XlCellFormatting cellFormatting = new XlCellFormatting();
                cellFormatting.Font = new XlFont();
                cellFormatting.Font.Name = "Century Gothic";
                cellFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                var list = new List<Record>()
                {
                    new Record()
                    {
                        ProductNumber = "SE601258",
                        ProductDescription = "描述描述描述描述描述描述描述描述描述描述描述描述描述描述描述描述描述",
                        VdwNumbers = new []{ "8028.08.42", "26B34.01", "54R55.01" },
                        Description = new[] { "STAR G2张力器" ,"鳄鱼嘴张力器", "双叶片张力器", },
                        Quantity = new[] { 1,3,2 },
                        Price = 1800,
                        Currency = "CNY"
                    },
                    new Record()
                    {
                        ProductNumber = "SE601259",
                        ProductDescription = "描述描述描述描述描述描述描述描述描述描述描述描述描述描述描述描述描述",
                        VdwNumbers = new []{ "8828.08.40.N" },
                        Description = new[] { "引纬送经装置" },
                        Quantity = new[] { 1 },
                        Price = 600,
                        Currency = "CNY"
                    },
                    new Record()
                    {
                        ProductNumber = "SE601290",
                        ProductDescription = "描述描述描述描述描",
                        VdwNumbers = new []{ "56C654", "A0001", "LCSW_G3 plus" },
                        Description = new[] { "电控箱", "OEMProduct1", "4色卧式纱架+储纬器架一体 4c creel and stand" },
                        Quantity = new[] { 1,1,1 },
                        Price = 3000,
                        Currency = "CNY"
                    }
                };

                CellRange range = sheet.Range[$"A2:F2"];
                range.Merge();
                sheet.Cells["A2"].Value = "客户名称 ： 勿删白标公司"; //replace with cusotmer 
                sheet.Cells["A2"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                sheet.Cells["A2"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;

                sheet.Cells["A4"].Value = "序列";
                sheet.Cells["B4"].Value = "客户物料编号";
                sheet.Cells["C4"].Value = "客户物料描述";
                sheet.Cells["D4"].Value = "VDW物料编号";
                sheet.Cells["E4"].Value = "物料描述";
                sheet.Cells["F4"].Value = "数量";
                sheet.Cells["G4"].Value = "单价";

                var count = 1;
                foreach (var item in list)
                {
                    var currentNumber = sheet.GetUsedRange().RowCount + 2;
                    PopulateRecord(item, sheet, currentNumber, count);
                    count++;
                }

                var pageSetup = sheet.PrintOptions;
                pageSetup.FitToPage = true;
                pageSetup.FitToWidth = 1;
                pageSetup.FitToHeight = 1;
                pageSetup.Scale = 100;

                workbook.SaveDocument(filePath, DocumentFormat.Xlsx);
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }

        static void PopulateRecord(Record record, Worksheet sheet, int currentNumber, int count)
        {
            // Sample data
            string[] vdwNumbers = record.VdwNumbers;
            string[] description = record.Description;
            int[] quantities = record.Quantity;

            // Insert VDW numbers in merged cells
            for (int i = 0; i < vdwNumbers.Length; i++)
            {
                sheet.Cells[$"D{currentNumber + i}"].Value = vdwNumbers[i];
                sheet.Cells[$"D{currentNumber + i}"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                sheet.Cells[$"D{currentNumber + i}"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            }

            //Insert Description in merged cells
            for (int i = 0; i < description.Length; i++)
            {
                sheet.Cells[$"E{currentNumber + i}"].Value = description[i];
                sheet.Cells[$"E{currentNumber + i}"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                sheet.Cells[$"E{currentNumber + i}"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            }

            //Insert Description in merged cells
            for (int i = 0; i < quantities.Length; i++)
            {
                sheet.Cells[$"F{currentNumber + i}"].Value = quantities[i];
                sheet.Cells[$"F{currentNumber + i}"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                sheet.Cells[$"F{currentNumber + i}"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            }

            // Merge and set the description cell
            MergeCell(sheet, count.ToString(), "A", currentNumber, currentNumber + vdwNumbers.Length - 1);
            MergeCell(sheet, record.ProductNumber, "B", currentNumber, currentNumber + vdwNumbers.Length - 1);
            MergeCell(sheet, record.ProductDescription, "C", currentNumber, currentNumber + vdwNumbers.Length - 1);
            MergeCell(sheet, $"{record.Currency} {record.Price:N2}", "G", currentNumber, currentNumber + vdwNumbers.Length - 1);


            // Adjust column widths for better visibility
            sheet.Columns[0].WidthInCharacters = GetColumnLength(count.ToString().Length);
            sheet.Columns[1].WidthInCharacters = GetColumnLength(record.ProductNumber.Length);
            sheet.Columns[2].WidthInCharacters = GetColumnLength(record.ProductDescription.Length);
            sheet.Columns[3].WidthInCharacters = GetColumnLength(vdwNumbers.Max(f => f.Length));
            sheet.Columns[4].WidthInCharacters = GetColumnLength(description.Max(f => f.Length));
            sheet.Columns[5].WidthInCharacters = GetColumnLength(quantities.Max(f => f.ToString().Length));
            sheet.Columns[6].WidthInCharacters = GetColumnLength(($"{record.Currency} {record.Price:N2}").Length);


            var usedRange = sheet.GetUsedRange();
            var range = sheet.Range[$"A4:G{usedRange.RowCount + 1}"];
            var rangeFormatting = range.BeginUpdateFormatting();
            rangeFormatting.Alignment.WrapText = true;
            rangeFormatting.Borders.SetAllBorders(Color.Black, DevExpress.Spreadsheet.BorderLineStyle.Thin);
            range.EndUpdateFormatting(rangeFormatting);
        }

        static void MergeCell(Worksheet sheet, string value, string column, int startNumber, int endNumber)
        {
            CellRange range = sheet.Range[$"{column}{startNumber}:{column}{endNumber}"];
            range.Merge();
            sheet.Cells[$"{column}{startNumber}"].Value = value;
            sheet.Cells[$"{column}{startNumber}"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
            sheet.Cells[$"{column}{startNumber}"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
        }

        static int GetColumnLength(int length)
        {
            if (length >= 30)
                return 40;

            if (length <= 5)
                return 10;

            if (length <= 10)
                return 20;

            return length + 5;
        }
    }
}
