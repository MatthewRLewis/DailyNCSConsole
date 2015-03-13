using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;
using System.IO;
using System.Globalization;
using System.Drawing;
using JayMuntzCom;


namespace DailyNCSConsole
{
    class EPHelper
    {

        public static void GenerateExcel(DataSet dtPrinter, DataSet dtCopier, String fileName)
        {
            string currentDirectorypath = Environment.CurrentDirectory + "\\Resources\\tempfiles\\";
            string finalFileNameWithPath = string.Empty;

            finalFileNameWithPath = string.Format("{0}\\{1}.xlsx", currentDirectorypath, fileName);

            //Delete existing file with same file name (unlikely, as this is done in Program.cs.
            if (File.Exists(finalFileNameWithPath))
                File.Delete(finalFileNameWithPath);

            var newFile = new FileInfo(finalFileNameWithPath);

            //The ExcelPackage class and pass file path to constructor.
            using (var package = new ExcelPackage(newFile))
            {
                sheetMaker(dtPrinter, package, "Printer");
                sheetMaker(dtCopier, package, "Copier");

                //File properties
                package.Workbook.Properties.Title = fileName;
                package.Workbook.Properties.Author = "DailyNCSConsole";
                //optional subject
                //package.Workbook.Properties.Subject = 



                //Save changes to ExcelPackage object which will create the spreadsheet.
                package.Save();

            }
         
        }
        public static void sheetMaker(DataSet dtS, ExcelPackage package, string excelSheetName)
        {
            //set formatting for display in spreadsheet
            string rowStyle = "$###,###,##0";

            DateTime FirstDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime firstOfNextMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(1);
            DateTime dateNow = firstOfNextMonth.AddDays(-1);
            double busDays = GetBusinessDays(FirstDate, dateNow);
            //create holiday INT
            int holidays = 0;
            //This DLL will return the number of holidays in the given month, XML file updated by HR
            HolidayCalculator hc = new HolidayCalculator(FirstDate, @"\\nwgdeploy\Scripts\HolidayCalculator\Holidays.xml", dateNow);
            foreach (HolidayCalculator.Holiday h in hc.OrderedHolidays)
            {
                holidays++;
            }

            int intBusHol = (Convert.ToInt16(busDays) - holidays);

            int colFin = 0;

            string avFinCol = "";
            string finCol = "";

            //spreadsheet begins at row 2
            int counter = 2;
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month) +  " " + excelSheetName);
            foreach (DataTable dt in dtS.Tables)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (string.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            //row is empty, convert it to a 0
                            dt.Rows[i][j] = 0.00M;
                        }
                        else
                        {
                           //row is NOT empty, convert it to a decimal
                            dt.Rows[i][j] = Decimal.ToInt32(Convert.ToDecimal(dt.Rows[i][j]));
                        }
                    }
                }

                int colCount = dt.Columns.Count;
                string tableName = dt.TableName.ToString();
               
                //grab the budget target amount from the settings file
                decimal target;
                //check to see if we're dealing with a copier, or a printer table
                if (tableName.IndexOf("Printer") != -1)
                {
                    target = NCSSettings.Default.PrinterTarget;
                }
                else
                {
                    target = NCSSettings.Default.CopierTarget;
                }

                decimal QBSTarget = Math.Round((target * NCSSettings.Default.QBSPerc), 0);
                decimal CTXTarget = (target - QBSTarget);


                string calcPlace = "A" + counter;
                string tablePlace = "A" + (counter + 2) + ":" + GetColumnName(colCount - 1) + (counter + 2);
                var calcCol = worksheet.Cells[calcPlace];
                calcCol.Value = tableName;

                var Col1 = worksheet.Cells["E" + counter];
                Col1.Value = "Target :";
               
                //Check the table name to see which regional target will be applied
                decimal col2Val;
                if (tableName.IndexOf("QBS") != -1)
                {
                    col2Val = QBSTarget;
                }
                else if (tableName.IndexOf("CTX") != -1)
                {
                    col2Val = CTXTarget;
                }
                else
                {
                    col2Val = target;
                }

                var Col2 = worksheet.Cells["F" + counter];
                Col2.Value = col2Val;
                var Col3 = worksheet.Cells["G" + counter];
                Col3.Value = "/";
                Col3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                var Col4 = worksheet.Cells["H" + counter];
                Col4.Value = intBusHol;
                var Col5 = worksheet.Cells["I" + counter];
                Col5.Value = "=";
                Col5.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                var Col6 = worksheet.Cells["J" + counter];
                Col6.Formula = "F" + counter + "/" + "H" + counter;
                Col6.Style.Numberformat.Format = rowStyle;
                var Col7 = worksheet.Cells["K" + counter];
                Col7.Value = "per day";


                string topRange = calcPlace + ":" + "K" + counter;
                blckFormatter(worksheet, topRange);

                //add one to the counter
                counter++;
                
                //using GetColumnName, find the size of the table
                string formulaRows = "A" + (counter + 1) + ":" + GetColumnName(colCount -1 ) + (counter + 1);

                finCol = (GetColumnName(colCount) + (counter + 1)).ToString();
                string sheytPlace = (GetColumnName(colCount) + (counter + 1)).ToString();
                var sheyt = worksheet.Cells[sheytPlace];
                   sheyt.Formula = "SUM(" + formulaRows + ")";
                   greyFormatter(worksheet, sheytPlace);
                   sheyt.Style.Numberformat.Format = rowStyle;
                
                    
                string totalPlace = (GetColumnName(colCount) + (counter)).ToString();
                var totalCol = worksheet.Cells[totalPlace];
                totalCol.Value = "Total";
                blckFormatter(worksheet, totalPlace);


                avFinCol = (GetColumnName(colCount + 1) + (counter + 1)).ToString();
                string shoytPlace = (GetColumnName(colCount + 1) + (counter + 1)).ToString();
                var shoyt = worksheet.Cells[shoytPlace];
                   shoyt.Formula = "AVERAGEIF(" + formulaRows + ", \">0\")";
                   greyFormatter(worksheet, shoytPlace);
                   shoyt.Style.Numberformat.Format = rowStyle;


                   string averagePlace = (GetColumnName(colCount + 1) + (counter)).ToString();
                   var AverageCol = worksheet.Cells[averagePlace];
                   AverageCol.Value = "Average";
                   blckFormatter(worksheet, averagePlace);


                   calcPlace = "A" + counter;
                   worksheet.Cells[calcPlace].LoadFromDataTable(dt, true, TableStyles.Medium1);
                   worksheet.Cells[tablePlace].Style.Numberformat.Format = rowStyle;

            
                 List<string> rown  = new List<string>();
                   for (int i = 0; i < colCount; i++)
                   {
                        rown.Add(GetColumnName(i) + (counter + 1));   
                   }

                condFormat(worksheet, rown, (col2Val / Convert.ToDecimal(intBusHol)).ToString());
                //add 4 for spacing
                counter += 4;
                colFin = colCount;
            }

            //Add text
            string workPlace = ("A" + counter).ToString();
            var workCol = worksheet.Cells[workPlace];
            workCol.Value = "Work days remaining: ";
            blckFormatter(worksheet, workPlace);

            //Calcualte the number of work days remaining in the month, accounting for holidays
            string calcprojPlace = ("B" + counter).ToString();
            var calcprojCol = worksheet.Cells[calcprojPlace];
            int busDaysToday = Convert.ToInt16(GetBusinessDays(FirstDate, DateTime.Now));
            int daysLeft = (intBusHol - busDaysToday);
            calcprojCol.Value = (daysLeft + 1);
            calcprojCol.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            blckFormatter(worksheet, calcprojPlace); 

            string projPlace = ("C" + counter).ToString();
            var projCol = worksheet.Cells[projPlace];
            projCol.Value = "Monthly Projected: ";
            blckFormatter(worksheet, projPlace);


            
            string finPlace = ("D" + counter).ToString();
            var fin = worksheet.Cells[finPlace];
            fin.Formula = "SUM(" + finCol + "," + avFinCol + " * " + calcprojPlace + ")";
            blckFormatter(worksheet, finPlace);
            fin.Style.Numberformat.Format = rowStyle;
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

          

        }


        static void blckFormatter(ExcelWorksheet worksheet, string loc)
        {
            var totalCol = worksheet.Cells[loc];
            totalCol.Style.Fill.PatternType = ExcelFillStyle.Solid;
            totalCol.Style.Fill.BackgroundColor.SetColor(Color.Black);
            totalCol.Style.Font.Color.SetColor(Color.White);
            totalCol.Style.Font.Bold = true;
        }

        static void greyFormatter(ExcelWorksheet worksheet, string loc)
        {
            var totalCol = worksheet.Cells[loc];
            totalCol.Style.Fill.PatternType = ExcelFillStyle.Solid;
            totalCol.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            totalCol.Style.Font.Color.SetColor(Color.Black);
            totalCol.Style.Font.Bold = true;
        }
        static void condFormat(ExcelWorksheet worksheet, List<string> expressRange, string finVal)
        {
            //conditional formatting for any rows that exceed budget
            foreach (string e in expressRange)
            {
                var sheet = worksheet.Cells[e];
                var _cond1 = sheet.ConditionalFormatting.AddGreaterThan();
                _cond1.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                _cond1.Style.Fill.BackgroundColor.Color = Color.Red;
                _cond1.Formula = finVal;
            }
        }

        static string GetColumnName(int index)
        {
            //uses the index of the string to find the excel position
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }

        public static double GetBusinessDays(DateTime startD, DateTime endD)
            //calculates the number of working days by subtracting weekends using DayOfWeek's position
        {
            double calcBusinessDays =
                1 + ((endD - startD).TotalDays * 5 -
                (startD.DayOfWeek - endD.DayOfWeek) * 2) / 7;

            if ((int)endD.DayOfWeek == 6) calcBusinessDays--;
            if ((int)startD.DayOfWeek == 0) calcBusinessDays--;

            return calcBusinessDays;
        }
    }
}
