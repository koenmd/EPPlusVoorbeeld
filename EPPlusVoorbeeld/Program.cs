using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;

namespace EPPlusVoorbeeld
{
    class Program
    {
        static void Main(string[] args)
        {

            Utils.OutputDir = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}Voorbeeld");

            using (var package = new ExcelPackage(new MemoryStream()))
            {
                var ws1 = package.Workbook.Worksheets.Add("ws1");
                //// Add some values to sum
                //ws1.Cells["A1"].Formula = "(2*2)/2";
                //ws1.Cells["A2"].Formula = "A1+2";
                //ws1.Cells["A3"].Formula = "A2+2";
                //ws1.Cells["A4"].Formula = "SUM(A1:A3)";

                string cell = "";
                string cellvorig = "";
                for (int i = 0; i < 100; i++)
                {
                    //cellvorig = "";
                    for (int j = 0; j < 10; j++)
                    {
                        // bepaal cel
                        string kolom = "A";
                        switch (j)
                        {
                            case 0:
                                kolom = "A";
                                break;
                            case 1:
                                kolom = "B";
                                break;
                            case 2:
                                kolom = "C";
                                break;
                            case 3:
                                kolom = "D";
                                break;
                            case 4:
                                kolom = "E";
                                break;
                            case 5:
                                kolom = "F";
                                break;
                            case 6:
                                kolom = "G";
                                break;
                            case 7:
                                kolom = "H";
                                break;
                            case 8:
                                kolom = "I";
                                break;
                            case 9:
                                kolom = "J";
                                break;
                        }
                        cell = kolom + (i + 1).ToString();

                        if (!string.IsNullOrEmpty(cellvorig))
                            ws1.Cells[cell].Formula = cellvorig + "+1";
                        else
                            ws1.Cells[cell].Value = i + 1;

                        //OfficeOpenXml.Style.ExcelFill fill = OfficeOpenXml.Style.ExcelFill.

                        ws1.Cells[cell].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws1.Cells[cell].Style.Fill.BackgroundColor.SetColor(255, 150+i, j * 25, 0);
                        cellvorig = cell;
                    }
                }


                // calculate all formulas on  the worksheet
                ws1.Calculate();
                ws1.Cells.AutoFitColumns(10, 100);

                //ws1.Column(1).Width = 10;
                //ws1.Column(2).Width = 20;
                //ws1.Column(3).Width = 30;
                //ws1.Column(4).Width = 40;

                //// Print the calculated value
                //Console.WriteLine("SUM(A1:A3) evaluated to {0}", ws1.Cells["A4"].Value);

                //// Add another worksheet
                //var ws2 = package.Workbook.Worksheets.Add("ws2");
                //ws2.Cells["A1"].Value = 3;
                //ws2.Cells["A2"].Formula = "SUM(A1,ws1!A4)";

                //// calculate all formulas in the entire workbook
                //package.Workbook.Calculate();

                //// Print the calculated value
                //Console.WriteLine("SUM(A1,ws1!A4) evaluated to {0}", ws2.Cells["A2"].Value);

                //// Calculate a range
                //ws1.Cells["B1"].Formula = "IF(TODAY()<DATE(2014,6,1),\"BEFORE\" &\" FIRST\",CONCATENATE(\"FIRST\",\" OF\",\" JUNE 2014 OR LATER\"))";
                //ws1.Cells["B1"].Calculate();

                //// Print the calculated value
                //Console.WriteLine("IF(TODAY()<DATE(2014,6,1),\"BEFORE\" &\" FIRST\",CONCATENATE(\"FIRST\",\" OF\",\" JUNE 2014 OR LATER\")) evaluated to {0}", ws1.Cells["B1"].Value);

                //// Evaluate a formula string (without calculate depending cells).
                //// That means that if A1 contains a formula that hasn't been calculated it take the value from a1, blank or zero if it's a new formula.
                //// In this case A1 has been calculated (2), so everything should be ok!
                //const string formula = "(2+4)*ws1!A1";
                //var result = package.Workbook.FormulaParserManager.Parse(formula);

                //// Print the calculated value
                //Console.WriteLine("(2+4)*ws1!A2 evaluated to {0}", result);

                //// Evaluate a formula string (Calculate depending cells)
                //// A1 will be recalculated.
                //var result2 = ws1.Calculate("(2+4)*A1");

                package.SaveAs(Utils.GetFileInfo("VoorbeeldFormule.xlsx"));
            }

        }
    }
}
