/*Quantity2 - software for data analysis
 Copyright(C) 2018  Georgi Danovski

 This program is free software: you can redistribute it and/or modify
 it under the terms of the GNU General Public License as published by
 the Free Software Foundation, either version 3 of the License, or
 (at your option) any later version.

 This program is distributed in the hope that it will be useful,
 but WITHOUT ANY WARRANTY; without even the implied warranty of
 MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
 GNU General Public License for more details.

 You should have received a copy of the GNU General Public License
 along with this program.If not, see<http://www.gnu.org/licenses/>.*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace Quantity2
{
    class ExcelCode
    {
        public static void CreateWorkbook(string suf, int MaxColumnLength, string dir, int MinG, int MaxG, int MinR, int MaxR)
        {


            var dirs = FileReader.DigForFiles(dir, suf);

            // creating Excel Application
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            Sheets xlSheets = null;

            xlSheets = workbook.Sheets as Sheets;

            // see the excel sheet behind the program
            app.Visible = false;

            //Select the sheet
            worksheet = workbook.Worksheets[1];
            //Rename the sheet
            worksheet.Name = "Summary";

            string[] names = GetNames(dirs);

            Console.WriteLine("Processing Files:");

            for (int i = 0; i < names.Length; i++)
            {
                Console.WriteLine(names[i]);
                Microsoft.Office.Interop.Excel._Workbook csvWorkbook = app.Workbooks.Open(dirs[i]);
                Microsoft.Office.Interop.Excel._Worksheet worksheetCSV = ((Microsoft.Office.Interop.Excel._Worksheet)csvWorkbook.Worksheets[1]);

                worksheetCSV.Copy(xlSheets[1]);
                xlSheets[1].Name = names[i];

                ((_Worksheet)xlSheets[1]).Cells[1, 24] = "Mean_Cell - Mean_Noise";
                for (int row = 2; row < MaxColumnLength; row++)
                {
                    ((_Worksheet)xlSheets[1]).Cells[row, 24] = "=E" + row + "-U" + row;
                }
                ((_Worksheet)xlSheets[1]).Cells[1, 25] = "Max_Cell";
                ((_Worksheet)xlSheets[1]).Cells[2, 25] = "=MAX(X2:X" + MaxColumnLength + ")";
                // Exit from the application
                csvWorkbook.Close();
            }

            worksheet.Move(Before: workbook.Sheets[1]);
            Console.WriteLine("Preparing the summary...");
            //worksheet.Cells[row, column] = "=cell57_Q2_Ch0_Green_Results!E2";

            int currentColumn = 1;
            int currentRow = 1;
            int increment = names.Length;

            for (currentRow = 1; currentRow < MaxColumnLength; currentRow++)
                worksheet.Cells[currentRow, currentColumn] = "=" + names[0] + "!C" + currentRow;

            currentRow = 1;

            for (int i = 0; i < names.Length; i++)
            {
                //Area cell
                currentColumn = i + 2;

                worksheet.Cells[currentRow, currentColumn] = "Area_Cell_" + names[i];

                //Mean cell
                currentColumn += increment;

                worksheet.Cells[currentRow, currentColumn] = "Mean_Cell_" + names[i];
                //Area noise
                currentColumn += increment;

                worksheet.Cells[currentRow, currentColumn] = "Area_Noise_" + names[i];
                //Mean noise
                currentColumn += increment;

                worksheet.Cells[currentRow, currentColumn] = "Mean_Noise_" + names[i];
                //Area spots
                currentColumn += increment;

                worksheet.Cells[currentRow, currentColumn] = "Area_Spot_" + names[i];
                //Mean spots
                currentColumn += increment;

                worksheet.Cells[currentRow, currentColumn] = "Mean_Spot_" + names[i];

                //Mean-Noise cell
                currentColumn += increment;

                worksheet.Cells[currentRow, currentColumn] = "Mean-Noise_Cell_" + names[i];
                //Max cell
                currentColumn += increment;

                worksheet.Cells[currentRow, currentColumn] = "Max_Cell_" + names[i];
                worksheet.Cells[currentRow + 1, currentColumn] = "=" + names[i] + "!Y2";
                //Max-Results
                currentColumn += increment;

                worksheet.Cells[currentRow, currentColumn] = "Max-Results_" + names[i];

                //Results
                currentColumn += increment;

                worksheet.Cells[currentRow, currentColumn] = "Results_" + names[i];
                //ResultsTo0
                currentColumn += increment;

                worksheet.Cells[currentRow, currentColumn] = "To0_Results_" + names[i];
                
                //ResultsTo1
                currentColumn += increment;

                worksheet.Cells[currentRow, currentColumn] = "nTo0_Results_" + names[i];

            }
            //Mean result
            
            worksheet.Cells[currentRow, currentColumn + 1] = "nAvgMob";
            worksheet.Cells[currentRow, currentColumn + 2] = "nnAvgMob";
            worksheet.Cells[currentRow, currentColumn + 3] = "nStDevMob";
            worksheet.Cells[currentRow, currentColumn + 4] = "AvgMob";
            worksheet.Cells[currentRow, currentColumn + 5] = "StDevMob";

            int spotSignal, MaxCell, CellMinusNoise;

            for (currentRow = 2; currentRow < MaxColumnLength; currentRow++)
                for (int i = 0; i < names.Length; i++)
                {
                    //Area cell
                    currentColumn = i + 2;

                    worksheet.Cells[currentRow, currentColumn] = "=" + names[i] + "!D" + currentRow;

                    //Mean cell
                    currentColumn += increment;

                    worksheet.Cells[currentRow, currentColumn] = "=" + names[i] + "!E" + currentRow;
                    //Area noise
                    currentColumn += increment;

                    worksheet.Cells[currentRow, currentColumn] = "=" + names[i] + "!T" + currentRow;
                    //Mean noise
                    currentColumn += increment;

                    worksheet.Cells[currentRow, currentColumn] = "=" + names[i] + "!U" + currentRow;
                    //Area spots
                    currentColumn += increment;

                    worksheet.Cells[currentRow, currentColumn] =
                        "=AVERAGE(" + names[i] + "!H" + currentRow + ","
                        + names[i] + "!L" + currentRow + ","
                        + names[i] + "!P" + currentRow + ")";
                    //Mean spots
                    currentColumn += increment;

                    worksheet.Cells[currentRow, currentColumn] =
                        "=AVERAGE(" + names[i] + "!I" + currentRow + ","
                        + names[i] + "!M" + currentRow + ","
                        + names[i] + "!Q" + currentRow + ")";

                    spotSignal = currentColumn;
                    //Cell-Noise mean
                    currentColumn += increment;

                    worksheet.Cells[currentRow, currentColumn] = "=" + names[i] + "!X" + currentRow;
                    CellMinusNoise = currentColumn;
                    //Max Cell
                    currentColumn += increment;
                    MaxCell = currentColumn;
                    //Min-Results

                    currentColumn += increment;

                    worksheet.Cells[currentRow, currentColumn] = "=" + ColumnLabel(MaxCell) + "2-(" + ColumnLabel(MaxCell) + "2/" +
                        ColumnLabel(CellMinusNoise) + currentRow + ")*(" + ColumnLabel(spotSignal) + currentRow + "-" +
                        names[i] + "!U" + currentRow +
                        ")";
                    //Results

                    currentColumn += increment;

                    worksheet.Cells[currentRow, currentColumn] = "=(" + ColumnLabel(MaxCell) + "2/" +
                        ColumnLabel(CellMinusNoise) + currentRow + ")*(" + ColumnLabel(spotSignal) + currentRow + "-" +
                        names[i] + "!U" + currentRow +
                        ")";
                }

            //calculations

            currentColumn++;
            int MaxMinResults = 2 + 8 * increment;
            int ResultsTo0 = currentColumn;

            for (int i = 0; i < names.Length; i++)
            {
                for (currentRow = 2; currentRow < MaxColumnLength; currentRow++)
                {
                    if (suf == "Green")
                    {
                        worksheet.Cells[currentRow, currentColumn] = "=(" +
                            ColumnLabel(MaxMinResults) + currentRow +
                            "-AVERAGE(" + ColumnLabel(MaxMinResults) + MinG + ":" + ColumnLabel(MaxMinResults) + MaxG + ")" +
                            ")";
                    }
                    else
                    {
                        worksheet.Cells[currentRow, currentColumn] = "=(" +
                            ColumnLabel(MaxMinResults) + currentRow +
                            "-AVERAGE(" + ColumnLabel(MaxMinResults) + MinR + ":" + ColumnLabel(MaxMinResults) + MaxR + ")" +
                            ")";
                    }
                }
                currentColumn++;
                MaxMinResults++;
            }

            //NormTo1
            MaxMinResults = currentColumn - increment;
            for (int i = 0; i < names.Length; i++)
            {
                for (currentRow = 2; currentRow < MaxColumnLength; currentRow++)
                {
                    worksheet.Cells[currentRow, currentColumn] = "=(" +
                        ColumnLabel(MaxMinResults) + currentRow +
                        "/MAX(" + ColumnLabel(MaxMinResults) + "2" + ":" + ColumnLabel(MaxMinResults) + MaxColumnLength + ")" +
                        ")";
                }
                currentColumn++;
                MaxMinResults++;
            }

            {
                int MinInd = currentColumn - increment;
                int MaxInd = currentColumn-1;
                int Mean = currentColumn;
                int nMean = currentColumn + 1;
                int StDev = currentColumn + 2;

                for (currentRow = 2; currentRow < MaxColumnLength; currentRow++)
                {
                    worksheet.Cells[currentRow, Mean] = "=" +
                        "AVERAGE(" + ColumnLabel(MinInd) + currentRow + ":" + ColumnLabel(MaxInd) + currentRow + ")";

                    worksheet.Cells[currentRow, StDev] = "=" +
                        "STDEV.S(" + ColumnLabel(MinInd) + currentRow + ":" + ColumnLabel(MaxInd) + currentRow + ")";
                }
                for (currentRow = 2; currentRow < MaxColumnLength; currentRow++)
                {
                    worksheet.Cells[currentRow, nMean] = "=(" +
                                ColumnLabel(Mean) + currentRow +
                                "/MAX(" + ColumnLabel(Mean) + "2:" + ColumnLabel(Mean) + MaxColumnLength + ")" +
                                ")";
                }

                 MinInd -= increment;
                 MaxInd -= increment;
                Mean = currentColumn+3;
                StDev = currentColumn + 4;

                for (currentRow = 2; currentRow < MaxColumnLength; currentRow++)
                {
                    worksheet.Cells[currentRow, Mean] = "=" +
                        "AVERAGE(" + ColumnLabel(MinInd) + currentRow + ":" + ColumnLabel(MaxInd) + currentRow + ")";

                    worksheet.Cells[currentRow, StDev] = "=" +
                        "STDEV.S(" + ColumnLabel(MinInd) + currentRow + ":" + ColumnLabel(MaxInd) + currentRow + ")";
                }
            }
            string name = dir.Substring(dir.LastIndexOf("\\") + 1, dir.Length - dir.LastIndexOf("\\") - 1);
            name = dir + "\\" + "Res_" + suf + ".xlsx";

            app.Visible = true;
            
            workbook.SaveAs(name, Type.Missing,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Console.WriteLine("Results saved to :\n" + name);
            app.Quit();//Check is working
            Console.WriteLine("Done!");
        }

        private static string[] GetNames(List<string> dirs)
        {
            string[] names = new string[dirs.Count];

            for (int i = 0; i < names.Length; i++)
            {
                string dir = dirs[i];
                string name = dir.Substring(dir.LastIndexOf("\\") + 1, dir.LastIndexOf(".") - dir.LastIndexOf("\\") - 1);

                name = name.Replace("_CompositeRegistred", "");

                if (!names.Contains(name))
                {
                    names[i] = name;
                }
                else
                {
                    int ind = 1;
                    while (names.Contains(name + "_" + ind))
                    {
                        ind++;
                    }
                    names[i] = name + "_" + ind;
                }
            }
            return names;
        }

        public static string ColumnLabel(int col)
        {
            var dividend = col;
            var columnLabel = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnLabel = Convert.ToChar(65 + modulo).ToString() + columnLabel;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnLabel;
        }
        public static int ColumnIndex(string colLabel)
        {
            // "AD" (1 * 26^1) + (4 * 26^0) ...
            var colIndex = 0;
            for (int ind = 0, pow = colLabel.Count() - 1; ind < colLabel.Count(); ++ind, --pow)
            {
                var cVal = Convert.ToInt32(colLabel[ind]) - 64; //col A is index 1
                colIndex += cVal * ((int)Math.Pow(26, pow));
            }
            return colIndex;
        }

    }
}
