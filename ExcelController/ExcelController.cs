using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Collections.Generic;
using System;
using System.IO;
using System.Drawing;

namespace Fresh
{
    class ExcelController
    {
        public ExcelController()
        {

        }

        #region Generic Excel Functions
        private object missing = System.Reflection.Missing.Value;

        private static void ExcelDisableCalculations(Excel.Application app)
        {
            app.ScreenUpdating = false;
            app.Calculation = Excel.XlCalculation.xlCalculationManual;
        }

        private static void ExcelEnableCalculations(Excel.Application app)
        {
            app.ScreenUpdating = true;
            app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        }

        public static void AddBordersToRange(Excel.Range range, bool darkMarjorBorder)
        {
            int innerBorderweight = 2;
            int edgeBorderweight = 3;

            if (!darkMarjorBorder)
                edgeBorderweight = 2;

            if (range.Rows.Count > 1)
            {
                range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Color = Color.Red.ToArgb();
                range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Weight = innerBorderweight;
            }

            if (range.Columns.Count > 1)
            {
                range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Color = Color.Black.ToArgb();
                range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Weight = innerBorderweight;
            }

            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Color = Color.Black.ToArgb();
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Color = Color.Black.ToArgb();
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Color = Color.Black.ToArgb();
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Color = Color.Black.ToArgb();
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = edgeBorderweight;
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = edgeBorderweight;
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = edgeBorderweight;
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = edgeBorderweight;
        }

        private static void WriteDataArrayToSheet(object[,] data, Point startCellIndex, Excel.Worksheet worksheet)
        {
            //var data = new object[rows, columns];
            //for (var row = 1; row <= rows; row++)
            //{
            //    for (var column = 1; column <= columns; column++)
            //    {
            //        data[row - 1, column - 1] = "Test";
            //    }
            //}

            Point endCellIndex = new Point(startCellIndex.X + data.GetLength(0) - 1, startCellIndex.Y + data.GetLength(1) - 1);

            var startCell = (Excel.Range)worksheet.Cells[startCellIndex.X, startCellIndex.Y];
            var endCell = (Excel.Range)worksheet.Cells[endCellIndex.X, endCellIndex.Y];
            var writeRange = worksheet.Range[startCell, endCell];

            writeRange.Value2 = data;
        }

        public static bool ValidXLSFile(ref string filePath)
        {
            string appPath = Path.GetDirectoryName(Application.ExecutablePath);

            bool validFile = false;
            if (filePath != null)
            {
                if (File.Exists(filePath))
                {
                    validFile = true;
                }
                else
                {
                    filePath = appPath + Path.DirectorySeparatorChar + filePath;
                    if (File.Exists(filePath))
                    {
                        validFile = true;
                    }
                }
            }
            if (validFile)
            {
                validFile = false;
                filePath = Path.GetFullPath(filePath);

                //this is to test for invalid file types
                if (Path.GetExtension(filePath).ToLower() != ".xls")
                    validFile = true;
                if (Path.GetExtension(filePath).ToLower() != ".xlsx")
                    validFile = true;
                if (Path.GetExtension(filePath).ToLower() != ".csv")
                    validFile = true;
            }
            return validFile;
        }

        /// <summary>
        /// Enter a number between 0-25 and get the corisponding char. Over 25 will be modded down
        /// </summary>
        /// <param name="x">char value between 0 and 25</param>
        /// <returns></returns>
        private static string GetExcelColAZoZ(int x)
        {
            x %= 26;
            x += 1;
            switch (x)
            {
                default:
                case 1: return "A";
                case 2: return "B";
                case 3: return "C";
                case 4: return "D";
                case 5: return "E";
                case 6: return "F";
                case 7: return "G";
                case 8: return "H";
                case 9: return "I";
                case 10: return "J";
                case 11: return "K";
                case 12: return "L";
                case 13: return "M";
                case 14: return "N";
                case 15: return "O";
                case 16: return "P";
                case 17: return "Q";
                case 18: return "R";
                case 19: return "S";
                case 20: return "T";
                case 21: return "U";
                case 22: return "V";
                case 23: return "W";
                case 24: return "X";
                case 25: return "Y";
                case 26: return "Z";
            }
        }

        private static int GetExcelColIndex(string colName)
        {
            colName = colName.ToUpper();

            if (colName.Length > 1)
            {
                int colVal1 = GetExcelColIndex(colName[0]);
                int colVal2 = GetExcelColIndex(colName[1]);

                int result = (colVal1 * 26) + colVal2;
                return result;
            }
            else
            {
                return GetExcelColIndex(colName[0]);
            }
        }

        private static int GetExcelColIndex(char c)
        {
            switch (c)
            {
                default:
                case 'A': return 1;
                case 'B': return 2;
                case 'C': return 3;
                case 'D': return 4;
                case 'E': return 5;
                case 'F': return 6;
                case 'G': return 7;
                case 'H': return 8;
                case 'I': return 9;
                case 'J': return 10;
                case 'K': return 11;
                case 'L': return 12;
                case 'M': return 13;
                case 'N': return 14;
                case 'O': return 15;
                case 'P': return 16;
                case 'Q': return 17;
                case 'R': return 18;
                case 'S': return 19;
                case 'T': return 20;
                case 'U': return 21;
                case 'V': return 22;
                case 'W': return 23;
                case 'X': return 24;
                case 'Y': return 25;
                case 'Z': return 26;
            }
        }

        public static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion

        #region Power_RPP_Audit
        private PowerRecordComparator powerCircuitComparer = new PowerRecordComparator();

        public void CreatePower_RPP_Audit(string dataFile, string templateFile)
        {
            List<PowerRecord> dataRecords = LoadPowerData_RPP_Audit(dataFile);

            Excel.Application excelApplication = new Excel.Application();
            //excelApplication.Visible = true;//hidden for users safty, could click in or off window breaking things

            Excel.Workbook workbook = excelApplication.Workbooks.Open(templateFile, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            Excel.Worksheet templateSheet = (Excel.Worksheet)workbook.Worksheets[1];

            DateTime maxPanelDate = new DateTime(0);
            string lastPanel = null;
            object[,] data = new object[41, 10];
            Excel.Worksheet panelSheet = null;
            foreach (PowerRecord rec in dataRecords)
            {
                if (rec.panel != lastPanel)
                {
                    //new panel/sheet
                    maxPanelDate = new DateTime(0);
                    //Excel.Worksheet newSheet = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    templateSheet.Copy(templateSheet);
                    if(panelSheet!=null)
                        ReleaseObject(panelSheet);
                    panelSheet = (Excel.Worksheet)workbook.Worksheets[workbook.Worksheets.Count-1];
                    panelSheet.Name = rec.panel;

                    //site
                    //panelSheet.get_Range("C8").Value = "IRV01";
                    //panel
                    panelSheet.get_Range("C9").Value = rec.panel;

                    lastPanel = rec.panel;
                }

                //update to max date //ignore 0 value readings that may be a result of just turning the circuit off and not a new reading on the panel
                if (rec.date > maxPanelDate && rec.reading>0)
                {
                    maxPanelDate = rec.date;
                    panelSheet.get_Range("C10").Value = maxPanelDate;
                }

                string cellOfCircuit = GetDataCellForCircuit_RPP_Audit(rec.circuit);
                panelSheet.get_Range(cellOfCircuit).Value = rec.reading;

                //update percentage formula
                if (rec.amps != 20)
                {
                    string formulaCell = GetFormulaCellForCircuit_RPP_Audit(rec.circuit);
                    panelSheet.get_Range(formulaCell).Formula = "=" + cellOfCircuit+"/"+rec.amps;
                }
            }
            ReleaseObject(panelSheet);

            //delete template sheet
            templateSheet.Delete();

            //excelApplication.ActiveWindow.DisplayGridlines = false;
            excelApplication.Visible = true;

            //workbook
            ReleaseObject(workbook);
            //app
            ReleaseObject(excelApplication);
        }

        private static string GetDataCellForCircuit_RPP_Audit(int circuit)
        {
            int letterCol = (int)(((circuit-1)%6)/2);
            if (circuit % 2 == 0) letterCol += 5;//skip two in middle - put even on the right
            letterCol+=1;//offset from 0
            string letter = GetExcelColAZoZ(letterCol);
            int number = 14+(int)((circuit-1)/2);
            return letter + number;
        }

        private static string GetFormulaCellForCircuit_RPP_Audit(int circuit)
        {
            string letter;
            if (circuit % 2 == 0)
                letter = "J";
            else
                letter = "E";

            int number = 14 + (int)((circuit - 1) / 2);
            return letter + number;
        }

        private List<PowerRecord> LoadPowerData_RPP_Audit(string dataFile)
        {
            /* parse data from CSV file
             * *NOTE this creates duplicate entries for 208v circuits (circuits that continan "/" in the circuit name like "14/16")
             */
            var databasePassword = missing;

            Excel.Application excelApplication = new Excel.Application();
            //excelApplication.Visible = true;
            Excel.Workbooks dataWorkbook = excelApplication.Workbooks;
            dataWorkbook.OpenText(dataFile, missing, 3, Excel.XlTextParsingType.xlDelimited, Excel.XlTextQualifier.xlTextQualifierNone, missing, missing, missing, true, missing, missing, missing, missing, missing, missing, missing, missing, missing);

            Excel.Worksheet dataSheet = dataWorkbook[1].Worksheets.get_Item(1);

            //load data from sheet
            var powerData = new List<PowerRecord>();

            const int dataStart = 3;

            Excel.Range dataRange = dataSheet.UsedRange;
            int rowCount = dataRange.Rows.Count;
            object[,] valueArray = (object[,])dataRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

            for (int row = rowCount; row >= dataStart; row--)
            {
                try
                {
                    var circuits = valueArray[row, 2].ToString().Split('/');
                    foreach (string circuit in circuits)
                    {
                        var rec = new PowerRecord();
                        rec.panel = valueArray[row, 1].ToString();
                        rec.circuit = int.Parse(circuit);
                        rec.reading = float.Parse(valueArray[row, 3].ToString());
                        rec.amps = int.Parse(valueArray[row, 5].ToString());
                        rec.volts = int.Parse(valueArray[row, 6].ToString());
                        rec.on = valueArray[row, 7].ToString() == "On";
                        rec.date = (DateTime)valueArray[row, 8];
                        powerData.Add(rec);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Add error has occurred durring the load process. Please check the data is valid and try again.","Error");
                    Console.WriteLine("failed");
                }
            }
            powerData.Sort(powerCircuitComparer);
            
            
            //release objects
            dataWorkbook.Close();
            excelApplication.Quit();

            //sheet
            ReleaseObject(dataSheet);
            //workbook
            ReleaseObject(dataWorkbook);
            //app
            ReleaseObject(excelApplication);

            return powerData;
        }
        #endregion
    }

    #region Power Record for HDC RCC Audit
    struct PowerRecord
    {
        public string panel;
        public int circuit;
        public float reading;
        public int amps;
        public int volts;
        public bool on;
        public DateTime date;

        public override string ToString()
        {
            return "Circuit: " + panel + " " + circuit + " - " + reading + "A - " + volts + "V " + amps + "A  "+(on)+" "+date;
        }
    }

    class PowerRecordComparator : IComparer<PowerRecord>
    {
        public int Compare(PowerRecord t1, PowerRecord t2)
        {
            return (t1.panel.PadLeft(3, '0') + (t1.circuit + 100)).CompareTo(t2.panel.PadLeft(3, '0') + (t2.circuit + 100));
        }
    }
    #endregion
}
