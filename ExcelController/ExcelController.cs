using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Collections.Generic;
using System;

namespace Fresh
{
    class ExcelController
    {
        private PowerRecordComparator powerCircuitComparer = new PowerRecordComparator();

        private object missing = System.Reflection.Missing.Value;

        public ExcelController()
        {

        }

        public void CreatePower_HDC_RCC_Audit(string dataFile, string template)
        {
            List<PowerRecord> data = LoadPowerData(dataFile);

        }

        public List<PowerRecord> LoadPowerData(string dataFile)
        {
            /* parse data from CSV file
             */
            var databasePassword = missing;

            Excel.Application excelApplication = new Excel.Application();
            //excelApplication.Visible = true;
            Excel.Workbooks dataWorkbook = excelApplication.Workbooks;
            dataWorkbook.OpenText(dataFile, missing, 3, Excel.XlTextParsingType.xlDelimited, Excel.XlTextQualifier.xlTextQualifierNone, missing, missing, missing, true, missing, missing, missing, missing, missing, missing, missing, missing, missing);

            Excel.Worksheet dataSheet = dataWorkbook[1].Worksheets.get_Item(1);

            //load data from sheet
            var powerData = new List<PowerRecord>();

            const int dataStart = 2;

            Excel.Range dataRange = dataSheet.UsedRange;
            int rowCount = dataRange.Rows.Count;
            object[,] valueArray = (object[,])dataRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

            for (int row = rowCount; row >= dataStart; row--)
            {
                try
                {
                    var rec = new PowerRecord();
                    rec.panel =      valueArray[row, 1].ToString();
                    rec.circuit = int.Parse(valueArray[row, 2].ToString());
                    rec.reading = float.Parse(valueArray[row, 3].ToString());
                    rec.amps = int.Parse(valueArray[row, 5].ToString());
                    rec.volts = int.Parse(valueArray[row, 6].ToString());
                    rec.on = valueArray[row, 7].ToString() == "On";
                    rec.date = valueArray[row, 8].ToString();
                    powerData.Add(rec);
                }
                catch (Exception)
                {
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
    }

    class PowerRecordComparator : IComparer<PowerRecord>
    {
        public int Compare(PowerRecord t1, PowerRecord t2)
        {
            return (t1.panel+(t1.circuit+100)).CompareTo(t2.panel+(t2.circuit+100));
        }
    }

    struct PowerRecord
    {
        public string panel;
        public int circuit;
        public float reading;
        public int amps;
        public int volts;
        public bool on;
        public string date;

        public override string ToString()
        {
            return "Circuit: " + panel + " " + circuit + " - " + reading + "A - " + volts + "V " + amps + "A  "+(on)+" "+date;
        }
    }
}
