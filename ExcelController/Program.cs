using System.Windows.Forms;
using System.Reflection;
using System;

namespace Fresh
{
    class Program
    {
        enum ApplicationModes
        {
            Power_HDC_RCC_Audit,
            Test,
            None,
        }

        [STAThread]
        static void Main(string[] args)
        {
            ApplicationModes mode = ApplicationModes.None;
            string templateFile = "";
            string dataFile = "";

            if (args.Length == 0)
            {
                //HACK test command lines
                args = new string[] { "-pa" };
            }

            //bool lastArgWasMode = false;
            for (int i=0; i < args.Length; i++)
            {
                string arg = args[i].ToLower();
                if (arg == "-pa")
                    mode = ApplicationModes.Power_HDC_RCC_Audit;
                else if (arg == "-?")
                    mode = ApplicationModes.Test;
            }

            if (mode == ApplicationModes.None && false)
            {
                //HACK test modes
                //mode = ApplicationModes.ImportGas;
            }

            switch (mode)
            {
                case ApplicationModes.None:
                    MessageBox.Show("Excel Controller \nVersion: " + Assembly.GetExecutingAssembly().GetName().Version + "\n No mode specified. Program will now close.", "Excel Contrller", MessageBoxButtons.OK);
                    break;
                case ApplicationModes.Power_HDC_RCC_Audit:
                    //dataFile = @"C:\Fresh Temp\Excel data\PowerHistory-2015-04-10.csv";
                    //templateFile = @"C:\Fresh Temp\Excel data\HDC RPP Audit - All COLO - Template.xlsx";

                    if (dataFile == "")
                    {
                        OpenFileDialog file = new OpenFileDialog();
                        file.Title = "Select a Data CSV File";
                        if (file.ShowDialog() == DialogResult.OK)
                        {
                            dataFile = file.FileName;
                        }
                    }
                    if (templateFile == "")
                    {
                        OpenFileDialog file = new OpenFileDialog();
                        file.Title = "Select a template excel file";
                        if (file.ShowDialog() == DialogResult.OK)
                        {
                            templateFile = file.FileName;
                        }
                    }

                    if (!ExcelController.ValidXLSFile(ref dataFile))
                        MessageBox.Show("Cannot start without a data file.");
                    else if (!ExcelController.ValidXLSFile(ref templateFile))
                        MessageBox.Show("Cannot start without a summary template.");
                    else
                        new ExcelController().CreatePower_HDC_RCC_Audit(dataFile, templateFile);
                    break;

            }
        }
    }
}
