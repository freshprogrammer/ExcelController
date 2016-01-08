using System.Windows.Forms;
using System.Reflection;
using System;

namespace Fresh
{
    class Program
    {
        enum ApplicationModes
        {
            Power_Panel_Audit_Creator,
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
                args = new string[] { "-ppa" };
            }

            //bool lastArgWasMode = false;
            for (int i=0; i < args.Length; i++)
            {
                string arg = args[i].ToLower();
                if (arg == "-ppa")
                    mode = ApplicationModes.Power_Panel_Audit_Creator;
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
                case ApplicationModes.Power_Panel_Audit_Creator:
                    if (dataFile == "")
                    {
                        OpenFileDialog file = new OpenFileDialog();
                        file.Title = "Select a Data CSV File";
                        file.Multiselect = false;
                        file.CheckFileExists = true;
                        if (file.ShowDialog() == DialogResult.OK)
                            dataFile = file.FileName;
                    }
                    if (dataFile != "" && templateFile == "")
                    {
                        OpenFileDialog file = new OpenFileDialog();
                        file.Title = "Select a template excel file";
                        file.Multiselect = false;
                        file.CheckFileExists = true;
                        if (file.ShowDialog() == DialogResult.OK)
                            templateFile = file.FileName;
                    }

                    if (!ExcelController.ValidXLSFile(ref dataFile))
                        MessageBox.Show("Cannot start without a data file.", "Missing File");
                    else if (!ExcelController.ValidXLSFile(ref templateFile))
                        MessageBox.Show("Cannot start without a template file.", "Missing File");
                    else
                        new ExcelController().CreatePower_RPP_Audit(dataFile, templateFile);
                    break;

            }
        }
    }
}
