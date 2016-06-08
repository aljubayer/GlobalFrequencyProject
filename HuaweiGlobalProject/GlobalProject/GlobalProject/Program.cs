using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using HuaweiPlanEditorForFrequencyChange;
using HuaweiPlanEditorForHoppingSettings;
using HuaweiPlanEditorForExternal2G;

namespace GlobalProject
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                File.Delete(Directory.GetCurrentDirectory() + "\\errorLog.txt");
                File.Delete(Directory.GetCurrentDirectory() + "\\log.txt");
                File.Delete(Directory.GetCurrentDirectory() + "\\nonquery.txt");
                File.Delete(Directory.GetCurrentDirectory() + "\\output.txt");
                File.Delete(Directory.GetCurrentDirectory() + "\\Validation.txt");





                string dbFileName = Directory.GetCurrentDirectory() + "\\2G Cell Frequency Data Template.xls";
                string inputFile = Directory.GetCurrentDirectory() + "\\WO_Input.xlsx";
                HuaweiPlanProcess aPlanProcess = new HuaweiPlanProcess();
                string message = aPlanProcess.ProcessPlan(dbFileName, inputFile);
                Console.WriteLine(message);
                Console.WriteLine("");
                Console.WriteLine("");
                aPlanProcess.WriteLogFiles();
                if (message.Trim().Length > 0)
                {
                    Console.ReadKey();
                    return;
                }

                dbFileName = Directory.GetCurrentDirectory() + "\\2G Cell Frequency Data Template.xls";
                inputFile = Directory.GetCurrentDirectory() + "\\WO_Input.xlsx";
                HuaweiPlanEditorForHoppingModeSettings aPlanProcessHop = new HuaweiPlanEditorForHoppingModeSettings();
                message = aPlanProcessHop.ProcessWOInputFile(dbFileName, inputFile);
                Console.WriteLine(message);
                Console.WriteLine("");
                Console.WriteLine("");
                aPlanProcessHop.WriteLogFiles();


                string dbFileNameNeighbor = Directory.GetCurrentDirectory() + "\\2G Radio Network Planning Data Template.xls";
                string dbFileNameCell = Directory.GetCurrentDirectory() + "\\2G Cell Frequency Data Template.xls";
                inputFile = Directory.GetCurrentDirectory() + "\\WO_Input.xlsx";
                HuaweiProcessExternalPlan aPlan = new HuaweiProcessExternalPlan();
                aPlan.ProcessPlan(dbFileNameCell, dbFileNameNeighbor, inputFile);
                aPlan.WriteValidationFile();
                aPlan.WriteLogFiles();

                Console.WriteLine("");
                Console.WriteLine("");
               
                Console.WriteLine("Complete...Press any key to exit");

                Console.ReadKey();
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception Occured: " + exception.Message);
                Console.ReadKey();
            }
        }
    }
}
