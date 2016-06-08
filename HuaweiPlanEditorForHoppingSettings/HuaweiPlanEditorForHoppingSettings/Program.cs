using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using HuaweiPlanEditorForHoppingSettings;


namespace HuaweiPlanEditorForFrequencyChange
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string dbFileName = Directory.GetCurrentDirectory() + "\\2G Cell Frequency Data Template.xls";
                string inputFile = Directory.GetCurrentDirectory() + "\\WO_Input.xlsx";
                HuaweiPlanEditorForHoppingModeSettings aPlanProcess = new HuaweiPlanEditorForHoppingModeSettings();
                string message = aPlanProcess.ProcessWOInputFile(dbFileName, inputFile);
                Console.WriteLine(message);
                Console.WriteLine("");
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
