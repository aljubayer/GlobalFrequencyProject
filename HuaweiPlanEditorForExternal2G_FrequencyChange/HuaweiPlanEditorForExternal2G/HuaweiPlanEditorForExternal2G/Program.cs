using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace HuaweiPlanEditorForExternal2G
{
    class Program
    {
        static void Main(string[] args)
        {
            string dbFileNameNeighbor = Directory.GetCurrentDirectory() + "\\2G Radio Network Planning Data Template.xls";
            string dbFileNameCell = Directory.GetCurrentDirectory() + "\\2G Cell Frequency Data Template.xls";
            string inputFile = Directory.GetCurrentDirectory() + "\\WO_Input.xlsx";
            HuaweiProcessExternalPlan aPlan = new HuaweiProcessExternalPlan();
            aPlan.ProcessPlan(dbFileNameCell,dbFileNameNeighbor, inputFile);
        }
    }
}
