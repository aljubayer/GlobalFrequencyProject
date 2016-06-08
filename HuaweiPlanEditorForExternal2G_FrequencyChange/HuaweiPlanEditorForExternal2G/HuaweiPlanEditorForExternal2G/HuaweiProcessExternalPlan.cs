using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using ManiacProject.Libs;
using Microsoft.Office.Interop.Excel;

namespace HuaweiPlanEditorForExternal2G
{
    public class HuaweiProcessExternalPlan
    {
        private string dbFileNameNeighbor = string.Empty;
        private string dbFileNameGCELL = string.Empty;
        private string inputFile = string.Empty;
        private List<FrequencyChangeData> woInputFrequency = new List<FrequencyChangeData>();
        private List<BSICChangeData> woInputBSIC = new List<BSICChangeData>();
        private List<Dictionary<string,string>> woInputNeighborDelete = new List<Dictionary<string, string>>();
        
        List<string> nonQueryCommandList = new List<string>();
        private Dictionary<string, string> logDictionary = new Dictionary<string, string>();
        private Dictionary<string, string> errorLogDictionary = new Dictionary<string, string>();
        Dictionary<string,Dictionary<string,string>> gcellByLacCi = new Dictionary<string, Dictionary<string, string>>();
        Dictionary<string, Dictionary<string, string>> gcellByCellName = new Dictionary<string, Dictionary<string, string>>();
        Dictionary<string, List<Dictionary<string, string>>> g2GnCell = new Dictionary<string, List<Dictionary<string, string>>>();
        List<string> neighborMismatch = new List<string>();

        public void ProcessPlan(string dbFileNameCell,string dbFileNameNeighbor, string inputFile)
        {
            this.dbFileNameGCELL = dbFileNameCell;
            this.dbFileNameNeighbor = dbFileNameNeighbor;
            this.inputFile = inputFile;
            Console.WriteLine("Loading GCELL Data.....");
            LoadGcellData();
            Console.WriteLine("Reading WO Input.....");
            ReadWOInput(inputFile);
            LoadG2GNCell();
            Console.WriteLine("Generating Non-Query .....");
            GenerateExternalUpdateQuery();
            GenerateCommandForDeleteG2GNCELL();
            Console.WriteLine("Running Non-Query .....");
            RunNonQueryCommandsInDBFile();
            Console.WriteLine("Validating G2GNCELL.....");
            ValidateG2GNCELL();

        }

        private void GenerateCommandForDeleteG2GNCELL()
        {
            List<string> g2gnCols = new List<string>();
            if (g2GnCell.Count > 0)
            {
                Dictionary<string,string> aDictionary = g2GnCell.First().Value[0];
                foreach (KeyValuePair<string, string> keyValuePair in aDictionary)
                {
                    g2gnCols.Add(keyValuePair.Key);
                }
            }

            foreach (Dictionary<string, string> dictionary in woInputNeighborDelete)
            {
                string temp = "set ";
                foreach (string g2GnCol in g2gnCols)
                {
                    temp += g2GnCol + "='',";
                }
                temp = temp.Trim().Substring(0, temp.LastIndexOf(','));
                string nonQuery = "update [G2GNCELL$] " + temp + " where SRCCELLNAME='" + dictionary["SourceCell"] + "' and NBRCELLNAME='" + dictionary["NeighborCell"] + "';";
                nonQueryCommandList.Add(nonQuery);
            
            }
        }


        public void WriteValidationFile()
        {
            using (StreamWriter sw = new StreamWriter("Validation.txt"))
            {
                foreach (string mismatch in neighborMismatch)
                {
                    sw.WriteLine(mismatch);
                }

                sw.Close();
            }
        }

       

        private void LoadG2GNCell()
        {
            DataSet aSet = IOFileOperation.ReadExcelFile(dbFileNameNeighbor, "G2GNCELL");

            List<string> cols = new List<string>();
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }
            
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                Dictionary<string,string> aDictionary = new Dictionary<string, string>();
                foreach (string col in cols)
                {
                    aDictionary.Add(col,dataRow[col].ToString());
                }

                string index = "BSC-" + dataRow["BSCName"].ToString() + "_CELLNAME-" + dataRow["SRCCELLNAME"].ToString();
                if (g2GnCell.ContainsKey(index))
                {
                    g2GnCell[index].Add(aDictionary);
                }
                else
                {
                    List<Dictionary<string,string>> aList = new List<Dictionary<string, string>>();
                    aList.Add(aDictionary);
                    g2GnCell.Add(index,aList);
                }
            }

        }

        private void LoadGcellData()
        {
            DataSet aSet = IOFileOperation.ReadExcelFile(dbFileNameGCELL, "GCELL");

            List<string> cols = new List<string>();
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }
           
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                string index = "LAC-" + dataRow["LAC"].ToString() + "_CI-" + dataRow["CI"].ToString();

                Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                foreach (string col in cols)
                {
                    aDictionary.Add(col, dataRow[col].ToString());
                }

                if (!gcellByLacCi.ContainsKey(index))
                {
                    gcellByLacCi.Add(index, aDictionary);
                }


                index = "CellName-" + dataRow["CELLNAME"].ToString();

                if (!gcellByCellName.ContainsKey(index))
                {
                    gcellByCellName.Add(index, aDictionary);
                }
            }

            Dictionary<string,Dictionary<string,string>> gtrx = new Dictionary<string, Dictionary<string, string>>();
                
            aSet = IOFileOperation.ReadExcelFile(dbFileNameGCELL, "GTRX");
            cols = new List<string>();
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                string index = "BSC-" + dataRow["BSCName"].ToString() + "_CELLID-" + dataRow["CELLID"].ToString();

                if (dataRow["ISMAINBCCH"].ToString() == "YES")
                {
                    Dictionary<string,string> aDictionary = new Dictionary<string, string>();
                    foreach (string col in cols)
                    {
                        aDictionary.Add(col,dataRow[col].ToString());
                    }
                    if (!gtrx.ContainsKey(index))
                    {
                        gtrx.Add(index, aDictionary);
                    }
                }
            }

            foreach (KeyValuePair<string, Dictionary<string, string>> keyValuePair in gcellByLacCi)
            {
                Dictionary<string, string> aDictionary = (Dictionary<string, string>) keyValuePair.Value;
                string gtrxIndex = "BSC-" + aDictionary["BSCName"].ToString() + "_CELLID-" + aDictionary["CELLID"].ToString();

                if (gtrx.ContainsKey(gtrxIndex))
                {
                    aDictionary.Add("BCCH", gtrx[gtrxIndex]["FREQ"]);
                }
                else
                {
                    aDictionary.Add("BCCH","");
                }
            }
        }

        public void ResetFields()
        {
            woInputFrequency = new List<FrequencyChangeData>();
            woInputBSIC = new List<BSICChangeData>();
            gcellByLacCi = new Dictionary<string, Dictionary<string, string>>();
            gcellByCellName = new Dictionary<string, Dictionary<string, string>>();
            g2GnCell = new Dictionary<string, List<Dictionary<string, string>>>();
            neighborMismatch = new List<string>();
            LoadGcellData();
            ReadWOInput(inputFile);
            LoadG2GNCell();
        }

        private bool ValidateG2GNCELL()
        {
            ResetFields();


            Dictionary<string,Dictionary<string,string>> woInputDictionary = new Dictionary<string, Dictionary<string, string>>();
            foreach (FrequencyChangeData frequencyChangeData in woInputFrequency)
            {
                Dictionary<string,string> aDictionary = new Dictionary<string, string>();
                aDictionary.Add("BSC",frequencyChangeData.BSCName);
                aDictionary.Add("CELLID", frequencyChangeData.CELLID);
                aDictionary.Add("LAC",frequencyChangeData.LAC);
                aDictionary.Add("CI", frequencyChangeData.CI);

                aDictionary.Add("BCCH",frequencyChangeData.NewFrequency);
                var bsicVar = 
                    woInputBSIC.Where(
                        i => i.BSCName == frequencyChangeData.BSCName && i.CELLID == frequencyChangeData.CELLID);


                
                string ncc = string.Empty;
                string bcc = string.Empty;
                foreach (BSICChangeData bsicChangeData in bsicVar)
                {
                    ncc = GetNCC(bsicChangeData.NewBSIC);
                    bcc = GetBCC(bsicChangeData.NewBSIC);
                }
              
                
                aDictionary.Add("NCC", ncc);
                aDictionary.Add("BCC",bcc);
                aDictionary.Add("CELLNAME",frequencyChangeData.CELLNAME);

                if (!woInputDictionary.ContainsKey(aDictionary["CELLNAME"]))
                {
                    woInputDictionary.Add(aDictionary["CELLNAME"],aDictionary);
                }
                
            }
            
            // BCCH,NCC,BCC not changed in wo, then load system value;
            foreach (KeyValuePair<string, Dictionary<string, string>> keyValuePair in woInputDictionary)
            {
                Dictionary<string, string> aDictionary = (Dictionary<string, string>) keyValuePair.Value;

                if (aDictionary["BCCH"].Trim().Length == 0)
                {
                    string index = "LAC-" + aDictionary["LAC"] + "_CI-" + aDictionary["CI"];
                    if (gcellByLacCi.ContainsKey(index))
                    {
                        aDictionary["BCCH"] = gcellByLacCi[index]["BCCH"];
                    }
                }
                if (aDictionary["NCC"].Trim().Length == 0)
                {
                    string index = "LAC-" + aDictionary["LAC"] + "_CI-" + aDictionary["CI"];
                    if (gcellByLacCi.ContainsKey(index))
                    {
                        aDictionary["NCC"] = gcellByLacCi[index]["NCC"];
                    }

                }
                if (aDictionary["BCC"].Trim().Length == 0)
                {
                    string index = "LAC-" + aDictionary["LAC"] + "_CI-" + aDictionary["CI"];
                    if (gcellByLacCi.ContainsKey(index))
                    {
                        aDictionary["BCC"] = gcellByLacCi[index]["BCC"];
                    }
                }
                string mismatch = GetNeighborMismatch(aDictionary);
                if (mismatch.Trim().Length != 0)
                {
                    neighborMismatch.Add(mismatch);
                }
            }

            Dictionary<string, Dictionary<string, string>> sourceCellsForNeighborCellsOfWOInput =
                GetSourceCellsForNeighborCellsOfWOInput(woInputDictionary);



            foreach (KeyValuePair<string, Dictionary<string, string>> keyValuePair in sourceCellsForNeighborCellsOfWOInput)
            {
                Dictionary<string, string> aDictionary = (Dictionary<string, string>)keyValuePair.Value;
                string mismatch = GetNeighborMismatch(aDictionary);
                if (mismatch.Trim().Length != 0)
                {
                    neighborMismatch.Add(mismatch);
                }
            }



            if (neighborMismatch.Count > 0)
            {
                List<string> formattedMismatch = new List<string>();
                foreach (string aMismatch in neighborMismatch)
                {
                    string[] arrayString = Regex.Split(aMismatch, "\r\n");
                    foreach (string aString in arrayString)
                    {
                        if (!formattedMismatch.Contains(aString.Trim()))
                        {
                            formattedMismatch.Add(aString);
                        }
                    }
                }

                neighborMismatch = formattedMismatch;
                return false;
            }

            return true;
        }

        private Dictionary<string, Dictionary<string, string>> GetSourceCellsForNeighborCellsOfWOInput(Dictionary<string, Dictionary<string, string>> woInputDictionary)
        {
            
            Dictionary<string, Dictionary<string, string>> sourceCells =
                new Dictionary<string, Dictionary<string, string>>();
            List<Dictionary<string, string>> listNeighborCells = new List<Dictionary<string, string>>();

            foreach (KeyValuePair<string, Dictionary<string, string>> keyValuePair in woInputDictionary)
            {
              
                foreach (KeyValuePair<string, List<Dictionary<string, string>>> valuePair in g2GnCell)
                {
                    foreach (Dictionary<string, string> dictionary in valuePair.Value)
                    {
                        if (dictionary["NBRCELLNAME"] == keyValuePair.Key
                            && dictionary["SRCCELLNAME"] != keyValuePair.Key
                            && dictionary["ISNCELL"] == "INNCELL")
                        {
                            listNeighborCells.Add(dictionary);
                        }

                    }
                }
            }




            foreach (Dictionary<string, string> aNeigbor in listNeighborCells)
            {
                if (!sourceCells.ContainsKey(aNeigbor["SRCCELLNAME"]))
                {
                    string gcellIndex = "CellName-" + aNeigbor["SRCCELLNAME"];
                    if (gcellByCellName.ContainsKey(gcellIndex))
                    {
                   
                        Dictionary<string,string> aDictionary = new Dictionary<string, string>();
                        aDictionary.Add("BSC", gcellByCellName[gcellIndex]["BSCName"]);
                        aDictionary.Add("LAC", gcellByCellName[gcellIndex]["LAC"]);
                        aDictionary.Add("CI", gcellByCellName[gcellIndex]["CI"]);
                        aDictionary.Add("CELLNAME", gcellByCellName[gcellIndex]["CELLNAME"]);
                        aDictionary.Add("CELLID", gcellByCellName[gcellIndex]["CELLID"]);
                        aDictionary.Add("BCCH", gcellByCellName[gcellIndex]["BCCH"]);
                        aDictionary.Add("NCC", gcellByCellName[gcellIndex]["NCC"]);
                        aDictionary.Add("BCC", gcellByCellName[gcellIndex]["BCC"]);
                        sourceCells.Add(aNeigbor["SRCCELLNAME"],aDictionary);

                    }
                }
            }

         
            return sourceCells;
        }

        
        private string GetNeighborMismatch(Dictionary<string, string> neighborSouceCell)
        {
            string output = string.Empty;
            string sourceCellName = neighborSouceCell["CELLNAME"];
            string bsc = neighborSouceCell["BSC"];

            if (g2GnCell.ContainsKey("BSC-" + bsc + "_CELLNAME-"+sourceCellName))
            {
              
                foreach (Dictionary<string, string> dictionary in g2GnCell["BSC-" + bsc + "_CELLNAME-" + sourceCellName])
                {
                    string neighborIndex = "CellName-" + dictionary["NBRCELLNAME"];
                    if (gcellByCellName.ContainsKey(neighborIndex) 
                        && dictionary["SRCCELLNAME"] != dictionary["NBRCELLNAME"])
                    {
                        if (gcellByCellName[neighborIndex]["BCCH"] == neighborSouceCell["BCCH"]
                            && gcellByCellName[neighborIndex]["NCC"] == neighborSouceCell["NCC"]
                            && gcellByCellName[neighborIndex]["BCC"] == neighborSouceCell["BCC"])
                        {
                            output += "Mismatch: Source Cell: " + sourceCellName
                                + ", Neighbor Cell: " + dictionary["NBRCELLNAME"] + ", BCCH: "
                                + neighborSouceCell["BCCH"] + ", NCC: " + neighborSouceCell["NCC"]
                                + ", BCC: " + neighborSouceCell["BCC"] + ";BCCH-BSIC Conflict\r\n"; 
                        }
                    }

                    output += GetMismatchOnInsideNeighboringCells(g2GnCell["BSC-" + bsc + "_CELLNAME-" + sourceCellName], dictionary);
                    
                }

            }
            return output;
        }

        private string GetMismatchOnInsideNeighboringCells(List<Dictionary<string, string>> neighborCells,
            Dictionary<string, string> referenceCell)
        {

            string referenceBCCH = string.Empty;
            string referenceNCC = string.Empty;
            string referenceBCC = string.Empty;
            string gcellIndex = "CellName-" + referenceCell["NBRCELLNAME"];
            if (gcellByCellName.ContainsKey(gcellIndex))
            {
                referenceBCCH = gcellByCellName[gcellIndex]["BCCH"];
                referenceNCC = gcellByCellName[gcellIndex]["NCC"];
                referenceBCC = gcellByCellName[gcellIndex]["BCC"];
            }
            string output = string.Empty;

            foreach (Dictionary<string, string> neighborCell in neighborCells)
            {
                string neighborBCCH = string.Empty;
                string neighborNCC = string.Empty;
                string neighborBCC = string.Empty;

                gcellIndex = "CellName-" + neighborCell["NBRCELLNAME"];
                if (gcellByCellName.ContainsKey(gcellIndex))
                {
                    neighborBCCH = gcellByCellName[gcellIndex]["BCCH"];
                    neighborNCC = gcellByCellName[gcellIndex]["NCC"];
                    neighborBCC = gcellByCellName[gcellIndex]["BCC"];
                }

                if (referenceBCCH == neighborBCCH && referenceNCC == neighborNCC 
                    && referenceBCC == neighborBCC && referenceCell["NBRCELLNAME"] !=neighborCell["NBRCELLNAME"]
                    && neighborBCCH.Trim().Length != 0)
                {
                    output += "Mismatch Beetween 2 Neighbors:Source Cell: " + referenceCell["SRCCELLNAME"] + ", Neighbor Cell 1: " + referenceCell["NBRCELLNAME"]
                                + ", Neighbor Cell 2: " + neighborCell["NBRCELLNAME"] + ", BCCH: "
                                + neighborBCCH + ", NCC: " + neighborNCC
                                + ", BCC: " + neighborBCC + ";BCCH-BSIC Conflict\r\n";
 
                }

            }



            return output;
        }
        public void WriteLogFiles()
        {
            StreamWriter aWriterLog = new StreamWriter("log.txt",true);
            StreamWriter aWriterErrorLog = new StreamWriter("errorLog.txt",true);

            foreach (KeyValuePair<string, string> keyValuePair in errorLogDictionary)
            {
                aWriterErrorLog.WriteLine(keyValuePair.Value);
            }

            foreach (KeyValuePair<string, string> keyValuePair in logDictionary)
            {
                aWriterLog.WriteLine(keyValuePair.Value);
            }
            aWriterErrorLog.Close();
            aWriterLog.Close();

            StreamWriter aqStreamWriter = new StreamWriter("nonquery.txt",true);
            foreach (string s in nonQueryCommandList)
            {
                aqStreamWriter.WriteLine(s);
            }
            aqStreamWriter.Close();


          

        }

        private void RunNonQueryCommandsInDBFile()
        {
            int index = 0;
            if (logDictionary.Keys.Count != 0)
            {
                index = Convert.ToInt16(logDictionary.Keys.Last()) + 1;
            }

            int totalCommand = nonQueryCommandList.Where(i => i != "").Count();
            string log = string.Empty;
            string errorLog = string.Empty;

            ExecuteNonQueryOnExcel aNonQueryOnExcel = new ExecuteNonQueryOnExcel(dbFileNameNeighbor);


            foreach (string nonQuery in nonQueryCommandList)
            {
                if (nonQuery.Trim().Length != 0)
                {
                    log = "Running Command: " + nonQuery + "\r\n";
                    Console.WriteLine("Running Command(" + ++index + "/" + totalCommand + "): " + nonQuery);
                    int affectedRows = aNonQueryOnExcel.ExecuteCommandOnExcelFile(nonQuery);

                    Console.WriteLine("Affected Rows: " + affectedRows);
                    log += "Affected Rows: " + affectedRows + "\n";
                    logDictionary.Add(index.ToString(), log);

                    if (affectedRows != 1)
                    {
                        errorLog = "Error: " + nonQuery + "\r\n Affected Rows: " + affectedRows;
                        errorLogDictionary.Add(index.ToString(), errorLog);

                    }

                }
            }
            aNonQueryOnExcel.CloseConnection();

        }


        private void GenerateExternalUpdateQuery()
        {
            foreach (FrequencyChangeData frequencyChangeData in woInputFrequency)
            {
                string nonQuery = "update [GEXT2GCELL$] set BCCH = '" + frequencyChangeData.NewFrequency + "' where LAC='" + frequencyChangeData.LAC + "' and CI='"+frequencyChangeData.CI+"';";
                nonQueryCommandList.Add(nonQuery);
            }

            foreach (BSICChangeData bsicChangeData in woInputBSIC)
            {
                string ncc = GetNCC(bsicChangeData.NewBSIC);
                string bcc = GetBCC(bsicChangeData.NewBSIC);
                string nonQuery = "update [GEXT2GCELL$] set NCC = '" + ncc + "', BCC='" + bcc + "' where LAC='" + bsicChangeData.LAC + "' and CI='" + bsicChangeData.CI + "';";
                nonQueryCommandList.Add(nonQuery);
            }
        }


        private void ReadWOInput(string inputFile)
        {
            


            DataSet aSet = IOFileOperation.ReadExcelFile(inputFile, "Input");
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
               
                if (dataRow["NEW BCCH"].ToString().Trim().Length != 0)
                {
                    FrequencyChangeData aData = new FrequencyChangeData();
                    aData.LAC = dataRow["LAC"].ToString().Trim();
                    aData.CI = dataRow["CI"].ToString().Trim();
                    aData.OldFrequency = dataRow["BCCH"].ToString().Trim();
                    aData.NewFrequency = dataRow["NEW BCCH"].ToString().Trim();
                    string index = "LAC-" + aData.LAC + "_CI-" + aData.CI;
                    if (gcellByLacCi.ContainsKey(index))
                    {
                        aData.BSCName = gcellByLacCi[index]["BSCName"];
                        aData.CELLNAME = gcellByLacCi[index]["CELLNAME"];
                        aData.CELLID = gcellByLacCi[index]["CELLID"];

                    }
                    woInputFrequency.Add(aData);
                }
                string bsic = dataRow["NEW BSIC"].ToString().Trim();

                if (dataRow["NEW BSIC"].ToString().Trim().Length != 0)
                {
                    BSICChangeData aData = new BSICChangeData();
                    aData.LAC = dataRow["LAC"].ToString().Trim();
                    aData.CI = dataRow["CI"].ToString().Trim();
                    aData.OldBSIC = dataRow["BSIC"].ToString().Trim();
                    aData.NewBSIC = dataRow["NEW BSIC"].ToString().Trim();
                    string index = "LAC-" + aData.LAC + "_CI-" + aData.CI;
                    if (gcellByLacCi.ContainsKey(index))
                    {
                        aData.BSCName = gcellByLacCi[index]["BSCName"];
                        aData.CELLNAME = gcellByLacCi[index]["CELLNAME"];
                        aData.CELLID = gcellByLacCi[index]["CELLID"];
                    }

                    woInputBSIC.Add(aData);
                }
            }




            aSet = IOFileOperation.ReadExcelFile(inputFile, "NeighborDelete");
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                if (dataRow["Cell ID"].ToString().Trim().Length != 0)
                {
                    Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                    aDictionary.Add("SourceCell", dataRow["Cell ID"].ToString().Trim());
                    aDictionary.Add("LAC", dataRow["Cell ID"].ToString().Trim());
                    aDictionary.Add("CI", dataRow["Cell ID"].ToString().Trim());
                    aDictionary.Add("NeighborCell", dataRow["Neighbour Cell ID"].ToString().Trim());
                    aDictionary.Add("NeighborLAC", dataRow["NLAC"].ToString().Trim());
                    aDictionary.Add("NeigborCI", dataRow["NCI"].ToString().Trim());
                    woInputNeighborDelete.Add(aDictionary);
                }
            }
        }


        private string GetBCC(string bsic)
        {


            if (bsic.Contains("-"))
            {
                return bsic.Trim().Split('-')[1];
            }


            if (bsic.Trim().Length == 1)
            {
                return bsic;
            }
            return bsic.Trim().ToCharArray()[1].ToString();

            //return bcc;
        }

        private string GetNCC(string bsic)
        {


            if (bsic.Contains("-"))
            {
                return bsic.Trim().Split('-')[0];
            }


            if (bsic.Trim().Length == 1)
            {
                return "0";
            }
            return bsic.Trim().ToCharArray()[0].ToString();

            //return ncc;
        }

    }


    public class FrequencyChangeData
    {

        public string BSCName { set; get; }
        public string CELLID { set; get; }
        public string LAC { set; get; }
        public string CI { set; get; }
        public string TRXId { set; get; }
        public string OldFreqColumnNameInGCELLMAGRP { set; get; }
        public string OldFrequency { set; get; }
        public string NewFrequency { set; get; }


        public string CELLNAME { get; set; }
    }
    public class BSICChangeData
    {
        public string BSCName { set; get; }
        public string CELLID { set; get; }
        public string LAC { set; get; }
        public string CI { set; get; }
        public string OldBSIC { set; get; }
        public string NewBSIC { set; get; }

        public string CELLNAME { get; set; }
    }
    public class ExecuteNonQueryOnExcel
    {
        System.Data.OleDb.OleDbConnection MyConnection;
        System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();

        public ExecuteNonQueryOnExcel(string dbFileName)
        {
            MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + dbFileName + "';Extended Properties=Excel 8.0;");
            MyConnection.Open();
        }
        public int ExecuteCommandOnExcelFile(string nonQuery)
        {
            myCommand.Connection = MyConnection;
            myCommand.CommandText = nonQuery;
            int affectedRows = myCommand.ExecuteNonQuery();
            return affectedRows;

        }

        public static DataSet ReadFromExcelFile(string query, string file)
        {
            OleDbConnection con =
                new System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + file + "';Extended Properties=Excel 8.0;");
            OleDbDataAdapter da = new OleDbDataAdapter(query, con);
            DataSet aDataObjectSet = new DataSet();
            da.Fill(aDataObjectSet);
            con.Close();
            return aDataObjectSet;
        }

        public void CloseConnection()
        {
            MyConnection.Close();
        }
        public static List<string> RemoveTRXIDFromGCELLFREQ_FREQColumn(string dbFile, string table)
        {
            List<string> xlData = new List<string>();
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;


            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(dbFile, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(table);


            Console.WriteLine(xlWorkSheet.Name);


            xlWorkSheet.Cells[1, 5] = "";
            //xlWorkBook.
            xlWorkBook.SaveAs(dbFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            xlWorkBook.Close(false, false, false);
            xlApp.Quit();

            return xlData;


        }

    }
}
