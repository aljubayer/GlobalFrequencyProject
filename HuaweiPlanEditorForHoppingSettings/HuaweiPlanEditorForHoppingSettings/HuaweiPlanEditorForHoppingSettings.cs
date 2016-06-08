using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using ManiacProject.Libs;
using Microsoft.Office.Interop.Excel;

namespace HuaweiPlanEditorForHoppingSettings
{
    public class HuaweiPlanEditorForHoppingModeSettings
    {

        List<Dictionary<string, string>> DeleteGCELLMAGRP = new List<Dictionary<string, string>>();
        List<Dictionary<string, string>> CreateGCELLMAGRP = new List<Dictionary<string, string>>();
        List<Dictionary<string, string>> GTRXHOP = new List<Dictionary<string, string>>();
        List<Dictionary<string, string>> GTRXCHANHOP = new List<Dictionary<string, string>>();
        Dictionary<string, string> logDictionary = new Dictionary<string, string>();
        Dictionary<string, string> errorLogDictionary = new Dictionary<string, string>();
        List<string> allNonQuery = new List<string>();
        private string dbFile = string.Empty;

        List<string> nonQueryCommands = new List<string>(); 

        public string ProcessWOInputFile(string dbFile, string inputFile)
        {
            this.dbFile = dbFile;
            string fileName = string.Empty;
            Console.WriteLine("Loading Input file.....");
            LoadInputFile(inputFile);

            Console.WriteLine("Generating Command for DeleteGCELLMAGRP.....");
            GenerateNonQueryCommandsForDeleteGCELLMAGRP(dbFile);
            Console.WriteLine("Generating Command for CreateGCELLMAGRP.....");
            GenerateNonQueryCommandsForCreateGCELLMAGRP(dbFile);
            Console.WriteLine("Generating Command for GTRXHOP.....");
            GenerateNonQueryCommandsForGTRXHOP();
            Console.WriteLine("Generating Command for GTRXCHANHOP.....");
            GenerateNonQueryCommandsForGTRXCHANHOP();
            Console.WriteLine("Run Non Query Commands.....");
            RunNonQueryCommandsInDBFile(dbFile);
            allNonQuery.AddRange(nonQueryCommands);

            Console.WriteLine("Processing GCELLFREQ_FREQ......");
            ProcessDataInGCELLFREQ();
            allNonQuery.AddRange(nonQueryCommands);

            WriteLogFiles();
            return fileName;
        }

        List<Dictionary<string, string>> GCELL = new List<Dictionary<string, string>>();
        List<Dictionary<string, string>> GCELLFREQ = new List<Dictionary<string, string>>();
        List<Dictionary<string, string>> GTRX = new List<Dictionary<string, string>>();
        List<Dictionary<string, string>> GCELLMAGRP = new List<Dictionary<string, string>>();
        private void ProcessDataInGCELLFREQ()
        {
            nonQueryCommands = new List<string>();
            ResetTemplateDataFields();
            LoadTemplateData(dbFile);
            RemoveAllDataFromGCELLFREQ_FREQ();
            LoadGCELLFREQData();
            RunNonQueryCommandsInDBFile(dbFile);
        }




        private void LoadGCELLFREQData()
        {
            List<Dictionary<string, string>> gcellFreqNewData = new List<Dictionary<string, string>>();
            foreach (Dictionary<string, string> dictionary in GCELL)
            {
                if (dictionary["BSCName"] != "BSC Name")
                {
                    string bscName = dictionary["BSCName"];
                    string cellId = dictionary["CELLID"];
                    string cellName = dictionary["CELLNAME"];

                    List<string> freqs = new List<string>();
                    List<Dictionary<string, string>> freqDictionary = new List<Dictionary<string, string>>();
                    freqDictionary = GTRX.Where(i => i["BSCName"] == bscName && i["CELLID"] == cellId).ToList();
                    foreach (Dictionary<string, string> aFreqDictionary in freqDictionary)
                    {
                        freqs.Add(aFreqDictionary["FREQ"]);
                    }

                    freqDictionary = new List<Dictionary<string, string>>();
                    freqDictionary = GCELLMAGRP.Where(i => i["BSCName"] == bscName && i["CELLID"] == cellId).ToList();

                    foreach (Dictionary<string, string> aFreqDictionary in freqDictionary)
                    {
                        foreach (KeyValuePair<string, string> keyValuePair in aFreqDictionary)
                        {
                            if (keyValuePair.Key.Contains("FREQ") && keyValuePair.Value.Trim().Length > 0)
                            {
                                if (!freqs.Contains(keyValuePair.Value))
                                {
                                    freqs.Add(aFreqDictionary[keyValuePair.Key]);
                                }
                            }
                        }

                    }

                    foreach (string freq in freqs)
                    {
                        Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                        aDictionary.Add("BSCName", bscName);
                        aDictionary.Add("CELLNAME", cellName);
                        aDictionary.Add("FREQ", freq);
                        aDictionary.Add("CELLID", cellId);
                        gcellFreqNewData.Add(aDictionary);
                    }
                }

            }

            foreach (Dictionary<string, string> dictionary in gcellFreqNewData)
            {
                string nonQuery = "insert into [GCELLFREQ_FREQ$] (BSCName,CELLNAME,FREQ,CELLID) VALUES('"
                    + dictionary["BSCName"] + "','" + dictionary["CELLNAME"] + "','"
                    + dictionary["FREQ"] + "','" + dictionary["CELLID"] + "')";
                nonQueryCommands.Add(nonQuery);
            }
        }



        private void RemoveAllDataFromGCELLFREQ_FREQ()
        {
            string nonQuery = "update [GCELLFREQ_FREQ$] set FREQ = '', BSCName = '', CELLID='',CELLNAME='' where BSCName <> 'BSC Name';";
            nonQueryCommands.Add(nonQuery);
            RunNonQueryCommandsInDBFile(dbFile);
            nonQueryCommands = new List<string>();
        }


        private void LoadTemplateData(string dbFileName)
        {
            DataSet aSet = IOFileOperation.ReadExcelFile(dbFileName, "GCELL");
            List<string> cols = new List<string>();
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }

            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                foreach (string col in cols)
                {
                    aDictionary.Add(col, dataRow[col].ToString());
                }
                GCELL.Add(aDictionary);
            }

            cols = new List<string>();
            aSet = IOFileOperation.ReadExcelFile(dbFileName, "GCELLFREQ_FREQ");
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }

            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                foreach (string col in cols)
                {
                    aDictionary.Add(col, dataRow[col].ToString());
                }
                GCELLFREQ.Add(aDictionary);
            }



            cols = new List<string>();
            aSet = IOFileOperation.ReadExcelFile(dbFileName, "GTRX");
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }

            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                foreach (string col in cols)
                {
                    aDictionary.Add(col, dataRow[col].ToString());
                }
                GTRX.Add(aDictionary);
            }


            cols = new List<string>();
            aSet = IOFileOperation.ReadExcelFile(dbFileName, "GCELLMAGRP");
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }

            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                foreach (string col in cols)
                {
                    aDictionary.Add(col, dataRow[col].ToString());
                }
                GCELLMAGRP.Add(aDictionary);
            }

        }


        private void ResetTemplateDataFields()
        {
            GCELL = new List<Dictionary<string, string>>();
            GCELLFREQ = new List<Dictionary<string, string>>();
            GTRX = new List<Dictionary<string, string>>();
            GCELLMAGRP = new List<Dictionary<string, string>>();
        }


        private void GenerateNonQueryCommandsForGTRXCHANHOP()
        {
            foreach (Dictionary<string, string> dictionary in GTRXCHANHOP)
            {
                string command = "update [GTRXCHANHOP$] set TRXHOPINDEX='" + dictionary["TRXHOPINDEX"] + "', TRXMAIO='" + dictionary["TRXMAIO"]
                    + "' where BSCName='" + dictionary["BSCName"] + "' and TRXID='" + dictionary["TRXID"] + "' and CHNO='" + dictionary["CHNO"] + "';";
                nonQueryCommands.Add(command);
            }
        }

        private void GenerateNonQueryCommandsForGTRXHOP()
        {
            foreach (Dictionary<string, string> dictionary in GTRXHOP)
            {
                string command = "update [GTRXHOP$] set HOPTYPE='" + dictionary["HOPTYPE"] + "' where BSCName='"+dictionary["BSCName"]+"' and TRXID='"+dictionary["TRXID"]+"';";
                nonQueryCommands.Add(command);
            }
        }

        private void GenerateNonQueryCommandsForCreateGCELLMAGRP(string dbFile)
        {
            List<string> cols = new List<string>();
            DataSet aSet = IOFileOperation.ReadExcelFile(dbFile, "GCELLMAGRP");
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }

            Dictionary<string, Dictionary<string, string>> gcellMaGrpDataExistData = new Dictionary<string, Dictionary<string, string>>();
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                if (dataRow["BSCName"].ToString().Trim().Length > 0)
                {
                    if (dataRow["BSCName"].ToString() != "BSC Name" && dataRow["BSCName"].ToString().Trim().Length > 0)
                    {
                        string index = "BSC-" + dataRow["BSCName"].ToString() + "_CELLID-" + dataRow["CELLID"].ToString() +
                                   "_HOPINDEX-" + dataRow["HOPINDEX"].ToString();
                        Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                        foreach (string col in cols)
                        {
                            aDictionary.Add(col, dataRow[col].ToString());
                        }
                        gcellMaGrpDataExistData.Add(index, aDictionary);
                    }
                }
            }

            string command = string.Empty;
            foreach (Dictionary<string, string> dictionary in CreateGCELLMAGRP)
            {

                string index = "BSC-" + dictionary["BSCName"] + "_CELLID-" + dictionary["CELLID"] +
                               "_HOPINDEX-" + dictionary["HOPINDEX"];

                //if (gcellMaGrpDataExistData.ContainsKey(index))
                //{
                //    string tempCommand = "set ";

                //    foreach (KeyValuePair<string, string> keyValuePair in dictionary)
                //    {
                //        tempCommand += keyValuePair.Key + "='" + keyValuePair.Value + "', ";
                //    }
                //    tempCommand = tempCommand.Trim().Substring(0, tempCommand.Trim().LastIndexOf(','));


                //    command = "update [GCELLMAGRP$]  " + tempCommand + " where BSCName='" + gcellMaGrpDataExistData[index]["BSCName"]
                //           + "' and CELLID='" + gcellMaGrpDataExistData[index]["CELLID"] + "' and HOPINDEX='" + gcellMaGrpDataExistData[index]["HOPINDEX"] + "';";
                //    nonQueryCommands.Add(command);
                 
                //}
           

                command = "insert into [GCELLMAGRP$] ( ";
                foreach (KeyValuePair<string, string> keyValuePair in dictionary)
                {
                    command += keyValuePair.Key + ", ";
                }
                command = command.Trim().Substring(0, command.LastIndexOf(',')) + ") values (";
                foreach (KeyValuePair<string, string> keyValuePair in dictionary)
                {
                    command += "'" + keyValuePair.Value + "', ";
                }
                command = command.Trim().Substring(0, command.LastIndexOf(',')) + ");";
                nonQueryCommands.Add(command);

            }

           
        }

        private void GenerateNonQueryCommandsForDeleteGCELLMAGRP(string dbFile)
        {
            List<string> cols = new List<string>();
            DataSet aSet = IOFileOperation.ReadExcelFile(dbFile, "GCELLMAGRP");

            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }


            string tempCommand = "set ";
            foreach (string col in cols)
            {
                tempCommand += col + "=' ', ";
            }
            tempCommand = tempCommand.Trim().Substring(0, tempCommand.Trim().LastIndexOf(','));
            

            foreach (Dictionary<string, string> dictionary in DeleteGCELLMAGRP)
            {
                string command = "update [GCELLMAGRP$]  " + tempCommand + " where BSCName='" + dictionary["BSCName"]
                    + "' and CELLID='" + dictionary["CELLID"] + "' and HOPINDEX='" + dictionary["HOPINDEX"] + "';";
                nonQueryCommands.Add(command);
            }
        }
        private void RunNonQueryCommandsInDBFile(string dbFileName)
        {
            int index = 0;
            if (logDictionary.Keys.Count != 0)
            {
                index = Convert.ToInt32(logDictionary.Keys.Last()) + 1;
            }
            int totalCommand = nonQueryCommands.Where(i => i != "").Count();
            string log = string.Empty;
            string errorLog = string.Empty;
           
            ExecuteNonQueryOnExcel aNonQueryOnExcel = new ExecuteNonQueryOnExcel(dbFileName);


            foreach (string nonQuery in nonQueryCommands)
            {
                if (nonQuery.Trim().Length != 0)
                {
                    log = "Running Command(HOP): " + nonQuery + "\r\n";
                    Console.WriteLine("Running Command(" + ++index + "/" + totalCommand + "): " + nonQuery);
                    int affectedRows = aNonQueryOnExcel.ExecuteCommandOnExcelFile(nonQuery);

                    Console.WriteLine("Affected Rows(HOP): " + affectedRows);
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
        private void LoadInputFile(string inputFile)
        {
            int index = 0;
            DataSet aSet = IOFileOperation.ReadExcelFile(inputFile, "Delete GCELLMAGRP");
            List<string> cols = new List<string>();
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                if (index != 0)
                {
                    Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                    foreach (string col in cols)
                    {
                        aDictionary.Add(col, dataRow[col].ToString());
                    }
                    DeleteGCELLMAGRP.Add(aDictionary);
                }
                index++;
            }



            index = 0;
            aSet = IOFileOperation.ReadExcelFile(inputFile, "Create GCELLMAGRP");
            cols = new List<string>();
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                if (index != 0)
                {
                    Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                    foreach (string col in cols)
                    {
                        aDictionary.Add(col, dataRow[col].ToString());
                    }

                    CreateGCELLMAGRP.Add(aDictionary);
                }
                index++;
            }



            index = 0;
            aSet = IOFileOperation.ReadExcelFile(inputFile, "GTRXHOP");
            cols = new List<string>();
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                if (index != 0)
                {
                    Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                    foreach (string col in cols)
                    {
                        aDictionary.Add(col, dataRow[col].ToString());
                    }

                    GTRXHOP.Add(aDictionary);
                }
                index++;
            }


            index = 0;
            aSet = IOFileOperation.ReadExcelFile(inputFile, "GTRXCHANHOP");
            cols = new List<string>();
            foreach (DataColumn dataColumn in aSet.Tables[0].Columns)
            {
                cols.Add(dataColumn.ColumnName);
            }
            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
                if (index != 0)
                {
                    Dictionary<string, string> aDictionary = new Dictionary<string, string>();
                    foreach (string col in cols)
                    {
                        aDictionary.Add(col, dataRow[col].ToString());
                    }

                    GTRXCHANHOP.Add(aDictionary);
                }
                index++;
            }



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
            foreach (string s in allNonQuery)
            {
                aqStreamWriter.WriteLine(s);
            }
            aqStreamWriter.Close();

        }
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

        public void CloseConnection()
        {
            MyConnection.Close();
        }

    }
}
