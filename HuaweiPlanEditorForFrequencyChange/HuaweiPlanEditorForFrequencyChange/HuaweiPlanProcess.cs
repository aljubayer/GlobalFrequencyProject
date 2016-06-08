using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using ManiacProject.Libs;
//using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using MoreLinq;


namespace HuaweiPlanEditorForFrequencyChange
{
    public class HuaweiPlanProcess
    {

        private void ResetTemplateDataFields()
        {
            GCELL = new List<Dictionary<string, string>>();
            GCELLFREQ = new List<Dictionary<string, string>>();
            GTRX = new List<Dictionary<string, string>>();
            GCELLMAGRP = new List<Dictionary<string, string>>();
        }

        private List<Dictionary<string, string>> GCELL = new List<Dictionary<string, string>>();
        private List<Dictionary<string, string>> GCELLFREQ = new List<Dictionary<string, string>>();
        private List<Dictionary<string, string>> GTRX = new List<Dictionary<string, string>>();
        private List<Dictionary<string, string>> GCELLMAGRP = new List<Dictionary<string, string>>();
        private List<FrequencyChangeData> woInputFrequency = new List<FrequencyChangeData>();
        private List<BSICChangeData> woInputBSIC = new List<BSICChangeData>();
        private List<string> nonQueryCommandList = new List<string>();
        private string dbFileName = string.Empty;

        List<Dictionary<string, string>> GTRXCHANHOP = new List<Dictionary<string, string>>();


        private Dictionary<string, string> logDictionary = new Dictionary<string, string>();
        private Dictionary<string, string> errorLogDictionary = new Dictionary<string, string>();
        private List<string> allNonQuery = new List<string>();


        public string ProcessPlan(string dbFileName, string inputFile)
        {
            this.dbFileName = dbFileName;
            Console.WriteLine("Reading WO Input.....");
            ReadWOInput(inputFile);
            Console.WriteLine("Reading Template.....");
            LoadTemplateData(dbFileName);
            Console.WriteLine("Loading Cell Information.....");
            LoadBSCNameCellIdFromGCELL();
            Console.WriteLine("Loading TRXID Information.....");
            LoadTRXIDFromGTRX();
            Console.WriteLine("Loading MA Group information.....");
            LoadOldFreqColumnNameInGCELLMAGRP();
            if (HasOutOfBandFrequency())
            {
                return "WO input contains out of band frequency. Please check \"Output.txt\" file.";
            }

            if (!IsCurrentConfigurationOfWOOK())
            {
                return "WO input contains configuration mismatch. Please check \"Output.txt\" file.";
            }

            //Console.WriteLine("Adding TRXID Data in GCELLFREQ_FREQ......");
            //AddNewColumnTRXIDInGCELLFREQ_FREQ();

            Console.WriteLine("Generating non Query Except GCELLFREQ......");
            GenerateNonQueryCommandsExceptGCELLFREQ();
            allNonQuery.AddRange(nonQueryCommandList);
            Console.WriteLine("Running non query commands......");
            RunNonQueryCommandsInDBFile(dbFileName);

            Console.WriteLine("Processing GCELLFREQ_FREQ......");
            //ProcessDataInGCELLFREQ();
            //allNonQuery.AddRange(nonQueryCommandList);

      

            //WriteLogFiles();
            return "";
        }

        public void WriteLogFiles()
        {
            StreamWriter aWriterLog = new StreamWriter("log.txt");
            StreamWriter aWriterErrorLog = new StreamWriter("errorLog.txt");

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

            StreamWriter aqStreamWriter = new StreamWriter("nonquery.txt");
            foreach (string s in allNonQuery)
            {
                aqStreamWriter.WriteLine(s);
            }
            aqStreamWriter.Close();

        }

        private void ProcessDataInGCELLFREQ()
        {
            nonQueryCommandList = new List<string>();
            ResetTemplateDataFields();
            LoadTemplateData(dbFileName);
            RemoveAllDataFromGCELLFREQ_FREQ();
            LoadGCELLFREQData();
            RunNonQueryCommandsInDBFile(dbFileName);
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
                            if (keyValuePair.Key.Contains("FREQ"))
                            {
                                if (!freqs.Contains(keyValuePair.Value) && keyValuePair.Value.Trim().Length != 0)
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
                nonQueryCommandList.Add(nonQuery);
            }
        }

        private void RemoveAllDataFromGCELLFREQ_FREQ()
        {
            string nonQuery =
                "update [GCELLFREQ_FREQ$] set FREQ = '', BSCName = '', CELLID='',CELLNAME='' where FREQ <> 'Frequency';";
            nonQueryCommandList.Add(nonQuery);
            RunNonQueryCommandsInDBFile(dbFileName);
            nonQueryCommandList = new List<string>();
        }



        private void RemoveColumnTRXIDInGCELLFREQ_FREQ()
        {

            string nonQuery = "UPDATE [GCELLFREQ_FREQ$] SET TRXID='';";
            ExecuteNonQueryOnExcel aExcel = new ExecuteNonQueryOnExcel(dbFileName);
            aExcel.ExecuteCommandOnExcelFile(nonQuery);
            aExcel.CloseConnection();
            ExecuteNonQueryOnExcel.RemoveTRXIDFromGCELLFREQ_FREQColumn(dbFileName, "GCELLFREQ_FREQ");


        }

        private void LoadTRXIDFromGTRX()
        {
            foreach (FrequencyChangeData frequencyChangeData in woInputFrequency)
            {
                var trxData =
                    GTRX.Where(
                        i => i["BSCName"] == frequencyChangeData.BSCName && i["CELLID"] == frequencyChangeData.CELLID
                             && i["FREQ"] == frequencyChangeData.OldFrequency);
                foreach (Dictionary<string, string> dictionary in trxData)
                {
                    frequencyChangeData.TRXId = dictionary["TRXID"];
                }
            }
        }

        private bool IsCurrentConfigurationOfWOOK()
        {
            StreamWriter sw = new StreamWriter("output.txt", true);
            string logData = string.Empty;
            foreach (FrequencyChangeData frequencyChangeData in woInputFrequency)
            {
                if (!HasFrequencyInProvidedCell(frequencyChangeData))
                {
                    logData += "BSC: " + frequencyChangeData.BSCName + ",LAC: " + frequencyChangeData.LAC + ",CI: " +
                               frequencyChangeData.CI
                               + " provided current configuration does not match with actual configuration(" +
                               frequencyChangeData.OldFrequency + ").\r\n";
                }
            }
            sw.WriteLine(logData);
            sw.Close();

            if (logData.Trim().Length == 0)
            {
                return true;
            }
            return false;
        }

        private bool HasFrequencyInProvidedCell(FrequencyChangeData frequencyChangeData)
        {

            var currentData =
                GTRX.Where(
                    i =>
                        i["FREQ"] == frequencyChangeData.OldFrequency && i["BSCName"] == frequencyChangeData.BSCName &&
                        i["CELLID"] == frequencyChangeData.CELLID);
            foreach (Dictionary<string, string> dictionary in currentData)
            {
                return true;
            }
            return false;
        }



        private bool HasOutOfBandFrequency()
        {
            string outOfBandLog = string.Empty;
            foreach (FrequencyChangeData frequencyChangeData in woInputFrequency)
            {

                if (((Convert.ToInt32(frequencyChangeData.NewFrequency) < 27 ||
                      Convert.ToInt32(frequencyChangeData.NewFrequency) > 50) &&
                     frequencyChangeData.NewFrequency.Length == 2) ||
                    (Convert.ToInt32(frequencyChangeData.NewFrequency) < 27
                     || Convert.ToInt32(frequencyChangeData.NewFrequency) > 50) &&
                    frequencyChangeData.NewFrequency.Length == 1)
                {
                    outOfBandLog += "BSC: " + frequencyChangeData.BSCName + ",LAC: " + frequencyChangeData.LAC + ",CI: " +
                                    frequencyChangeData.CI
                                    + " Contains Out Of Band Frequency(" + frequencyChangeData.NewFrequency + ").\r\n";

                }

                if ((Convert.ToInt32(frequencyChangeData.NewFrequency) < 722 ||
                     Convert.ToInt32(frequencyChangeData.NewFrequency) > 770) &&
                    frequencyChangeData.NewFrequency.Length == 3)
                {
                    outOfBandLog += "BSC: " + frequencyChangeData.BSCName + ",LAC: " + frequencyChangeData.LAC + ",CI: " +
                                    frequencyChangeData.CI
                                    + " Contains Out Of Band Frequency(" + frequencyChangeData.NewFrequency + ").\r\n";
                }

            }

            StreamWriter sw = new StreamWriter("output.txt");
            if (outOfBandLog.Length != 0)
            {

                sw.WriteLine(outOfBandLog);
                sw.Close();
                return true;
            }
            sw.Close();
            return false;
        }

        private void RunNonQueryCommandsInDBFile(string dbFileName)
        {
            int index = 0;
            if (logDictionary.Keys.Count != 0)
            {
                index = Convert.ToInt16(logDictionary.Keys.Last()) + 1;
            }

            int totalCommand = nonQueryCommandList.Where(i => i != "").Count();
            string log = string.Empty;
            string errorLog = string.Empty;

            ExecuteNonQueryOnExcel aNonQueryOnExcel = new ExecuteNonQueryOnExcel(dbFileName);


            foreach (string nonQuery in nonQueryCommandList)
            {
                if (nonQuery.Trim().Length != 0)
                {
                    log = "Running Command(Freq Change): " + nonQuery + "\r\n";
                    Console.WriteLine("Running Command(" + ++index + "/" + totalCommand + "): " + nonQuery);
                    int affectedRows = aNonQueryOnExcel.ExecuteCommandOnExcelFile(nonQuery);

                    Console.WriteLine("Affected Rows(Freq Change): " + affectedRows);
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

        private void GenerateNonQueryCommandsExceptGCELLFREQ()
        {

            foreach (FrequencyChangeData frequencyChangeData in woInputFrequency)
            {

                string nonQuery = string.Empty;

                nonQuery =
                    "update [GTRX$] set FREQ = '" + frequencyChangeData.NewFrequency + "' where BSCName='" +
                    frequencyChangeData.BSCName
                    + "' and CELLID='" + frequencyChangeData.CELLID + "' and FREQ='" + frequencyChangeData.OldFrequency +
                    "' and TRXID='" + frequencyChangeData.TRXId + "'; ";

                nonQueryCommandList.Add(nonQuery);

                if (!Has900RFHopping(frequencyChangeData))
                {
                    if (frequencyChangeData.OldFreqColumnNameInGCELLMAGRP != null)
                    {

                        foreach (Dictionary<string, string> dictionary in frequencyChangeData.HopIndex)
                        {
                            if (dictionary.Count > 1)
                            {
                                nonQuery = "update [GCELLMAGRP$] set " + dictionary["ColumnName"] +
                                  " = '"
                                  + frequencyChangeData.NewFrequency + "' where BSCName='" +
                                  frequencyChangeData.BSCName +
                                  "' and CELLID='" + frequencyChangeData.CELLID + "' and " +
                                  dictionary["ColumnName"] + " = '"
                                  + frequencyChangeData.OldFrequency + "' and HOPINDEX='" + dictionary["HopIndex"] + "';";

                                nonQueryCommandList.Add(nonQuery);
                            }
                            
                        }
                       
                       
                    }
                }

                nonQueryCommandList.Add("");
                nonQueryCommandList.Add("");

            }



            foreach (BSICChangeData bsicChangeData in woInputBSIC)
            {

            

                string bcc = GetBCC(bsicChangeData.NewBSIC);
                string ncc = GetNCC(bsicChangeData.NewBSIC);
                string nonQuery = "update [GCELL$] set NCC='" + ncc + "' where BSCName='" + bsicChangeData.BSCName +
                                  "' and CELLID='" + bsicChangeData.CELLID + "'";
                nonQueryCommandList.Add(nonQuery);
                nonQuery = "update [GCELL$] set BCC='" + bcc + "' where BSCName='" + bsicChangeData.BSCName +
                           "' and CELLID='" + bsicChangeData.CELLID + "'";
                nonQueryCommandList.Add(nonQuery);

                nonQuery = "update [GCELLMAGRP$] set TSC='" + bcc + "' where BSCName='" + bsicChangeData.BSCName +
                           "' and CELLID='" + bsicChangeData.CELLID + "'";
                nonQueryCommandList.Add(nonQuery);



                nonQueryCommandList.Add("");
            }




        }



        private bool AlreadyFrequencyHasInGCELLFREQ(FrequencyChangeData frequencyChangeData)
        {

            if (GCELLFREQ.Exists(i => i["BSCName"] == frequencyChangeData.BSCName
                                      && i["CELLID"] == frequencyChangeData.CELLID
                                      && i["TRXID"].Trim().Length == 0
                                      && i["FREQ"] == frequencyChangeData.NewFrequency))
            {
                return true;
            }


            return false;
        }

        private bool Has900RFHopping(FrequencyChangeData frequencyChangeData)
        {
            if (frequencyChangeData.NewFrequency.Trim().Length == 3)
            {
                return false;
            }

            if (GCELLMAGRP.Exists(i => i["BSCName"] == frequencyChangeData.BSCName
                                       && i["CELLID"] == frequencyChangeData.CELLID
                                       && i["HOPMODE"] == "RF_FH"
                                       && i["FREQ1"].Trim().Length == 2))
            {

                return true;
            }

            return false;
        }




        private void AddNewColumnTRXIDInGCELLFREQ_FREQ()
        {

            ExecuteNonQueryOnExcel aExecuteNonQueryOnExcel = new ExecuteNonQueryOnExcel(dbFileName);
            string nQuery =
                "CREATE TABLE [GCELLFREQ_FREQ$] (BSCName TEXT(100),CELLNAME TEXT(100),FREQ TEXT(100),CELLID TEXT(100),TRXID TEXT(100),REQUIRE TEXT(100));";

            aExecuteNonQueryOnExcel.ExecuteCommandOnExcelFile(nQuery);
            aExecuteNonQueryOnExcel.CloseConnection();

            UpdateTRXIdInGCELLFREQ_FREQ();
        }

        private void UpdateTRXIdInGCELLFREQ_FREQ()
        {
            ExecuteNonQueryOnExcel aExecuteNonQueryOnExcel = new ExecuteNonQueryOnExcel(dbFileName);

            foreach (Dictionary<string, string> dictionary in GCELLFREQ)
            {
                var cellData =
                    GTRX.Where(i => i["BSCName"] == dictionary["BSCName"] && i["CELLID"] == dictionary["CELLID"]
                                    && i["FREQ"] == dictionary["FREQ"]);

                foreach (Dictionary<string, string> aCellData in cellData)
                {
                    dictionary["TRXID"] = aCellData["TRXID"];
                }

                if (!dictionary.ContainsKey("TRXID"))
                {
                    dictionary.Add("TRXID", "");
                }

            }

            foreach (Dictionary<string, string> dictionary in GCELLFREQ)
            {
                if (dictionary["TRXID"].Length != 0)
                {
                    string nonQuery = "update [GCELLFREQ_FREQ$] set TRXID = '" + dictionary["TRXID"] +
                                      "' where BSCName='" + dictionary["BSCName"] + "' and CELLID='" +
                                      dictionary["CELLID"] + "' and FREQ='" + dictionary["FREQ"] + "';";

                    aExecuteNonQueryOnExcel.ExecuteCommandOnExcelFile(nonQuery);
                    Console.WriteLine(nonQuery);
                }

            }
            aExecuteNonQueryOnExcel.CloseConnection();
        }



        private void LoadOldFreqColumnNameInGCELLMAGRP()
        {
            foreach (FrequencyChangeData frequencyChangeData in woInputFrequency)
            {
                var hopIndex = GTRXCHANHOP.Where(
                    i => i["BSCName"] == frequencyChangeData.BSCName && i["TRXID"] == frequencyChangeData.TRXId);

                foreach (Dictionary<string, string> dictionary in hopIndex)
                {
                   Dictionary<string,string> aDictionary = new Dictionary<string, string>();
                   aDictionary.Add("HopIndex", dictionary["TRXHOPINDEX"]);
                   frequencyChangeData.HopIndex.Add(aDictionary);
                }

                var cellData =
                    GCELLMAGRP.Where(
                        i => i["BSCName"] == frequencyChangeData.BSCName && i["CELLID"] == frequencyChangeData.CELLID);
                foreach (Dictionary<string, string> dictionary in cellData)
                {
                    foreach (KeyValuePair<string, string> keyValuePair in dictionary)
                    {
                        if (keyValuePair.Key.Contains("FREQ") &&
                            keyValuePair.Value.ToString() == frequencyChangeData.OldFrequency)
                        {
                            foreach (Dictionary<string,string> aHopIndex in frequencyChangeData.HopIndex)
                            {
                                if (aHopIndex["HopIndex"] == dictionary["HOPINDEX"] && !aHopIndex.ContainsKey("ColumnName"))
                                {
                                    aHopIndex.Add("ColumnName",keyValuePair.Key);
                                }
                            }
                            frequencyChangeData.OldFreqColumnNameInGCELLMAGRP = keyValuePair.Key;
                        }
                    }
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


        private void LoadBSCNameCellIdFromGCELL()
        {
            foreach (FrequencyChangeData frequencyChangeData in woInputFrequency)
            {
                var cellData =
                    GCELL.Where(i => i["LAC"] == frequencyChangeData.LAC && i["CI"] == frequencyChangeData.CI);
                foreach (Dictionary<string, string> dictionary in cellData)
                {
                    frequencyChangeData.BSCName = dictionary["BSCName"];
                    frequencyChangeData.CELLID = dictionary["CELLID"];
                }

            }


            foreach (BSICChangeData bsicChangeData in woInputBSIC)
            {
                var cellData =
                    GCELL.Where(i => i["LAC"] == bsicChangeData.LAC && i["CI"] == bsicChangeData.CI);
                foreach (Dictionary<string, string> dictionary in cellData)
                {
                    bsicChangeData.BSCName = dictionary["BSCName"];
                    bsicChangeData.CELLID = dictionary["CELLID"];
                }

            }
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


            cols = new List<string>();
            aSet = IOFileOperation.ReadExcelFile(dbFileName, "GTRXCHANHOP");
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
                GTRXCHANHOP.Add(aDictionary);
            }

        }


        private void ReadWOInput(string inputFile)
        {
            DataSet aSet = IOFileOperation.ReadExcelFile(inputFile, "Input");

            int count = aSet.Tables[0].Rows.Count;

            foreach (DataRow dataRow in aSet.Tables[0].Rows)
            {
               

                if (dataRow["NEW BCCH"].ToString().Trim().Length != 0)
                {
                    FrequencyChangeData aData = new FrequencyChangeData();
                    aData.LAC = dataRow["LAC"].ToString().Trim();
                    aData.CI = dataRow["CI"].ToString().Trim();
                    aData.OldFrequency = dataRow["BCCH"].ToString().Trim();
                    aData.NewFrequency = dataRow["NEW BCCH"].ToString().Trim();
                    woInputFrequency.Add(aData);
                }


                if (dataRow["NEW TCH1"].ToString().Trim().Length != 0)
                {
                    FrequencyChangeData aData1 = new FrequencyChangeData();
                    aData1.LAC = dataRow["LAC"].ToString().Trim();
                    aData1.CI = dataRow["CI"].ToString().Trim();
                    aData1.OldFrequency = dataRow["TCH1"].ToString().Trim();
                    aData1.NewFrequency = dataRow["NEW TCH1"].ToString().Trim();
                    woInputFrequency.Add(aData1);
                }

                if (dataRow["NEW TCH2"].ToString().Trim().Length != 0)
                {
                    FrequencyChangeData aData2 = new FrequencyChangeData();
                    aData2.LAC = dataRow["LAC"].ToString().Trim();
                    aData2.CI = dataRow["CI"].ToString().Trim();
                    aData2.OldFrequency = dataRow["TCH2"].ToString().Trim();
                    aData2.NewFrequency = dataRow["NEW TCH2"].ToString().Trim();
                    woInputFrequency.Add(aData2);
                }

                if (dataRow["NEW TCH3"].ToString().Trim().Length != 0)
                {
                    FrequencyChangeData aData3 = new FrequencyChangeData();
                    aData3.LAC = dataRow["LAC"].ToString().Trim();
                    aData3.CI = dataRow["CI"].ToString().Trim();
                    aData3.OldFrequency = dataRow["TCH3"].ToString().Trim();
                    aData3.NewFrequency = dataRow["NEW TCH3"].ToString().Trim();
                    woInputFrequency.Add(aData3);
                }


                if (dataRow["NEW TCH4"].ToString().Trim().Length != 0)
                {
                    FrequencyChangeData aData4 = new FrequencyChangeData();
                    aData4.LAC = dataRow["LAC"].ToString().Trim();
                    aData4.CI = dataRow["CI"].ToString().Trim();
                    aData4.OldFrequency = dataRow["TCH4"].ToString().Trim();
                    aData4.NewFrequency = dataRow["NEW TCH4"].ToString().Trim();
                    woInputFrequency.Add(aData4);
                }


                if (dataRow["NEW TCH5"].ToString().Length != 0)
                {
                    FrequencyChangeData aData5 = new FrequencyChangeData();
                    aData5.LAC = dataRow["LAC"].ToString().Trim();
                    aData5.CI = dataRow["CI"].ToString().Trim();
                    aData5.OldFrequency = dataRow["TCH5"].ToString().Trim();
                    aData5.NewFrequency = dataRow["NEW TCH5"].ToString().Trim();
                    woInputFrequency.Add(aData5);
                }

                if (dataRow["NEW TCH6"].ToString().Length != 0)
                {
                    FrequencyChangeData aData6 = new FrequencyChangeData();
                    aData6.LAC = dataRow["LAC"].ToString().Trim();
                    aData6.CI = dataRow["CI"].ToString().Trim();
                    aData6.OldFrequency = dataRow["TCH6"].ToString().Trim();
                    aData6.NewFrequency = dataRow["NEW TCH6"].ToString().Trim();
                    woInputFrequency.Add(aData6);
                }

                if (dataRow["NEW TCH7"].ToString().Length != 0)
                {
                    FrequencyChangeData aData7 = new FrequencyChangeData();
                    aData7.LAC = dataRow["LAC"].ToString().Trim();
                    aData7.CI = dataRow["CI"].ToString().Trim();
                    aData7.OldFrequency = dataRow["TCH7"].ToString().Trim();
                    aData7.NewFrequency = dataRow["NEW TCH7"].ToString().Trim();
                    woInputFrequency.Add(aData7);
                }


                if (dataRow["NEW TCH8"].ToString().Trim().Length != 0)
                {
                    FrequencyChangeData aData8 = new FrequencyChangeData();
                    aData8.LAC = dataRow["LAC"].ToString().Trim();
                    aData8.CI = dataRow["CI"].ToString();
                    aData8.OldFrequency = dataRow["TCH8"].ToString().Trim();
                    aData8.NewFrequency = dataRow["NEW TCH8"].ToString().Trim();
                    woInputFrequency.Add(aData8);
                }
                if (dataRow["NEW TCH9"].ToString().Length != 0)
                {
                    FrequencyChangeData aData9 = new FrequencyChangeData();
                    aData9.LAC = dataRow["LAC"].ToString().Trim();
                    aData9.CI = dataRow["CI"].ToString().Trim();
                    aData9.OldFrequency = dataRow["TCH9"].ToString().Trim();
                    aData9.NewFrequency = dataRow["NEW TCH9"].ToString().Trim();
                    woInputFrequency.Add(aData9);
                }


                if (dataRow["NEW TCH10"].ToString().Length != 0)
                {
                    FrequencyChangeData aData10 = new FrequencyChangeData();
                    aData10.LAC = dataRow["LAC"].ToString().Trim();
                    aData10.CI = dataRow["CI"].ToString().Trim();
                    aData10.OldFrequency = dataRow["TCH10"].ToString().Trim();
                    aData10.NewFrequency = dataRow["NEW TCH10"].ToString().Trim();
                    woInputFrequency.Add(aData10);
                }


                if (dataRow["NEW TCH11"].ToString().Length != 0)
                {
                    FrequencyChangeData aData11 = new FrequencyChangeData();
                    aData11.LAC = dataRow["LAC"].ToString().Trim();
                    aData11.CI = dataRow["CI"].ToString().Trim();
                    aData11.OldFrequency = dataRow["TCH11"].ToString().Trim();
                    aData11.NewFrequency = dataRow["NEW TCH11"].ToString().Trim();
                    woInputFrequency.Add(aData11);
                }

                string bsic = dataRow["NEW BSIC"].ToString().Trim();
                
                if (dataRow["NEW BSIC"].ToString().Trim().Length != 0)
                {
                    BSICChangeData aData12 = new BSICChangeData();
                    aData12.LAC = dataRow["LAC"].ToString().Trim();
                    aData12.CI = dataRow["CI"].ToString().Trim();
                    aData12.OldBSIC = dataRow["BSIC"].ToString().Trim();
                    aData12.NewBSIC = dataRow["NEW BSIC"].ToString().Trim();
                    woInputBSIC.Add(aData12);
                }
            }
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
        public List<Dictionary<string,string>> HopIndex = new List<Dictionary<string, string>>();

    }

    public class BSICChangeData
    {
        public string BSCName { set; get; }
        public string CELLID { set; get; }
        public string LAC { set; get; }
        public string CI { set; get; }
        public string OldBSIC { set; get; }
        public string NewBSIC { set; get; }
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
            OleDbDataAdapter da = new OleDbDataAdapter(query , con);
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
