using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;

namespace csvReader
{
    class Program
    {
        static void Main(string[] args)
        {
            excelWrier myW = new excelWrier();
            DataSet resultSet = new DataSet();
            #region 
            scvDataReader sdataReader = new scvDataReader();
            Console.WriteLine("Please write the NW result path down");
            resultSet.Tables.Add(sdataReader.getData(checkDirectory(Console.ReadLine()), "NW"));

            Console.WriteLine("Please write the SN result path down");
            resultSet.Tables.Add(sdataReader.getData(checkDirectory(Console.ReadLine()), "SN"));

            Console.WriteLine("Please write the SV result path down ");
            resultSet.Tables.Add(sdataReader.getData(checkDirectory(Console.ReadLine()), "SV"));

            Console.WriteLine("Please write the WV result path down");
            resultSet.Tables.Add(sdataReader.getData(checkDirectory(Console.ReadLine()), "WV"));

            Console.WriteLine("read data finished......");
            Console.WriteLine("Please write down the folder path where export data to : ");

            string exportPath = Console.ReadLine();

            for (int t = 0; t < resultSet.Tables.Count; t++)
            {
                myW.doExport(resultSet.Tables[t], exportPath + "\\resport_" + resultSet.Tables[t].TableName + ".xlsx");
                Console.WriteLine("export data finished to " + exportPath + "\\resport_" + resultSet.Tables[t].TableName + ".xlsx");
            }
            #endregion

            #region
            DataTable totalTimeCompare = new DataTable("TotalTimeCompare");
            totalTimeCompare.Columns.Add("AudioName", typeof(string));
            totalTimeCompare.Columns.Add("SpeexVlingo", typeof(string));
            totalTimeCompare.Columns.Add("WavVlingo", typeof(string));
            totalTimeCompare.Columns.Add("SpeexNuance", typeof(string));
            totalTimeCompare.Columns.Add("WavNuance", typeof(string));

            totalTimeCompare.Columns.Add("ExpectedResult", typeof(string));

            totalTimeCompare.Columns.Add("SVResult_best", typeof(string));
            totalTimeCompare.Columns.Add("SVResult_mid", typeof(string));
            totalTimeCompare.Columns.Add("SVResult_worst", typeof(string));

            totalTimeCompare.Columns.Add("WVResult_best", typeof(string));
            totalTimeCompare.Columns.Add("WVResult_mid", typeof(string));
            totalTimeCompare.Columns.Add("WVResult_worst", typeof(string));

            totalTimeCompare.Columns.Add("SNResult_best", typeof(string));
            totalTimeCompare.Columns.Add("SNResult_mid", typeof(string));
            totalTimeCompare.Columns.Add("SNResult_worst", typeof(string));

            totalTimeCompare.Columns.Add("WNResult_best", typeof(string));
            totalTimeCompare.Columns.Add("WNResult_mid", typeof(string));
            totalTimeCompare.Columns.Add("WNResult_worst", typeof(string));

            DataTable serverTimeCompare = totalTimeCompare.Clone();
            serverTimeCompare.TableName = "ServerTimeCompare";

            DataTable accuracyCompare = totalTimeCompare.Clone();
            accuracyCompare.TableName = "AccuracyCompare";

            int totalRowsCount = resultSet.Tables["NW"].Rows.Count;
            for (int r = 0; r < totalRowsCount; r++)
            {
                DataRow ttdr = totalTimeCompare.NewRow();
                ttdr["AudioName"] = resultSet.Tables["NW"].Rows[r]["fileName"];
                ttdr["ExpectedResult"] = resultSet.Tables["NW"].Rows[r]["expectedResult"];

                ttdr["WavNuance"] = resultSet.Tables["NW"].Rows[r]["totalTime_Avg"];
                ttdr["WNResult_best"] = resultSet.Tables["NW"].Rows[r]["bestObtainedResult"];
                ttdr["WNResult_mid"] = resultSet.Tables["NW"].Rows[r]["midObtainedResult"];
                ttdr["WNResult_worst"] = resultSet.Tables["NW"].Rows[r]["worstObtainedResult"];

                ttdr["SpeexVlingo"] = resultSet.Tables["SV"].Rows[r]["totalTime_Avg"];
                ttdr["SVResult_best"] = resultSet.Tables["SV"].Rows[r]["bestObtainedResult"];
                ttdr["SVResult_mid"] = resultSet.Tables["SV"].Rows[r]["midObtainedResult"];
                ttdr["SVResult_worst"] = resultSet.Tables["SV"].Rows[r]["worstObtainedResult"];

                ttdr["WavVlingo"] = resultSet.Tables["WV"].Rows[r]["totalTime_Avg"];
                ttdr["WVResult_best"] = resultSet.Tables["WV"].Rows[r]["bestObtainedResult"];
                ttdr["WVResult_mid"] = resultSet.Tables["WV"].Rows[r]["midObtainedResult"];
                ttdr["WVResult_worst"] = resultSet.Tables["WV"].Rows[r]["worstObtainedResult"];

                ttdr["SpeexNuance"] = resultSet.Tables["SN"].Rows[r]["totalTime_Avg"];
                ttdr["SNResult_best"] = resultSet.Tables["SN"].Rows[r]["bestObtainedResult"];
                ttdr["SNResult_mid"] = resultSet.Tables["SN"].Rows[r]["midObtainedResult"];
                ttdr["SNResult_worst"] = resultSet.Tables["SN"].Rows[r]["worstObtainedResult"];

                totalTimeCompare.Rows.Add(ttdr);
                totalTimeCompare.AcceptChanges();

                DataRow stdr = serverTimeCompare.NewRow();
                stdr["AudioName"] = resultSet.Tables["NW"].Rows[r]["fileName"];
                stdr["ExpectedResult"] = resultSet.Tables["NW"].Rows[r]["expectedResult"];

                stdr["WavNuance"] = resultSet.Tables["NW"].Rows[r]["serverTime_Avg"];
                stdr["WNResult_best"] = resultSet.Tables["NW"].Rows[r]["bestObtainedResult"];
                stdr["WNResult_mid"] = resultSet.Tables["NW"].Rows[r]["midObtainedResult"];
                stdr["WNResult_worst"] = resultSet.Tables["NW"].Rows[r]["worstObtainedResult"];

                stdr["SpeexVlingo"] = resultSet.Tables["SV"].Rows[r]["serverTime_Avg"];
                stdr["SVResult_best"] = resultSet.Tables["SV"].Rows[r]["bestObtainedResult"];
                stdr["SVResult_mid"] = resultSet.Tables["SV"].Rows[r]["midObtainedResult"];
                stdr["SVResult_worst"] = resultSet.Tables["SV"].Rows[r]["worstObtainedResult"];

                stdr["WavVlingo"] = resultSet.Tables["WV"].Rows[r]["serverTime_Avg"];
                stdr["WVResult_best"] = resultSet.Tables["WV"].Rows[r]["bestObtainedResult"];
                stdr["WVResult_mid"] = resultSet.Tables["WV"].Rows[r]["midObtainedResult"];
                stdr["WVResult_worst"] = resultSet.Tables["WV"].Rows[r]["worstObtainedResult"];

                stdr["SpeexNuance"] = resultSet.Tables["SN"].Rows[r]["serverTime_Avg"];
                stdr["SNResult_best"] = resultSet.Tables["SN"].Rows[r]["bestObtainedResult"];
                stdr["SNResult_mid"] = resultSet.Tables["SN"].Rows[r]["midObtainedResult"];
                stdr["SNResult_worst"] = resultSet.Tables["SN"].Rows[r]["worstObtainedResult"];

                serverTimeCompare.Rows.Add(stdr);
                serverTimeCompare.AcceptChanges();

                DataRow adr = accuracyCompare.NewRow();
                adr["AudioName"] = resultSet.Tables["NW"].Rows[r]["fileName"];
                adr["ExpectedResult"] = resultSet.Tables["NW"].Rows[r]["expectedResult"];

                adr["WavNuance"] = resultSet.Tables["NW"].Rows[r]["Accuracy_Avg"];
                adr["WNResult_best"] = resultSet.Tables["NW"].Rows[r]["bestObtainedResult"];
                adr["WNResult_mid"] = resultSet.Tables["NW"].Rows[r]["midObtainedResult"];
                adr["WNResult_worst"] = resultSet.Tables["NW"].Rows[r]["worstObtainedResult"];


                adr["SpeexVlingo"] = resultSet.Tables["SV"].Rows[r]["Accuracy_Avg"];
                adr["SVResult_best"] = resultSet.Tables["SV"].Rows[r]["bestObtainedResult"];
                adr["SVResult_mid"] = resultSet.Tables["SV"].Rows[r]["midObtainedResult"];
                adr["SVResult_worst"] = resultSet.Tables["SV"].Rows[r]["worstObtainedResult"];

                adr["WavVlingo"] = resultSet.Tables["WV"].Rows[r]["Accuracy_Avg"];
                adr["WVResult_best"] = resultSet.Tables["WV"].Rows[r]["bestObtainedResult"];
                adr["WVResult_mid"] = resultSet.Tables["WV"].Rows[r]["midObtainedResult"];
                adr["WVResult_worst"] = resultSet.Tables["WV"].Rows[r]["worstObtainedResult"];

                adr["SpeexNuance"] = resultSet.Tables["SN"].Rows[r]["Accuracy_Avg"];
                adr["SNResult_best"] = resultSet.Tables["SN"].Rows[r]["bestObtainedResult"];
                adr["SNResult_mid"] = resultSet.Tables["SN"].Rows[r]["midObtainedResult"];
                adr["SNResult_worst"] = resultSet.Tables["SN"].Rows[r]["worstObtainedResult"];

                accuracyCompare.Rows.Add(adr);
                accuracyCompare.AcceptChanges();

            }
            myW.doExport(totalTimeCompare, exportPath + "\\resport_" + totalTimeCompare.TableName + ".xlsx");
            Console.WriteLine("export data finished to " + exportPath + "\\resport_" + totalTimeCompare.TableName + ".xlsx");
              
            myW.doExport(serverTimeCompare, exportPath + "\\resport_" + serverTimeCompare.TableName + ".xlsx");
            Console.WriteLine("export data finished to " + exportPath + "\\resport_" + serverTimeCompare.TableName + ".xlsx");
              
            myW.doExport(accuracyCompare, exportPath + "\\resport_" + accuracyCompare.TableName + ".xlsx");
            Console.WriteLine("export data finished to " + exportPath + "\\resport_" + accuracyCompare.TableName + ".xlsx");
            #endregion
            Console.WriteLine("Done !");

            Console.ReadLine();
        }

        public static string checkDirectory(string path)
        {
            if (path!=""&&Directory.Exists(path))
            {
                return path;
            }
            else
            {
                Console.WriteLine("Please provide a valid path ...");
                return checkDirectory(Console.ReadLine());
            }

        }


    }
}
