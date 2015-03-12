using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;

namespace csvReader
{
    public class scvDataReader
    {
        public scvDataReader() {

        }

        public DataTable getData(string directoryFolder,string tableName)
        {
            DataTable dt = new DataTable(tableName);
            dt.Columns.Add("fileName", typeof(string));
            dt.Columns.Add("testCount", typeof(int));
            dt.Columns.Add("totalTime_Max", typeof(double));
            dt.Columns.Add("totalTime_Min", typeof(double));
            dt.Columns.Add("totalTime_Avg", typeof(double));
            dt.Columns.Add("serverTime_Max", typeof(double));
            dt.Columns.Add("serverTime_Min", typeof(double));
            dt.Columns.Add("serverTime_Avg", typeof(double));
            dt.Columns.Add("Accuracy_Max", typeof(double));
            dt.Columns.Add("Accuracy_Min", typeof(double));
            dt.Columns.Add("Accuracy_Avg", typeof(double));
            dt.Columns.Add("expectedResult", typeof(string));
            dt.Columns.Add("bestObtainedResult", typeof(string));
            dt.Columns.Add("worstObtainedResult", typeof(string));
            dt.Columns.Add("midObtainedResult", typeof(string));

            DirectoryInfo di = new DirectoryInfo(directoryFolder);
            FileInfo[] files = di.GetFiles();
            //FileInfo[] files = di.GetFiles("*.csv");
            for (int f = 0; f < files.Length; f++)
            {
                dt = updateResultTable(dt, files[f].FullName);
            }
            dt.AcceptChanges();
            return dt;
        }

        private DataTable updateResultTable(DataTable resultTable, string path)
        {
            string strline;
            string[] aryline;
            bool isResultTableEmpty = resultTable.Rows.Count == 0 ? true : false;//

            using (StreamReader mysr = new StreamReader(path))
            {
                int rowNum = 0;
                while ((strline = mysr.ReadLine()) != null)
                {
                    //aryline = strline.Split(new char[] { ',' });
                    strline =  dealCsvLine(strline).Replace("'", "''");//调用函数处理每一行内容
                    aryline = strline.Split(new char[] { '^' });//对处理后的内容进行特殊字符“^”分隔就得到了常规的字符数组了，你

                    #region if the first row is the column name row, skip it 
                    try
                    {
                        Convert.ToDouble(aryline[7].ToString().Trim());
                    }
                    catch
                    {
                        if (rowNum == 0)//begin with next line
                        {
                            continue;
                        }
                        //else
                        //{
                        //    if (!isResultTableEmpty)//skip to next line
                        //    {
                        //        rowNum += 1;
                        //        continue;
                        //    }
                        //}
                    }
                    #endregion

                    if (isResultTableEmpty)
                    {
                        #region first result
                        DataRow dr = resultTable.NewRow();
                        dr["fileName"] = aryline[0].ToString();
                        dr["testCount"] = 1;
                        try
                        {
                            double initialTotalTime = Convert.ToDouble(aryline[7].ToString().Trim());
                            dr["totalTime_Max"] = initialTotalTime;
                            dr["totalTime_Min"] = initialTotalTime;
                            dr["totalTime_Avg"] = initialTotalTime;
                        }
                        catch
                        {
                            dr["totalTime_Max"] = 0;
                            dr["totalTime_Min"] = 0;
                            dr["totalTime_Avg"] = 0;
                        }
                        try
                        {
                            double initialServerTime = Convert.ToDouble(aryline[13].ToString().Trim());
                            dr["serverTime_Max"] = initialServerTime;
                            dr["serverTime_Min"] = initialServerTime;
                            dr["serverTime_Avg"] = initialServerTime;
                        }
                        catch
                        {
                            dr["serverTime_Max"] = 0;
                            dr["serverTime_Min"] = 0;
                            dr["serverTime_Avg"] = 0;
                        }

                        try
                        {
                            double initialAccuracy = Convert.ToDouble(aryline[20].ToString().Trim());
                            dr["Accuracy_Max"] = initialAccuracy;
                            dr["Accuracy_Min"] = initialAccuracy;
                            dr["Accuracy_Avg"] = initialAccuracy;
                        }
                        catch
                        {
                            dr["Accuracy_Max"] = 0;
                            dr["Accuracy_Min"] = 0;
                            dr["Accuracy_Avg"] = 0;
                        }
                        try
                        {
                            dr["expectedResult"] = aryline[21].ToString();
                            dr["bestObtainedResult"] = aryline[22].ToString();
                            dr["worstObtainedResult"] = aryline[22].ToString();
                            dr["midObtainedResult"] = aryline[22].ToString();
                        }
                        catch {
                            dr["expectedResult"] = "";
                            dr["bestObtainedResult"] = "";
                            dr["worstObtainedResult"] = "";
                            dr["midObtainedResult"] = "";
                        }
                        resultTable.Rows.Add(dr);
                        #endregion
                    }
                    else
                    {
                        int count = int.Parse(resultTable.Rows[rowNum]["testCount"].ToString().Trim());
                        resultTable.Rows[rowNum]["testCount"] = count + 1;

                        #region update TotalTime
                        try
                        {
                            //if error ,will not update
                            double newTotalTime = Convert.ToDouble(aryline[7].ToString());

                            double oldAverageTotalTime = Convert.ToDouble(resultTable.Rows[rowNum]["totalTime_Avg"]);
                            if (oldAverageTotalTime == 0)
                            {
                                oldAverageTotalTime = newTotalTime;
                            }
                            if (newTotalTime > Convert.ToDouble(resultTable.Rows[rowNum]["totalTime_Max"]))
                            {
                                resultTable.Rows[rowNum]["totalTime_Max"] = newTotalTime;
                            }
                            else if (newTotalTime < Convert.ToDouble(resultTable.Rows[rowNum]["totalTime_Min"]))
                            {
                                resultTable.Rows[rowNum]["totalTime_Min"] = newTotalTime;
                            }
                            resultTable.Rows[rowNum]["totalTime_Avg"] = (oldAverageTotalTime * count + newTotalTime) / (count + 1);
                        }
                        catch
                        {
                            //dont update this info
                        }
                        #endregion

                        #region update ServerTime
                        try
                        {
                            double newServerTime = Convert.ToDouble(aryline[13].ToString());

                            double oldAverageServerTime = Convert.ToDouble(resultTable.Rows[rowNum]["serverTime_Avg"]);
                            if (oldAverageServerTime == 0)
                            {
                                oldAverageServerTime = newServerTime;
                            }
                            if (newServerTime > Convert.ToDouble(resultTable.Rows[rowNum]["serverTime_Max"]))
                            {
                                resultTable.Rows[rowNum]["serverTime_Max"] = newServerTime;
                            }
                            else if (newServerTime < Convert.ToDouble(resultTable.Rows[rowNum]["serverTime_Min"]))
                            {
                                resultTable.Rows[rowNum]["serverTime_Min"] = newServerTime;
                            }
                            resultTable.Rows[rowNum]["serverTime_Avg"] = (oldAverageServerTime * count + newServerTime) / (count + 1);
                        }
                        catch
                        {
                            //dont update this info
                        }
                        #endregion

                        #region update Accuracy/bestObtainedResult
                        try
                        {
                            double newAccuracy = Convert.ToDouble(aryline[20].ToString());

                            double oldAverageAccuracy = Convert.ToDouble(resultTable.Rows[rowNum]["Accuracy_Avg"]);
                            if (oldAverageAccuracy == 0)
                            {
                                oldAverageAccuracy = newAccuracy;
                            }

                            if (newAccuracy > Convert.ToDouble(resultTable.Rows[rowNum]["Accuracy_Max"]))
                            {
                                resultTable.Rows[rowNum]["Accuracy_Max"] = newAccuracy;
                                //get the most accury
                                resultTable.Rows[rowNum]["midObtainedResult"] = resultTable.Rows[rowNum]["bestObtainedResult"];
                                resultTable.Rows[rowNum]["bestObtainedResult"] = aryline[22].ToString();
                            }
                            else if (newAccuracy < Convert.ToDouble(resultTable.Rows[rowNum]["Accuracy_Min"]))
                            {
                                resultTable.Rows[rowNum]["Accuracy_Min"] = newAccuracy;
                                resultTable.Rows[rowNum]["worstObtainedResult"] = aryline[22].ToString();
                            }
                            resultTable.Rows[rowNum]["Accuracy_Avg"] = (oldAverageAccuracy * count + newAccuracy) / (count + 1);
                        }
                        catch
                        {
                            //dont update this info
                        }
                        #endregion                        
                    }

                    resultTable.AcceptChanges();
                    rowNum += 1;//read next line
                }
                mysr.Close();
            }
            return resultTable;
        }



        public  string dealCsvLine(string str)
        {
            string s = "";
            int k = 1;
            if (str.Length == 0) return "";
            str = str.Replace("^", "");
            for (int i = 0; i < str.Length; i++)
            {
                switch (str.Substring(i, 1))
                {
                    case "\"":
                        s += str.Substring(i, 1);
                        k++;
                        break;
                    case ",":
                        if (k % 2 == 0)
                            s += str.Substring(i, 1);
                        else
                            s += "^";
                        break;
                    default: s += str.Substring(i, 1); break;
                }
            }
            return s;
        }

    }



}
