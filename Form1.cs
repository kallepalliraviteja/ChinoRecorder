using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Data.SQLite;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace ChinoRecorder
{
    public partial class Form1 : Form
    {
        string logFilePath = Environment.CurrentDirectory + "/log.txt";
        JObject o1;
        Config config;
        string DB_Name;
        string DB_Location;
        string ConnectionString;
        string filePath;
        public Form1()
        {
            InitializeComponent();
            o1 = JObject.Parse(File.ReadAllText(@"./config.json"));
            config = Newtonsoft.Json.JsonConvert.DeserializeObject<Config>(o1.ToString());
            FillUnitCombo();
            DB_Name = config.DB_Name;
            DB_Location = config.DB_Location;
            dtpFrom.CustomFormat = "yyyy-MMM-dd";
            dtpTo.CustomFormat = "yyyy-MMM-dd";
            ConnectionString = "Data Source=" + DB_Name;
            AquireHeatRate();
            hourTimer.Interval = MilliSecondsLeftTilTheHour();
            lblError.Text = "Next fetch in" + hourTimer.Interval / (60000) + " minutes.";
        }

        private void FillUnitCombo()
        {
            foreach (var recorder in config.Recorders)
            {
                cmbUnit.Items.Add(recorder.Name);
            }
        }
        private void hourTimer_Tick(object sender, EventArgs e)
        {
            AquireHeatRate();
            lblError.Text = "Next fetch in" + hourTimer.Interval / (60000)+" minutes.";
        }
        private void AquireHeatRate()
        {
            try
            {
                lblLast.Text = "Last updated" + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss");
                //StreamWriter sw = File.CreateText("./data.txt");
                foreach (var recorder in config.Recorders)
                {
                    string url = "http://" + recorder.RecorderIp + recorder.RecorderWebURL;
                    var user = recorder.UserName;
                    var password = recorder.Password;
                    var totalParams = recorder.NoOfParams;
                    var base64UserNameAndPassword = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{user}:{password}"));                    
                    Console.WriteLine("Fetching url and reading data");
                    string response = CallUrl(url, base64UserNameAndPassword);
                    if (response != "")
                    {
                        var linkList = ParseHtml(response, totalParams);
                        WriteToDB(linkList,recorder.Name);
                        //WriteToExcel(fileDir, linkList);
                    }
                    else
                    {
                        lblError.Text = "There is no response from the recorder.";
                    }
                }
                hourTimer.Interval = MilliSecondsLeftTilTheHour();
            }
            catch (Exception ex)
            {
                WriteToLog(ex.Message);
                lblError.Text = ex.Message;
                hourTimer.Interval = 2 * 60 * 1000;
            }

        }

        private void WriteToDB(List<string> values,string unitName)
        {
            try {
                string query = "";
                string queryvalues = "";
                for(int i = 0; i < values.Count; i++)
                {
                    queryvalues += values[i] + ",";
                }
                queryvalues +="'"+ DateTime.Now.ToString("yyyy-MM-dd HH:mm") +"'";
                //queryvalues=queryvalues.Remove(queryvalues.Length - 1, 1);
                if (unitName.ToUpper() == "UNIT5")
                    query = "insert into unit5 values("+queryvalues+");";
                if (unitName.ToUpper() == "UNIT6")
                    query = "insert into unit6 values(" + queryvalues + ");";
                var con = new SQLiteConnection(ConnectionString);
                con.Open();
                var cmd = new SQLiteCommand(query, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string CallUrl(string fullUrl, string base64UserNameAndPassword)
        {
            string responseString = "";
            try
            {
                var client = WebRequest.Create(fullUrl);
                client.Headers.Add("Authorization", "Basic " + base64UserNameAndPassword);
                var response = client.GetResponse();
                var responseStream = response.GetResponseStream();
                if (responseStream == null) return null;
                var myStreamReader = new StreamReader(responseStream, Encoding.Default);
                responseString = myStreamReader.ReadToEnd();
                return responseString;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        private List<string> ParseHtml(string html, string totalParams)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);
            List<string> paramValues = new List<string>();
            int i = 0;
            foreach (HtmlNode row in doc.DocumentNode.SelectNodes("//table//tr//td//table//tr"))
            {
                if (i == 0) { }
                else if (i != 0 && i <= int.Parse(totalParams))
                {
                    var data = row.ChildNodes[5].InnerHtml;
                    paramValues.Add(data);
                }
                else
                {
                    return paramValues;
                }
                i++;
            }
            return new List<string>();
        }
        private void WriteToLog(string message)
        {
            if (!File.Exists(logFilePath))
            {
                File.Create(logFilePath);
            }
            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter(logFilePath, true))
            {
                file.WriteLine(DateTime.Now.ToString("dd-MMM-yyyy")+" :: "+message);
            }

        }
        private void WriteToExcel(string unitName, List<RecorderParameters> values)
        {
            string fileDir = Directory.GetCurrentDirectory();
            int hour = 2;
            string fileName =unitName+"_"+ DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
            filePath = Path.Combine(fileDir, fileName);
            if (!File.Exists(filePath))
            {
                File.Copy(fileDir + ".\\template.xlsx", filePath);
            }
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            _Workbook excelBook = xlApp.Workbooks.Open(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet xlWorkSheet = (_Worksheet)excelBook.Worksheets[1];
            foreach (RecorderParameters rp in values)
            {
                xlWorkSheet.Cells[1,hour] = rp.Timestamp;
                xlWorkSheet.Cells[2, hour]  = rp.MAIN_STEAM_FLOW;
                xlWorkSheet.Cells[3, hour]  = rp.LOAD;
                xlWorkSheet.Cells[4, hour]  = rp.MAIN_STM_PR;
                xlWorkSheet.Cells[5, hour]  = rp.MAIN_STEAM_TEMP;
                xlWorkSheet.Cells[6, hour]  = rp.FW_TMP_AT_ECOIL;
                xlWorkSheet.Cells[7, hour]  = rp.FW_PR_AT_ECOIL;
                xlWorkSheet.Cells[8, hour]  = rp.HRH_PRESSURE;
                xlWorkSheet.Cells[9, hour]  = rp.HRH_TEMPERATURE;
                xlWorkSheet.Cells[10, hour] = rp.CRH_PRESSURE;
                xlWorkSheet.Cells[11, hour] = rp.CRH_TEMPERATURE;
                xlWorkSheet.Cells[12, hour] = rp.REHEAT_SPRAY;
                xlWorkSheet.Cells[13, hour] = rp.FEED_WTR_FLOW;
                xlWorkSheet.Cells[14, hour] = rp.EXT_PR_HPH6_IL;
                xlWorkSheet.Cells[15, hour] = rp.EXT_TEMP_HPH6IL;
                xlWorkSheet.Cells[16, hour] = rp.HPH6_DRIP_TEMP;
                xlWorkSheet.Cells[17, hour] = rp.BFP_DISCH_HDRPR;
                xlWorkSheet.Cells[18, hour] = rp.FW_TMP_HPH6_IL;
                xlWorkSheet.Cells[19, hour] = rp.FW_TEMP_HPH6_OL;
                xlWorkSheet.Cells[20, hour] = rp.Per_DM_MAKEUP;
                xlWorkSheet.Cells[21, hour] = rp.SH_SPRAY;
                hour++;
            }
            excelBook.Save();
            excelBook.Close(true, Type.Missing, Type.Missing);
            xlApp.Quit();
        }
        private int MilliSecondsLeftTilTheHour()
        {
            int interval;

            int minutesRemaining = 59 - DateTime.Now.Minute;
            int secondsRemaining = 59 - DateTime.Now.Second;
            interval = ((minutesRemaining * 60) + secondsRemaining) * 1000;

            // If we happen to be exactly on the hour...
            if (interval == 0)
            {
                interval = 60 * 60 * 1000;
            }
            return interval;
        }

        private void btnGenerateExcel_Click(object sender, EventArgs e)
        {
            lblExcel.Text = "Reading from database";
            if (cmbUnit.SelectedItem == null)
            {
                MessageBox.Show("Please select unit");
                return;
            }
            string fromDate = dtpFrom.Value.ToString("yyyy-MM-dd")+" 00:00";
            string toDate = dtpTo.Value.ToString("yyyy-MM-dd")+" 23:59";
            string unitName = cmbUnit.SelectedItem.ToString();
            string query = "select * from "+unitName +" where timestamp between '"+fromDate+"' and '"+toDate+"';";
            var con = new SQLiteConnection(ConnectionString);
            con.Open();
            var cmd = new SQLiteCommand(query, con);
            var data=cmd.ExecuteReader();
            List<RecorderParameters> rpList = new List<RecorderParameters>();
            lblExcel.Text = "Creating Objects";
            while (data.Read())
            {
                RecorderParameters p = new RecorderParameters();
                p.MAIN_STEAM_FLOW = float.Parse(data["MAIN_STEAM_FLOW"].ToString());
                p.LOAD = float.Parse(data["LOAD"].ToString());
                p.MAIN_STM_PR = float.Parse(data["MAIN_STM_PR"].ToString());
                p.MAIN_STEAM_TEMP = float.Parse(data["MAIN_STEAM_TEMP"].ToString());
                p.FW_PR_AT_ECOIL = float.Parse(data["FW_PR_AT_ECOIL"].ToString());
                p.FW_TMP_AT_ECOIL = float.Parse(data["FW_TMP_AT_ECOIL"].ToString());
                p.HRH_PRESSURE = float.Parse(data["HRH_PRESSURE"].ToString());
                p.HRH_TEMPERATURE = float.Parse(data["HRH_TEMPERATURE"].ToString());
                p.CRH_PRESSURE = float.Parse(data["CRH_PRESSURE"].ToString());
                p.CRH_TEMPERATURE = float.Parse(data["CRH_TEMPERATURE"].ToString());
                p.REHEAT_SPRAY = float.Parse(data["REHEAT_SPRAY"].ToString());
                p.FEED_WTR_FLOW = float.Parse(data["FEED_WTR_FLOW"].ToString());
                p.EXT_PR_HPH6_IL = float.Parse(data["EXT_PR_HPH6_IL"].ToString());
                p.EXT_TEMP_HPH6IL = float.Parse(data["EXT_TEMP_HPH6IL"].ToString());
                p.HPH6_DRIP_TEMP = float.Parse(data["HPH6_DRIP_TEMP"].ToString());
                p.BFP_DISCH_HDRPR = float.Parse(data["BFP_DISCH_HDRPR"].ToString());
                p.FW_TMP_HPH6_IL = float.Parse(data["FW_TMP_HPH6_IL"].ToString());
                p.FW_TEMP_HPH6_OL = float.Parse(data["FW_TEMP_HPH6_OL"].ToString());
                p.Per_DM_MAKEUP = float.Parse(data["Per_DM_MAKEUP"].ToString());
                p.SH_SPRAY = float.Parse(data["SH_SPRAY"].ToString());
                p.Timestamp = Convert.ToDateTime(data["Timestamp"].ToString());
                rpList.Add(p);
            }
            con.Close();
            lblExcel.Text = "Writing to excel...";
            WriteToExcel(unitName,rpList);
            lblExcel.Text="File location "+ filePath;
        }
    }
}
