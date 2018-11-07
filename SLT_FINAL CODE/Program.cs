//using AutoIt;
using AutoItX3Lib;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
//using System.Data.OracleClient;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using Teradata.Client.Provider;
using System.Data.Odbc;


namespace SLT_Final_code
{
    class Program
    {
        static string sqlconnstring = "Server=ITSUSRAWSP03267; Database=ASFRA-DEV; User Id=ASFRA_UI_User; password= Jnj2013@";
        public static DataTable GetAppids()
        {
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {

                con.Open();

                SqlCommand getappids = new SqlCommand("select appid as APPID from [CONN_DETAILS] order by appid", con);
                SqlDataAdapter da = new SqlDataAdapter(getappids);
                da.Fill(dt);
                con.Close();

            }
            return dt;
        }
        public static IWebDriver Firefox_login(int appid, DataTable dt, IWebDriver driver, int prodDay)
        {
            String[] Array_Excel = new String[5];
            String[] Values = new String[20];

            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {
                con.Open();
                SqlCommand getappids = new SqlCommand("select [NTID], CAST([PASSWORD] AS VARCHAR(MAX)) as [Password], [AGENTS], [ENVIRONMENT],[URL] from [dbo].[CONN_DETAILS] where APPID=" + appid.ToString(), con);
                SqlDataReader dr = getappids.ExecuteReader();
                while (dr.Read())
                {
                    Array_Excel[0] = dr["NTID"].ToString();
                    Array_Excel[1] = dr["Password"].ToString();
                    Array_Excel[2] = dr["AGENTS"].ToString();
                    Array_Excel[3] = dr["ENVIRONMENT"].ToString();
                    Array_Excel[4] = dr["URL"].ToString();
                }
                con.Close();
            }

            if (Array_Excel[2] != null)
            {
                Values = Array_Excel[2].Split(',');
                for (int i = 0; i < Values.Length; i++)
                {
                    Values[i] = Values[i].Trim();
                }

            }

            string URL = "http://" + Array_Excel[4];


            driver = null;
            driver = new OpenQA.Selenium.Firefox.FirefoxDriver();
            driver.Navigate().GoToUrl(URL);
            Task.Delay(10000).Wait();
            AutoItX3 autoIT = new AutoItX3();

            autoIT.Send(Array_Excel[0]);

            Task.Delay(2000).Wait();

            autoIT.Send("{TAB}");

            Task.Delay(2000).Wait();

            autoIT.Send(Array_Excel[1]);

            Task.Delay(2000).Wait();

            autoIT.Send("{ENTER}");

            Task.Delay(15000).Wait();


            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(200));
            wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_RblEnvironments_0")));
            if ((Array_Excel[3] == "PROD"))
            {
                driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_RblEnvironments_0")).Click();
            }
            else
            {
                driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_RblEnvironments_1")).Click();
            }
            var Textbox1 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox2"));
            DateTime conv;
            //double d = double.Parse(Array_Excel[2]);
            if (appid == 5)
            {
                conv = DateTime.Now.AddDays(-5);
            }

            else
            {
                if (appid == 8)
                {
                    conv = DateTime.Now.AddDays(-2);
                }
                else
                {
                    conv = DateTime.Now.AddDays(-1);
                }

            }
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

            Textbox1.Clear();
            string date = conv.ToString("MM/dd/yyyy");
            Task.Delay(2000).Wait();
            try
            {

                driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox2")).SendKeys(date);
            }
            catch (Exception ex)
            {
                var ell = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox2"));
                js.ExecuteScript("document.getElementById('ctl00_ContentPlaceHolder1_TextBox2').setAttribute('value', '" + date + "')");

            }
            var Textbox2 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox6"));



            js.ExecuteScript("arguments[0].click();", Textbox2);


            var Textbox2_inner = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ListBox7"));

            for (int i = 0; i < Values.Length; i++)
            {
                Console.WriteLine(Values[i]);
                string Xpath_element = "//*[@id='ctl00_ContentPlaceHolder1_ListBox7']/option[contains(@value,'" + Values[i].Trim().ToString() + "')]";
                IWebElement agent = driver.FindElement(By.XPath(Xpath_element));
                agent.Click();
                Task.Delay(2000).Wait();
            }


            //div1.Click();
            //div2.Click();

            wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton7")));

            var Textbox_1 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton7"));

            js.ExecuteScript("arguments[0].click();", Textbox_1);

            //driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(200));

            Task.Delay(10000).Wait();

            //Third Textbox clicking

            var Textbox3 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox3"));

            js.ExecuteScript("arguments[0].click();", Textbox3);

            //wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton14")));

            var Textbox3_selectall = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton14"));

            js.ExecuteScript("arguments[0].click();", Textbox3_selectall);

            Task.Delay(10000).Wait();

            //wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton9")));

            var Textbox3_apply = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton9"));

            js.ExecuteScript("arguments[0].click();", Textbox3_apply);

            Task.Delay(10000).Wait();

            //Fourth Textbox clicking

            var Textbox4 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox5"));

            js.ExecuteScript("arguments[0].click();", Textbox4);

            //wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton14")));

            var Textbox4_selectall = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton15"));

            js.ExecuteScript("arguments[0].click();", Textbox4_selectall);

            Task.Delay(10000).Wait();

            //wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton9")));

            var Textbox4_apply = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton11"));

            js.ExecuteScript("arguments[0].click();", Textbox4_apply);

            Task.Delay(20000).Wait();



            // clicking Refresh data

            var button1 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_Button1"));

            js.ExecuteScript("arguments[0].click();", button1);

            Task.Delay(50000).Wait();

            return driver;
        }
        static void Main(string[] args)
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("JOBRUN_ID");
                dt.Columns["JOBRUN_ID"].AllowDBNull = true;
                dt.Columns.Add("PRODUCTION_DATE");
                dt.Columns["PRODUCTION_DATE"].DataType = typeof(DateTime);
                dt.Columns["PRODUCTION_DATE"].AllowDBNull = true;
                dt.Columns.Add("JOBGROUP");
                dt.Columns["JOBGROUP"].AllowDBNull = true;
                dt.Columns.Add("JOB");
                dt.Columns["JOB"].AllowDBNull = true;
                dt.Columns.Add("AGENT");
                dt.Columns["AGENT"].AllowDBNull = true;
                dt.Columns.Add("STATUS");
                dt.Columns["STATUS"].AllowDBNull = true;
                dt.Columns.Add("INSTANCE");
                dt.Columns["INSTANCE"].AllowDBNull = true;
                dt.Columns["INSTANCE"].DataType = typeof(Int32);
                dt.Columns.Add("RERUNS");
                dt.Columns["RERUNS"].AllowDBNull = true;
                dt.Columns["RERUNS"].DataType = typeof(Int32);
                dt.Columns.Add("COMMAND");
                dt.Columns["COMMAND"].AllowDBNull = true;
                dt.Columns.Add("PARAMS");
                dt.Columns["PARAMS"].AllowDBNull = true;
                dt.Columns.Add("ESTIMATEDSTARTTIME");
                dt.Columns["ESTIMATEDSTARTTIME"].AllowDBNull = true;
                dt.Columns["ESTIMATEDSTARTTIME"].DataType = typeof(DateTime);
                dt.Columns.Add("ESTIMATEDDURATION");
                dt.Columns["ESTIMATEDDURATION"].AllowDBNull = true;
                dt.Columns.Add("ACTUALSTARTTIME");
                dt.Columns["ACTUALSTARTTIME"].DataType = typeof(DateTime);
                dt.Columns["ACTUALSTARTTIME"].AllowDBNull = true;
                dt.Columns.Add("ACTUALDURATION");
                dt.Columns["ACTUALDURATION"].AllowDBNull = true;
                dt.Columns.Add("JOBID");
                dt.Columns["JOBID"].AllowDBNull = true;
                dt.Columns["JOBID"].DataType = typeof(Int32);
                dt.Columns.Add("DEPENDENCIES");
                dt.Columns["DEPENDENCIES"].AllowDBNull = true;
                dt.Columns.Add("APPID");
                dt.Columns["APPID"].DataType = typeof(Int32);
                dt.Columns["APPID"].AllowDBNull = true;
                dt.Columns.Add("SRC_DATA");
                dt.Columns["SRC_DATA"].AllowDBNull = true;
                int appid = Convert.ToInt32(args[0]);
              //  int appid = 14;
                string[] differentDaysProdRun = new string[5];
                using (SqlConnection con = new SqlConnection(sqlconnstring))
                {
                    con.Open();
                    string deletetidal = "delete SLT_DATA where   appid=" + appid + ";";
                    SqlCommand deletecmd = new SqlCommand(deletetidal, con);
                    deletecmd.ExecuteNonQuery();

                    string getDifferentrunday = "select distinct prod_day from slt_app_configuration where appid=" + appid;
                    SqlCommand getDifferentrundayCommand = new SqlCommand(getDifferentrunday, con);
                    SqlDataReader dr = getDifferentrundayCommand.ExecuteReader();

                    int i = 0;
                    while (dr.Read())
                    {
                        differentDaysProdRun[i] = dr["prod_day"].ToString();
                        i++;
                    }
                    con.Close();
                }
                if (appid == 3)
                {
                    dt.Clear();
                    CognosReport(appid, dt);
                }
                else
                {
                    if (appid == 11 || appid == 12)
                    {
                        dt.Clear();
                        QlikSense(appid, dt);

                    }
                    else
                    {
                        if (appid == 13)
                        {
                            GlobalSales(appid, dt);
                        }
                        else
                        {
                            dt.Clear();
                            // consumeredgetableaureport(dt, datarow["appid"].ToString());
                            // Cognosreport("3", dt);
                            // sendmail(appid);
                            foreach (var prodday in differentDaysProdRun)
                            {

                                if (prodday != null)
                                {
                                    TidalNavigationNavigate(appid, dt, Convert.ToInt32(prodday));
                                }
                            }
                        }
                        // consumeredgetableaureport(dt,i.ToString());
                    }
                }

                SendMail(appid);


            }
            catch (Exception ex)
            {
                string filePath = "Error.txt";
                using (StreamWriter writer = new StreamWriter(filePath, true))
                {
                    writer.WriteLine("Message :" + ex.Message + "<br/>" + Environment.NewLine + "StackTrace :" + ex.StackTrace +
                       "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());

                    writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);

                }
            }

        }
        public static void CognosReport(int appid, DataTable dt)
        {
            String[] Array_Excel = new String[5];
            String[] Values = new String[10];
            string[] joblastname = new string[10];
            string[] jobfirstname = new string[10];
            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {
                con.Open();
                SqlCommand getappids = new SqlCommand("select [NTID], CAST([PASSWORD] AS VARCHAR(MAX)) as [Password], [AGENTS], [ENVIRONMENT],[URL] from [dbo].[CONN_DETAILS] where APPID=" + appid + "and Source ='Cognos'", con);
                SqlDataReader dr = getappids.ExecuteReader();
                while (dr.Read())
                {
                    Array_Excel[0] = dr["NTID"].ToString();
                    Array_Excel[1] = dr["Password"].ToString();
                    Array_Excel[2] = dr["AGENTS"].ToString();
                    Array_Excel[3] = dr["ENVIRONMENT"].ToString();
                    Array_Excel[4] = dr["URL"].ToString();

                }
                con.Close();
                con.Open();
                SqlCommand getjobanme = new SqlCommand("SELECT distinct START_JOB from SLT_APP_CONFIGURATION where appid=" + appid, con);
                SqlDataReader jobnames = getjobanme.ExecuteReader();
                int i = 0;
                while (jobnames.Read())
                {
                    joblastname[i] = jobnames["START_JOB"].ToString();
                    i++;

                }
                con.Close();
                con.Open();
                SqlCommand getjobgroupname = new SqlCommand("SELECT distinct START_JOB_GROUP from SLT_APP_CONFIGURATION where appid=" + appid, con);
                SqlDataReader jobgroupnames = getjobgroupname.ExecuteReader();
                i = 0;
                while (jobgroupnames.Read())
                {
                    jobfirstname[i] = jobgroupnames["START_JOB_GROUP"].ToString();
                    i++;

                }

                con.Close();

            }

            OpenQA.Selenium.IWebDriver driver = null;
            driver = new OpenQA.Selenium.Firefox.FirefoxDriver();

            //string URL = "https://sbhamid3:June@2017@nacon.cogbi.jnj.com/1022/cgi-bin/cognosisapi.dll?b_action=xts.run&m=portal/cc.xts&m_folder=i9245D7C638894F6799F285DB45298167&refresh=";
            string URL = "https://" + Array_Excel[4];
            //  OpenQA.Selenium.Firefox.FirefoxProfile profile=null;
            // profile.SetPreference("network.automatic-ntlm-auth.trusted-uris", Array_Excel[4]);

            driver.Navigate().GoToUrl(URL);
            Task.Delay(1000);
           // IAlert alert = driver.SwitchTo().Alert();

            Task.Delay(5000).Wait();

            AutoItX3 autoIT = new AutoItX3();

            autoIT.Send(Array_Excel[0]);

            Task.Delay(2000).Wait();

            autoIT.Send("{TAB}");

            autoIT.Send(Array_Excel[1]);

            Task.Delay(2000).Wait();

            autoIT.Send("{ENTER}");

            Task.Delay(10000).Wait();
            //IWebElement modified = driver.FindElement(By.Id("columnHeaderModified"));

           // driver.FindElement(By.LinkText("Modified")).Click();
           // Task.Delay(5000).Wait();
           // driver.FindElement(By.LinkText("Modified")).Click();
           // Task.Delay(10000).Wait();
            //modified.Click();

            //joblastname[0] = "SYSTAGENIX";
            //joblastname[1] = "SYSCO";
            //joblastname[2] = "Perrigo";
            //joblastname[3] = "PRECISE PATCH";
            //joblastname[4] = "APOTEX";
            //joblastname[5] = "BIONPHARMA";
            //joblastname[6] = "CARIBBEAN";

            //jobfirstname[0] = "AE";
            //jobfirstname[1] = "PQC";
            string reportfirstname = "";
            string reportlastname = "";
            DateTime today = DateTime.Today;

            IWebElement table = driver.FindElement(By.XPath("//table[@summary='Table display of IBM Cognos content']/tbody"));
            List<IWebElement> tablerow = table.FindElements(By.TagName("tr")).ToList();
            // ReadOnlyCollection<IWebElement> lst = driver.FindElements(By.XPath("//a[contains(text(),'More')]"));
            for (int i = 1; i < 19; i++)
            {
                reportfirstname = "";
                reportlastname = "";
                var name = driver.FindElement(By.XPath("//table[@summary='Table display of IBM Cognos content']/tbody/tr[" + i + "]/td[5]")).Text;
                for (int j = 0; j < jobfirstname.Length; j++)
                {
                    if (jobfirstname[j] != null)
                    {
                        if (name.ToLower().Contains(jobfirstname[j].ToLower()))
                        {
                            reportfirstname = jobfirstname[j];
                        }
                    }
                }
                for (int j = 0; j < joblastname.Length; j++)
                {
                    if (joblastname[j] != null)
                    {
                        if (name.ToLower().Contains(joblastname[j].ToLower()))
                        {
                            reportlastname = joblastname[j];
                        }
                    }
                }
                if (reportfirstname != "" & reportlastname != "")
                {
                    driver.FindElement(By.XPath("//table[@summary='Table display of IBM Cognos content']/tbody/tr[" + i + "]/td/descendant::td[5]")).Click();

                    Task.Delay(5000).Wait();
                    var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(200));
                    wait.Until(drv => drv.FindElement(By.XPath("//a[contains(text(),'View run history')]")));

                    driver.FindElement(By.XPath("//a[contains(text(),'View run history')]")).Click();

                    Task.Delay(5000).Wait();
                    string starttimestamp;
                    string endtimestamp;
                    string status;
                    // 2017-05-29 05:39:29.000
                    //   ‪May 30, 2017 7:02:00 AM


                    starttimestamp = driver.FindElement(By.XPath("//table[@summary='Content']/tbody/tr/td[3]")).Text;
                    endtimestamp = driver.FindElement(By.XPath("//table[@summary='Content']/tbody/tr/td[5]")).Text;
                    status = driver.FindElement(By.XPath("//table[@summary='Content']/tbody/tr/td[7]")).Text;
                    driver.FindElement(By.XPath("//img[@title='Close']")).Click();
                    Task.Delay(10000).Wait();
                    driver.FindElement(By.XPath("//img[@title='Close']")).Click();
                    Task.Delay(10000).Wait();
                    string format = "MMMM d, yyyy";
                    int startdatelength = starttimestamp.Length - 12;
                    string str = starttimestamp.Substring(0, 1);
                    string datet = starttimestamp.Substring(1, startdatelength);
                    DateTime startdate = DateTime.ParseExact(datet.Trim(), format, new CultureInfo("en-US"));
                    DateTime enddate = DateTime.ParseExact(endtimestamp.Substring(1, endtimestamp.Length - 12).Trim(), format, new CultureInfo("en-US"));
                    string starttime = starttimestamp.Substring(starttimestamp.Length - 12, 11);
                    //string pm = starttime.Substring(starttime.Length - 3, 2);
                    //if (starttime.Substring(starttime.Length - 3, 2) == "AM")
                    //{
                    //    int hour;
                    //    if (starttime.Trim().Substring(1, 1) == ":")
                    //    {
                    //        hour = Convert.ToInt32(starttime.Trim().Substring(0, 1));
                    //        hour = hour + 12;
                    //        if (hour > 24)
                    //        {
                    //            hour = hour - 24;
                    //            starttime = hour + starttime.Trim().Substring(1, starttime.Length - 4).Trim();
                    //        }

                    //        else
                    //        {
                    //            hour = Convert.ToInt32(starttime.Trim().Substring(0, 1));

                    //            starttime = hour + starttime.Trim().Substring(1, starttime.Length - 4).Trim();

                    //        }

                    //    }
                    //    else
                    //    {
                    //        hour = Convert.ToInt32(starttime.Trim().Substring(0, 2));
                    //        hour = hour + 12;
                    //        if (hour > 24)
                    //        {
                    //            hour = hour - 24;
                    //            starttime = hour + starttime.Trim().Substring(2, starttime.Length - 3).Trim();
                    //        }

                    //        else
                    //        {
                    //            hour = Convert.ToInt32(starttime.Trim().Substring(0, 2));

                    //            starttime = hour + starttime.Trim().Substring(2, starttime.Length - 3).Trim();

                    //        }
                    //    }


                    //}
                    //else
                    //{

                    //    starttime = starttime.Trim().Substring(0, starttime.Length - 3);
                    //}
                    starttimestamp = startdate.ToString("yyyy-MM-dd").Trim() + " " + starttime.Trim();
                    //  startdate = DateTime.ParseExact(starttimestamp, "yyyy-mm-dd hh:mm:ss", new CultureInfo("en-US"));
                    startdate = Convert.ToDateTime(starttimestamp);
                    string endtime = endtimestamp.Substring(endtimestamp.Length - 12, 11);
                    //string am = endtime.Substring(endtime.Length - 3, 2);
                    //if (endtime.Substring(endtime.Length - 3, 2) == "AM")
                    //{
                    //    int hour;
                    //    if (endtime.Trim().Substring(1, 1) == ":")
                    //    {
                    //        hour = Convert.ToInt32(endtime.Trim().Substring(0, 1));
                    //        hour = hour + 12;
                    //        if (hour > 24)
                    //        {
                    //            hour = hour - 24;
                    //            endtime = hour + endtime.Trim().Substring(1, endtime.Length - 4).Trim();
                    //        }

                    //        else
                    //        {
                    //            hour = Convert.ToInt32(endtime.Trim().Substring(0, 1));

                    //            endtime = hour + endtime.Trim().Substring(1, endtime.Length - 4).Trim();

                    //        }

                    //    }
                    //    else
                    //    {
                    //        hour = Convert.ToInt32(endtime.Trim().Substring(0, 2));
                    //        hour = hour + 12;
                    //        if (hour > 24)
                    //        {
                    //            hour = hour - 24;
                    //            endtime = hour + endtime.Trim().Substring(2, endtime.Length - 3).Trim();
                    //        }

                    //        else
                    //        {
                    //            hour = Convert.ToInt32(endtime.Trim().Substring(0, 2));

                    //            endtime = hour + endtime.Trim().Substring(2, endtime.Length - 3).Trim();

                    //        }
                    //    }


                    //}
                    //else
                    //{

                    //    endtime = endtime.Trim().Substring(0, endtime.Length - 3);
                    //}
                    endtimestamp = enddate.ToString("yyyy-MM-dd") + " " + endtime.Trim();

                    enddate = Convert.ToDateTime(endtimestamp);
                    var Todays_Day = DateTime.Now.ToString("dd");

                    //if (Todays_Day == startdate.Day.ToString())
                    // {
                    TimeSpan ts = (enddate - startdate);
                    int days = ts.Days;
                    int hour = ts.Hours;
                    int min = ts.Minutes;
                    int sec = ts.Seconds;
                    DataRow dr = dt.NewRow();
                    dr["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                    dr["SRC_DATA"] = "Cognos report Website";
                    dr["STATUS"] = status;
                    dr["APPID"] = appid;
                    dr["ACTUALDURATION"] = days + ":" + hour.ToString("00") + ":" + min.ToString("00") + ":" + sec.ToString("00");
                    dr["ACTUALSTARTTIME"] = Convert.ToDateTime(startdate);
                    dr["JOB"] = reportlastname;
                    dr["JOBGROUP"] = reportfirstname;
                    dt.Rows.Add(dr);
                    //  }
                }

            }
            driver.Quit();

            BulkUpload(dt);







        }
        public static void BulkUpload(DataTable bulkuploaddatatable)
        {
            string sqlconnstring = "Server=ITSUSRAWSP03267; Database=ASFRA-DEV; User Id=ASFRA_UI_User; password= Jnj2013@";
            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {
                con.Open();

                using (SqlBulkCopy sqlbulk = new SqlBulkCopy(sqlconnstring, SqlBulkCopyOptions.FireTriggers))
                {
                    sqlbulk.DestinationTableName = "dbo.SLT_DATA";
                    sqlbulk.WriteToServer(bulkuploaddatatable);
                }

                SqlCommand cmd = new SqlCommand("dbo.spUpdateEndTime", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.ExecuteNonQuery();
                SqlCommand cmd1 = new SqlCommand("dbo.spUpdateEndTime", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.ExecuteNonQuery();

            }

        }
        public static void TidalNavigationNavigate(int appid, DataTable dt, int prodDay)
        {

            dt.Clear();
            String[] Array_Excel = new String[5];
            String[] Values = new String[20];

            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {
                con.Open();


                SqlCommand getappids = new SqlCommand("select [NTID], CAST([PASSWORD] AS VARCHAR(MAX)) as [Password], [AGENTS], [ENVIRONMENT],[URL] from [dbo].[CONN_DETAILS] where APPID=" + appid.ToString() + "and Source ='Tidal'", con);
                SqlDataReader dr = getappids.ExecuteReader();
                while (dr.Read())
                {
                    Array_Excel[0] = dr["NTID"].ToString();
                    Array_Excel[1] = dr["Password"].ToString();
                    Array_Excel[2] = dr["AGENTS"].ToString();
                    Array_Excel[3] = dr["ENVIRONMENT"].ToString();
                    Array_Excel[4] = dr["URL"].ToString();
                }
                con.Close();
            }

            if (Array_Excel[2] != null)
            {
                Values = Array_Excel[2].Split(',');
                for (int i = 0; i < Values.Length; i++)
                {
                    Values[i] = Values[i].Trim();
                }

            }
            OpenQA.Selenium.IWebDriver driver = null;

            DateTime currentTime = DateTime.Now;
            if (appid == 14)
            {
                 currentTime = TimeInSTZ();
            }
            int currentDay = currentTime.Day;
            DateTime productionDate  = currentTime.AddDays(prodDay);            
            int SystemDay = DateTime.Now.Day;          
            try
            {

                string URL = "http://" + Array_Excel[0] + ":" + Array_Excel[1] + "@" + Array_Excel[4];



                driver = new OpenQA.Selenium.Chrome.ChromeDriver();
                driver.Navigate().GoToUrl(URL);
                Task.Delay(10000).Wait();
                AutoItX3 autoIT = new AutoItX3();
                autoIT.Send("{F6}");
                Task.Delay(1000).Wait();

                autoIT.Send("{ENTER}");
                Task.Delay(20000).Wait();
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(200));
                wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_RblEnvironments_0")));
                if ((Array_Excel[3] == "PROD"))
                {
                    driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_RblEnvironments_0")).Click();
                }
                else
                {
                    driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_RblEnvironments_1")).Click();
                }
                var Textbox1 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox2"));
               
               
                //double d = double.Parse(Array_Excel[2]);
                if (appid == 5)
                {
                    productionDate = DateTime.Now.AddDays(-4);

                    if (DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
                    {
                        productionDate = DateTime.Now.AddDays(-5);
                    }
                    if (DateTime.Now.DayOfWeek == DayOfWeek.Sunday)
                    {
                        productionDate = DateTime.Now.AddDays(-6);
                    }
                    if (DateTime.Now.DayOfWeek == DayOfWeek.Monday)
                    {
                        productionDate = DateTime.Now.AddDays(-7);
                    }
                    if (DateTime.Now.DayOfWeek == DayOfWeek.Tuesday)
                    {
                        productionDate = DateTime.Now.AddDays(-1);
                    }
                    if (DateTime.Now.DayOfWeek == DayOfWeek.Wednesday)
                    {
                        productionDate = DateTime.Now.AddDays(-2);
                    }
                    if (DateTime.Now.DayOfWeek == DayOfWeek.Thursday)
                    {
                        productionDate = DateTime.Now.AddDays(-3);
                    }
                    if (DateTime.Now.DayOfWeek == DayOfWeek.Friday)
                    {
                        productionDate = DateTime.Now.AddDays(-4);
                    }
                }

                else
                {
                    if (appid == 8)
                    {


                        if (DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
                        {
                            productionDate = DateTime.Now.AddDays(-1);
                        }
                        if (DateTime.Now.DayOfWeek == DayOfWeek.Sunday)
                        {
                            productionDate = DateTime.Now.AddDays(-2);
                        }
                        if (DateTime.Now.DayOfWeek == DayOfWeek.Monday)
                        {
                            productionDate = DateTime.Now.AddDays(-3);
                        }
                        if (DateTime.Now.DayOfWeek == DayOfWeek.Tuesday)
                        {
                            productionDate = DateTime.Now.AddDays(-4);
                        }
                        if (DateTime.Now.DayOfWeek == DayOfWeek.Wednesday)
                        {
                            productionDate = DateTime.Now.AddDays(-5);
                        }
                        if (DateTime.Now.DayOfWeek == DayOfWeek.Thursday)
                        {
                            productionDate = DateTime.Now.AddDays(-6);
                        }
                        if (DateTime.Now.DayOfWeek == DayOfWeek.Friday)
                        {
                            productionDate = DateTime.Now.AddDays(-7);
                        }

                    }
                  


                }
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

                Textbox1.Clear();
                string date = productionDate.ToString("MM/dd/yyyy");
                Task.Delay(2000).Wait();


                driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox2")).SendKeys(date);


                var Textbox2 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox6"));



                js.ExecuteScript("arguments[0].click();", Textbox2);


                var Textbox2_inner = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_ListBox7"));

                for (int i = 0; i < Values.Length; i++)
                {
                    Console.WriteLine(Values[i]);
                    string Xpath_element = "//*[@id='ctl00_ContentPlaceHolder1_ListBox7']/option[contains(@value,'" + Values[i].Trim().ToString() + "')]";
                    IWebElement agent = driver.FindElement(By.XPath(Xpath_element));
                    agent.Click();
                    Task.Delay(2000).Wait();
                }


                //div1.Click();
                //div2.Click();

                wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton7")));

                var Textbox_1 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton7"));

                js.ExecuteScript("arguments[0].click();", Textbox_1);

                //driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(200));

                Task.Delay(10000).Wait();

                //Third Textbox clicking

                var Textbox3 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox3"));

                js.ExecuteScript("arguments[0].click();", Textbox3);

                //wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton14")));

                var Textbox3_selectall = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton14"));

                js.ExecuteScript("arguments[0].click();", Textbox3_selectall);

                Task.Delay(10000).Wait();

                //wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton9")));

                var Textbox3_apply = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton9"));

                js.ExecuteScript("arguments[0].click();", Textbox3_apply);

                Task.Delay(10000).Wait();

                //Fourth Textbox clicking

                var Textbox4 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_TextBox5"));

                js.ExecuteScript("arguments[0].click();", Textbox4);

                //wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton14")));

                var Textbox4_selectall = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton15"));

                js.ExecuteScript("arguments[0].click();", Textbox4_selectall);

                Task.Delay(10000).Wait();

                //wait.Until(drv => drv.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton9")));

                var Textbox4_apply = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_LinkButton11"));

                js.ExecuteScript("arguments[0].click();", Textbox4_apply);

                Task.Delay(20000).Wait();



                // clicking Refresh data

                var button1 = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_Button1"));

                js.ExecuteScript("arguments[0].click();", button1);

                Task.Delay(50000).Wait();
            }
            catch (Exception ex)
            {
                driver.Close();

                Task.Delay(5000).Wait();

                driver = Firefox_login(appid, dt, driver, prodDay);

            }

            bool tableexist = false;
            try
            {
                driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_GridView2"));
                tableexist = true;

            }
            catch (NoSuchElementException)
            {
                tableexist = false;
            }
            if (tableexist == true)
            {
                IWebElement table = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_GridView2"));

                List<IWebElement> tablerow = table.FindElements(By.TagName("tr")).ToList();

                DataRow datarow;
                foreach (var tr in tablerow)
                {
                    bool validrecord = true;
                    datarow = dt.NewRow();
                    int col = 0;
                    int i = 0;
                    List<IWebElement> td = tr.FindElements(By.TagName("td")).ToList();
                    if (td.Count > 0)
                    {

                        foreach (IWebElement tds in td)
                        {

                            if (i > 1 && i < 18)
                            {
                                if (i == 3 || i == 4 || i == 5 || i == 7 || i == 14 || i == 15)
                                {
                                    if (tds.TagName == "td")
                                    {

                                        string text = tds.Text.ToString().Trim();
                                        string proddate = productionDate.ToString("MM/dd/yyyy");
                                        if (i == 3 & proddate != tds.Text)
                                        {
                                            validrecord = false;
                                            break;
                                        }

                                        if (tds.Text.ToString().Trim() != "" & tds.Text.ToString().Trim() != null)
                                        {
                                            datarow[col] = tds.Text;
                                        }

                                        else
                                        {
                                            datarow[col] = DBNull.Value;
                                        }

                                    }

                                }
                                col++;
                            }
                            i++;


                        }
                        if (col != 0)
                        {
                            datarow["SRC_DATA"] = "TIDAL";
                            datarow["APPID"] = appid;
                        }
                        if (validrecord)

                            dt.Rows.Add(datarow);

                    }



                }

            }


            if (appid == 2)
            {


                DataTable cloneddt = dt.Clone();

                DataTable dt4 = ConsumerEdgeTableauReport(driver, appid, cloneddt);
                foreach (DataRow row in dt4.Rows)
                {
                    dt.ImportRow(row);
                }
                cloneddt.Clear();
                DataTable dt1 = ConsumerEdgeCubeRefresh(cloneddt,appid);
                foreach (DataRow row in dt1.Rows)
                {
                    dt.ImportRow(row);
                }
                cloneddt.Clear();
                DataTable dt2 = ConsumerEdgeReportDelievery(cloneddt, appid);
                foreach (DataRow row in dt2.Rows)
                {
                    dt.ImportRow(row);
                }
                cloneddt.Clear();
                DataTable dt3 = ConsumerEdgeTableauReport(cloneddt, appid);
                foreach (DataRow row in dt3.Rows)
                {
                    dt.ImportRow(row);
                }

            }
            BulkUpload(dt);
            driver.Quit();
            if (appid == 8)
            {
                dt.Clear();
                QlikSense(appid, dt);
            }


        }
        static void UpdateData(int appid)
        {
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {
                con.Open();

                SqlCommand storeprocedure = new SqlCommand("sp_InsertHistSltData", con);
                storeprocedure.CommandType = CommandType.StoredProcedure;
                storeprocedure.Parameters.Add("@Appid", SqlDbType.Int).Value = appid;
                storeprocedure.ExecuteNonQuery();
                con.Close();
            }

        }
        static void SendMail(int appid)
        {



            DataTable dt = new DataTable();
            DataTable maildetails = new DataTable();
            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {

                SqlCommand getrecordcount = new SqlCommand("select count(*) from sLT_app_configuration where appid=" + appid, con);
                con.Open();
                // string deletehist = "delete HIST_TIDAL_SLT_DATA where cast(SLT_LOADDATE_TIME as date)=cast(getdate() as date) and appid=" + appid + ";";
                // SqlCommand deletecmd = new SqlCommand(deletehist, con);
                int recordcount = Convert.ToInt32(getrecordcount.ExecuteScalar());
                //deletecmd.ExecuteNonQuery();


                UpdateData(appid);
                SqlCommand getappids = new SqlCommand("select top " + recordcount + " *  from HIST_SLT_DATA WHERE APPID=" + appid + " AND cast(SLT_LOADDATE_TIME as date)='" + DateTime.Today.Date + "' order by row_id desc", con);
                SqlDataAdapter da = new SqlDataAdapter(getappids);
                da.Fill(dt);


                SqlCommand sltreportdata = new SqlCommand("select *  from SLT_APP_MASTER WHERE APPID=" + appid, con);
                SqlDataAdapter reportdataadapter = new SqlDataAdapter(sltreportdata);
                reportdataadapter.Fill(maildetails);
                con.Close();
            }

         
            string htmlbody = "";
            htmlbody += maildetails.Rows[0]["Content"].ToString();
            int SN = 1;
          
            if (appid == 9 || appid == 10)
            {
               
                foreach (DataRow dr in dt.Rows)
                {
                    string statusstyle = "";
                    string status = "";
                    if (Convert.ToString(dr["SLT_STATUS"]).ToLower().Trim() == "met")
                    {

                        statusstyle = "style='background-color:#00ff00;'";
                        status = dr["STATUS"].ToString();
                    }

                    if (Convert.ToString(dr["SLT_STATUS"]).ToLower().Trim() == "not met")
                    {
                        statusstyle = "style='background-color:red;'";
                        status = dr["STATUS"].ToString();
                    }
                    if (Convert.ToString(dr["SLT_STATUS"]).ToLower().Trim() == "wip")
                    {
                        status = "WIP";
                        statusstyle = "style='background-color:yellow;'";
                    }

                    htmlbody += "<tr><td>" + SN + "</td><td>" +
                           dr["BUSINESS_DELIVERABLE"] + "</td><td>" +
                           dr["FREQUENCY"] + "</td><td>" +
                           dr["SOURCE_NAME"] + "</td><td>" +
                           dr["CATEGORY"] + "</td><td " + statusstyle + ">" +
                           status + "<td>" +
                          dr["ESTIMATEDSTARTTIME"] + "</td><td>" +
                          dr["ESTIMATEDENDTIME"] + "</td><td>" +
                            dr["ACTUALSTARTTIME"] + "</td><td>" +
                           dr["ACTUALENDTIME"] + "</td>" +
                           "<td></td><td></td></tr>";
                    SN = SN + 1;
                }
            }
            else
            {

                string deliverable = "business deliverable";
                if (!htmlbody.ToLower().Trim().Contains(deliverable.ToLower().Trim()))
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string img = "<img  src='http://asfra-dev.jnj.com/ASFRA/Content/Circle/Red.png' width:20pt;  alt='SLT MET' />";

                        if (Convert.ToString(dr["SLT_STATUS"]).ToLower().Trim() == "not met")
                        {
                            img = "<img  src='http://asfra-dev.jnj.com/ASFRA/Content/Circle/Green.png' width:20pt;  alt='SLT NOT MET' />";

                        }
                        if (Convert.ToString(dr["SLT_STATUS"]).ToLower().Trim() == "wip")
                        {
                            img = "<img  src='http://asfra-dev.jnj.com/ASFRA/Content/Circle/DrakGreen.png' width:20pt;  alt='WIP' />";

                        }
                        htmlbody += "<tr><td>" + SN + "</td><td>" +
                               dr["BUSINESS_DELIVERABLE"] + "</td><td>" +
                               dr["FREQUENCY"] + "</td><td>" +
                               dr["SOURCE_NAME"] + "</td><td>" +
                               dr["CATEGORY"] + "</td><td>" +
                               dr["STATUS"] + "</td><td >" +
                               img + " </td><td>" +
                              dr["ESTIMATEDSTARTTIME"] + "</td><td>" +
                              dr["ESTIMATEDENDTIME"] + "</td><td>" +
                                dr["ACTUALSTARTTIME"] + "</td><td>" +
                               dr["ACTUALENDTIME"] + "</td>" +
                               "<td></td><td></td></tr>";
                        SN = SN + 1;
                    }
                }
                else
                {
                    foreach (DataRow dr in dt.Rows)
                    {

                        string img = "";
                        if (Convert.ToString(dr["SLT_STATUS"]).ToLower().Trim() == "met")
                        {
                            img = "<img  src='http://asfra-dev.jnj.com/ASFRA/Content/Circle/Green.png' width:20pt;  alt='SLT MET' />";

                        }

                        if (Convert.ToString(dr["SLT_STATUS"]).ToLower().Trim() == "not met")
                        {
                            img = "<img  src='http://asfra-dev.jnj.com/ASFRA/Content/Circle/Red.png' width:20pt;  alt='SLT NOT MET' />";

                        }

                        if (Convert.ToString(dr["SLT_STATUS"]).ToLower().Trim() == "wip")
                        {
                            img = "<img  src='http://asfra-dev.jnj.com/ASFRA/Content/Circle/amber.png' width:20pt;  alt='WIP' />";

                        }

                        htmlbody += "<tr><td>" + SN + "</td><td>" +
                               dr["BUSINESS_DELIVERABLE"] + "</td><td>" +
                               dr["FREQUENCY"] + "</td><td>" +
                               dr["SOURCE_NAME"] + "</td><td>" +
                               dr["CATEGORY"] + "</td><td>" +
                               dr["STATUS"] + "</td><td>" +
                              "<center>" + img + "</center></td><td>" +
                               dr["ESTIMATEDSTARTTIME"] + "</td><td>" +
                               dr["ESTIMATEDENDTIME"] + "</td><td>" +
                               dr["ACTUALSTARTTIME"] + "</td><td>" +
                               dr["ACTUALENDTIME"] + "</td>" +
                               "<td></td><td></td></tr>";

                        SN = SN + 1;
                    }
                }
            }
            
            htmlbody += "</table></body></html>";
            
            string frequency = maildetails.Rows[0]["EMAIL_FREQUENCY"].ToString().ToLower();
            string day = DateTime.Today.DayOfWeek.ToString().ToLower();
            if (frequency.Contains(day))
            {

                //MailMessage mail = new MailMessage("lgupta3@its.jnj.com", "lgupta3@its.jnj.com,spt@its.jnj.com,ssakthi2@its.jnj.com");
                MailMessage mail = new MailMessage("lgupta3@its.jnj.com", maildetails.Rows[0]["EMAIL_TO_LIST"].ToString());
                SmtpClient client = new SmtpClient();
                client.Port = 25;
                mail.CC.Add(maildetails.Rows[0]["EMAIL_CC_LIST"].ToString());
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Host = "smtp1.eesus.jnj.com";
                mail.Subject = maildetails.Rows[0]["EMAIL_SUBJECT"].ToString();
                mail.IsBodyHtml = true;


                // mail.Bcc.Add("pgoyal4@its.jnj.com");
                AlternateView alternativeView = AlternateView.CreateAlternateViewFromString(htmlbody, null, MediaTypeNames.Text.Html);
                mail.AlternateViews.Add(alternativeView);
                //Sending an email
                Task.Delay(1000);
                client.Send(mail);
            }
        }
        public static DataTable ConsumerEdgeCubeRefresh(DataTable dt,int appid)
        {
            string userName = "", password = "", host = "";
            using (SqlConnection sqlcon = new SqlConnection(sqlconnstring))
            {
                sqlcon.Open();
                SqlCommand getappids = new SqlCommand("select [NTID], CAST([PASSWORD] AS VARCHAR(MAX)) as [Password], [AGENTS], [ENVIRONMENT],[URL] from [dbo].[CONN_DETAILS] where APPID=" + appid + "and Source ='Oracle'", sqlcon);
                SqlDataReader dr = getappids.ExecuteReader();
                while (dr.Read())
                {
                    userName = dr["NTID"].ToString();
                    password = dr["Password"].ToString();
                    host = dr["URL"].ToString();

                }
                sqlcon.Close();
            }
            string connectionstring = "User Id="+userName+";Password="+password+";Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="+host+")(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=PDDB5493.jnj.com)))";
            using (OracleConnection oraclecon = new OracleConnection(connectionstring))
            {
                DataRow dr1 = dt.NewRow();
                oraclecon.Open();
                OracleCommand cmd = new OracleCommand("select max(UPD_DT_TIME) , min(CUBE_REF_START)  from TD_CUBE_REFRESH_STATUS where NAME in ('Trend Topline','Customer P&L')", oraclecon);

                OracleDataReader rdr = cmd.ExecuteReader();
                //int i = 0;
                while (rdr.Read())
                {
                    DateTime endtime = Convert.ToDateTime(rdr["max(UPD_DT_TIME)"].ToString());
                    DateTime starttime = Convert.ToDateTime(rdr["min(CUBE_REF_START)"].ToString());
                    if (endtime.Date >= DateTime.Today.Date)
                    {
                        TimeSpan ts = (endtime - starttime);
                        int days = Math.Abs(ts.Days);
                        int hour = Math.Abs(ts.Hours);
                        int min = Math.Abs(ts.Minutes);
                        int sec = Math.Abs(ts.Seconds);
                        string duration = days + ":" + hour.ToString("00") + ":" + min.ToString("00") + ":" + sec.ToString("00");
                        dr1["ACTUALDURATION"] = duration;
                        if (days > 0 || hour > 0 || min > 0 || sec > 0)
                        {
                            dr1["STATUS"] = "Completed Normally";
                        }
                    }
                    else
                    {
                        dr1["ACTUALDURATION"] = DBNull.Value;
                        dr1["STATUS"] = DBNull.Value;
                    }
                    dr1["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                    dr1["SRC_DATA"] = "Consumer Edge Database";
                    dr1["APPID"] = "2";



                    dr1["ACTUALSTARTTIME"] = Convert.ToDateTime(rdr["min(CUBE_REF_START)"].ToString());
                    dr1["JOB"] = "P-AS-001_TREND_TOPLINE";
                    dr1["JOBGROUP"] = "P-AS-001_TREND_TOPLINE";


                }
                dt.Rows.Add(dr1);
                DataRow dr2 = dt.NewRow();
                OracleCommand cmd1 = new OracleCommand("select max(UPD_DT_TIME) , min(CUBE_REF_START) from TD_CUBE_REFRESH_STATUS where AREA ='SHIPMENT'", oraclecon);
                OracleDataReader rdr1 = cmd1.ExecuteReader();
                while (rdr1.Read())
                {
                    DateTime endtime = Convert.ToDateTime(rdr1["MAX(UPD_DT_TIME)"].ToString());
                    DateTime starttime = Convert.ToDateTime(rdr1["MIN(CUBE_REF_START)"].ToString());
                    if (endtime.Date <= DateTime.Today.Date)
                    {
                        TimeSpan ts = (endtime - starttime);
                        int days = Math.Abs(ts.Days);
                        int hour = Math.Abs(ts.Hours);
                        int min = Math.Abs(ts.Minutes);
                        int sec = Math.Abs(ts.Seconds);
                        string duration = days + ":" + hour.ToString("00") + ":" + min.ToString("00") + ":" + sec.ToString("00");
                        dr2["ACTUALDURATION"] = duration;
                        if (days > 0 || hour > 0 || min > 0 || sec > 0)
                        {
                            dr2["STATUS"] = "Completed Normally";
                        }
                    }
                    else
                    {
                        dr2["ACTUALDURATION"] = DBNull.Value;
                        dr2["STATUS"] = DBNull.Value;
                    }
                    dr2["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                    dr2["SRC_DATA"] = "Consumer Edge Database";
                    dr2["APPID"] = "2";

                    dr2["ACTUALSTARTTIME"] = Convert.ToDateTime(rdr1["min(CUBE_REF_START)"].ToString());
                    dr2["JOB"] = "S-AS-001_SHIP_ATTAINMENT";
                    dr2["JOBGROUP"] = "S-AS-001_SHIP_ATTAINMENT";

                }
                dt.Rows.Add(dr2);
                DataRow dr3 = dt.NewRow();
                OracleCommand cmd2 = new OracleCommand("select max(UPD_DT_TIME) , min(CUBE_REF_START)   from TD_CUBE_REFRESH_STATUS where NAME = 'Integrated Weekly'", oraclecon);
                OracleDataReader rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    DateTime endtime = Convert.ToDateTime(rdr2["max(UPD_DT_TIME)"].ToString());
                    DateTime starttime = Convert.ToDateTime(rdr2["min(CUBE_REF_START)"].ToString());
                    if (endtime.Date <= DateTime.Today.Date)
                    {
                        TimeSpan ts = (endtime - starttime);

                        int days = Math.Abs(ts.Days);
                        int hour = Math.Abs(ts.Hours);
                        int min = Math.Abs(ts.Minutes);
                        int sec = Math.Abs(ts.Seconds);
                        string duration = days + ":" + hour.ToString("00") + ":" + min.ToString("00") + ":" + sec.ToString("00");
                        dr3["ACTUALDURATION"] = duration;
                        if (days > 0 || hour > 0 || min > 0 || sec > 0)
                        {
                            dr1["STATUS"] = "Completed Normally";
                        }
                    }
                    else
                    {
                        dr3["ACTUALDURATION"] = DBNull.Value;
                        dr3["STATUS"] = DBNull.Value;
                    }
                    dr3["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                    dr3["SRC_DATA"] = "Consumer Edge Database";
                    dr3["APPID"] = "2";

                    dr3["ACTUALSTARTTIME"] = Convert.ToDateTime(rdr2["min(CUBE_REF_START)"].ToString());
                    dr3["JOB"] = "I-AS-001_ISIS_INTEGRATED_WEEKLY";
                    dr3["JOBGROUP"] = "I-AS-001_ISIS_INTEGRATED_WEEKLY";
                }

                dt.Rows.Add(dr3);
                oraclecon.Close();
                return dt;

            }

        }
        public static DataTable ConsumerEdgeReportDelievery( DataTable dt, int appid)
        {
            string password = "", userName = "", url = "";
            DataRow datarow;
            IWebDriver driver = new FirefoxDriver();
            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {
                con.Open();
                SqlCommand getappids = new SqlCommand("select [NTID], CAST([PASSWORD] AS VARCHAR(MAX)) as [Password], [AGENTS], [ENVIRONMENT],[URL] from [dbo].[CONN_DETAILS] where APPID=" + appid + "and Source ='.NET Portal'", con);
                SqlDataReader dr = getappids.ExecuteReader();
                while (dr.Read())
                {
                    userName = dr["NTID"].ToString();
                    password = dr["Password"].ToString();
                    url = dr["URL"].ToString();

                }
                //string URL = "http://" + username + ":" + password + "@" + "isis.na.jnj.com/isisrptdeliverystatus/ReportDeliveryStatus.aspx";
            }
        

            driver.Navigate().GoToUrl(url);

            Task.Delay(2000).Wait();

            // IAlert alert = driver.SwitchTo().Alert();


            Task.Delay(3000);




            AutoItX3 autoIT = new AutoItX3();

            autoIT.Send(userName);

            Task.Delay(2000).Wait();

            autoIT.Send("{TAB}");

            autoIT.Send(password);

            Task.Delay(2000).Wait();

            autoIT.Send("{ENTER}");





            datarow = dt.NewRow();

            //driver.Navigate().GoToUrl(URL);

            Task.Delay(5000).Wait();
            string startJobstatus = driver.FindElement(By.XPath("//*[@id='TblStatus']/tbody/tr[2]/td[5]")).Text.ToString();
            IWebElement table_element1 = driver.FindElement(By.Id("TblStatus"));
            IWebElement x1 = driver.FindElement(By.XPath("//*[@id='TblStatus']/tbody/tr[2]/td[4]"));

            string startJobTime = x1.Text;
            string endJobTime = "";
            string endJobstatus = "";
            if (DateTime.Today.DayOfWeek.ToString().ToLower() == "monday" & startJobstatus == "Delivered")
            {
                driver.FindElement(By.XPath("//*[@id='cmbFrequency']/option[2]")).Click();
                endJobstatus = driver.FindElement(By.XPath("//*[@id='TblStatus']/tbody/tr[126]/td[5]")).Text.ToString();
                endJobTime = driver.FindElement(By.XPath("//*[@id='TblStatus']/tbody/tr[126]/td[4]")).Text.ToString();
                endJobstatus = driver.FindElement(By.XPath("//*[@id='TblStatus']/tbody/tr[126]/td[5]")).Text.ToString();
            }
            else
            {
                if (startJobstatus == "Delivered")
                {

                    IWebElement x2 = driver.FindElement(By.XPath("//*[@id='TblStatus']/tbody/tr[107]/td[4]"));
                    endJobstatus = driver.FindElement(By.XPath("//*[@id='TblStatus']/tbody/tr[107]/td[5]")).Text.ToString();


                    endJobTime = x2.Text;
                }

            }
            DateTime startdate;
            DateTime enddate;
            if (startJobTime.Trim() != "" & endJobTime.Trim() != null)
            {
                startdate = Convert.ToDateTime(startJobTime);
                enddate = Convert.ToDateTime(endJobTime);
                if (enddate.Date >= DateTime.Today.Date)
                {
                    TimeSpan ts = (enddate - startdate);

                    int days = Math.Abs(ts.Days);
                    int hour = Math.Abs(ts.Hours);
                    int min = Math.Abs(ts.Minutes);
                    int sec = Math.Abs(ts.Seconds);
                    string duration = days + ":" + hour.ToString("00") + ":" + min.ToString("00") + ":" + sec.ToString("00");
                    datarow["ACTUALDURATION"] = duration;

                    datarow["STATUS"] = "Delivered";
                }
                else
                {
                    datarow["ACTUALDURATION"] = DBNull.Value;
                    datarow["STATUS"] = DBNull.Value;
                }
                datarow["ACTUALSTARTTIME"] = Convert.ToDateTime(startJobTime);

                datarow["JOB"] = "Report Delivery";
                datarow["SRC_DATA"] = "Report Delivery Website";

                datarow["APPID"] = "2";
                datarow["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                datarow["JOBGROUP"] = "Report Delivery";
            }
            else
            {
                datarow["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                datarow["ACTUALDURATION"] = "";
                datarow["JOB"] = "Report Delivery";
                datarow["SRC_DATA"] = "Report Delivery Website";
                datarow["STATUS"] = "Not Delivered";
                datarow["APPID"] = "2";
                datarow["JOBGROUP"] = "Report Delivery";
            }
            dt.Rows.Add(datarow);
            driver.Quit();
            return dt;
        }
        public static DataTable ConsumerEdgeTableauReport(DataTable dt, int appid)
        {
            string userName = "", password = "", url = "";
            string[] arr = new string[20];
            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {

                con.Open();
                SqlCommand getappids = new SqlCommand("select [NTID], CAST([PASSWORD] AS VARCHAR(MAX)) as [Password], [AGENTS], [ENVIRONMENT],[URL] from [dbo].[CONN_DETAILS] where APPID=" + appid + "and Source ='Tableau Third Portal'", con);
                SqlDataReader sdr = getappids.ExecuteReader();
                while (sdr.Read())
                {
                    userName = sdr["NTID"].ToString();
                    password = sdr["Password"].ToString();
                    url = @sdr["URL"].ToString();

                }
                //string URL = "http://" + username + ":" + password + "@" + "isis.na.jnj.com/isisrptdeliverystatus/ReportDeliveryStatus.aspx";
            }
            //DateTime starttime, endtime;
            DataRow dr = dt.NewRow();

            TimeSpan ts;
            int hours = 0, mins = 0, days = 0, secs = 0;
            
            OpenQA.Selenium.IWebDriver driver = null;

            driver = new OpenQA.Selenium.Firefox.FirefoxDriver();

           
            String script = "window.location = \'" + url + "\'";
            string[] reports = new string[3];
            reports[0] = "Integrated weekly (BP)";
            reports[1] = "Certified US Integrated Weekly";
            reports[2] = "Quota by Month";
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

            js.ExecuteScript(script);

            Task.Delay(30000).Wait();
            driver.FindElement(By.Id("username")).SendKeys(userName);
            driver.FindElement(By.Id("password")).SendKeys(password);
           

            Task.Delay(3000).Wait();

            driver.FindElement(By.XPath("//*/a[contains(text(),'Sign On')]")).Click();
            TimeZone localZone = TimeZone.CurrentTimeZone;
            Task.Delay(12000).Wait();
            string str = driver.FindElement(By.XPath("html/body/div[1]/div/div/div[1]/div/div/div[1]/div[2]/a[3]/span[2]")).Text;
            int numberofdatasource = Convert.ToInt32(driver.FindElement(By.XPath("html/body/div[1]/div/div/div[1]/div/div/div[1]/div[2]/a[3]/span[2]")).Text);
            for (int i = 1; i <= numberofdatasource; i++)
            {
                // "html/body/div[1]/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/   div[1]  /div[8]/span/span[2]"
                // "html/body/div[1]/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/   div[5]  /div[8]/span/span[2]"
                // "html/body/div[1]/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/   div[6]  /div[8]/span/span[2]"

                var endtime = "html/body/div[1]/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[8]/span/span[2]";
                var reportnamexpath = "html/body/div[1]/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[3]/div/div[1]/span/a";
                string reportname = driver.FindElement(By.XPath(reportnamexpath)).Text;

                if (DateTime.Today.ToString("dddd").ToLower() == "monday")
                {
                    if (reportname.ToLower().Contains(reports[0].ToLower()))
                    {
                        dr = dt.NewRow();
                        dr["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                        dr["SRC_DATA"] = "Tableaue Extractor";
                        dr["APPID"] = appid;
                        DateTime starttime1 = DateTime.Today.AddHours(8);
                        string date = starttime1.ToString("yyyy-MM-dd hh:mm:ss");
                        dr["ACTUALSTARTTIME"] = starttime1.ToString("yyyy-MM-dd hh:mm:ss");
                        //   dr["ACTUALDURATION"] = days + ":" + hour + ":" + min + ":" + sec;
                        // dr["ACTUALSTARTTIME"] = Convert.ToDateTime(starttime);
                        dr["JOB"] = "Shared Server -  Integrated weekly (BP)";
                        dr["JOBGROUP"] = "Shared Server -  Integrated weekly (BP)";

                        string endtimestamp = driver.FindElement(By.XPath(endtime)).Text;
                        if (endtimestamp.Length > 0)
                        {

                            DateTime endtime1 = ConvertDate(endtimestamp);
                            if (endtime1.Date >= DateTime.Today.Date)
                            {

                                ts = endtime1 - starttime1;



                                hours = Math.Abs(ts.Hours);
                                days = Math.Abs(ts.Days);
                                mins = Math.Abs(ts.Minutes);
                                secs = Math.Abs(ts.Seconds);
                                dr["ACTUALDURATION"] = days + ":" + hours.ToString("00") + ":" + mins.ToString("00") + ":" + secs.ToString("00");
                                if (days > 0 || hours > 0 || mins > 0 || secs > 0)
                                {
                                    dr["STATUS"] = "Completed Normally";
                                }
                            }
                        }
                        else
                        {
                            dr["ACTUALDURATION"] = DBNull.Value;
                            dr["STATUS"] = DBNull.Value;
                        }
                        dt.Rows.Add(dr);
                    }

                }

                if (DateTime.Today.ToString("dddd").ToLower() == "monday")
                {
                    if (reportname.ToLower() ==reports[1].ToLower())
                    {
                        dr = dt.NewRow();


                        dr["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                        dr["SRC_DATA"] = "Tableaue Extractor";
                        dr["APPID"] = appid;
                        DateTime starttime2 = DateTime.Today.AddHours(8);


                        dr["ACTUALSTARTTIME"] = starttime2.ToString("yyyy-MM-dd hh:mm:ss");
                        //   dr["ACTUALDURATION"] = days + ":" + hour + ":" + min + ":" + sec;
                        // dr["ACTUALSTARTTIME"] = Convert.ToDateTime(starttime);
                        dr["JOB"] = "Shared Server-  Integrated Weekly";
                        dr["JOBGROUP"] = "Shared Server - Integrated Weekly";

                        string endtimestamp = driver.FindElement(By.XPath(endtime)).Text;
                        if (endtimestamp.Length > 0)
                        {
                            DateTime endtime2 = ConvertDate(endtimestamp);
                            if (endtime2.Date >= DateTime.Today.Date)
                            {
                                ts = endtime2 - starttime2;
                                hours = Math.Abs(ts.Hours);
                                days = Math.Abs(ts.Days);
                                mins = Math.Abs(ts.Minutes);
                                secs = Math.Abs(ts.Seconds);
                                dr["ACTUALDURATION"] = days + ":" + hours.ToString("00") + ":" + mins.ToString("00") + ":" + secs.ToString("00");
                                if (days > 0 || hours > 0 || mins > 0 || secs > 0)
                                {
                                    dr["STATUS"] = "Completed Normally";
                                }
                            }
                        }
                        else
                        {
                            dr["ACTUALDURATION"] = DBNull.Value; dr["STATUS"] = DBNull.Value;
                        }
                        dt.Rows.Add(dr);
                    }

                }


                if (reportname.ToLower().Contains(reports[2].ToLower()))
                {
                    dr = dt.NewRow();
                    dr["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                    dr["SRC_DATA"] = "Tableaue Extractor";
                    dr["APPID"] = appid;
                    DateTime starttime3 = DateTime.Today.AddHours(8);




                    dr["JOB"] = "Shared server- Quota by Month";
                    dr["JOBGROUP"] = "Shared server- Quota by Month";

                    string endtimestamp1 = driver.FindElement(By.XPath(endtime)).Text;
                    if (endtimestamp1.Length > 0)
                    {
                        DateTime endtime3 = ConvertDate(endtimestamp1);
                        if (endtime3.Date >= DateTime.Today.Date)
                        {
                            dr["ACTUALSTARTTIME"] = starttime3;
                            ts = endtime3 - starttime3;
                            hours = Math.Abs(ts.Hours);
                            days = Math.Abs(ts.Days);
                            mins = Math.Abs(ts.Minutes);
                            secs = Math.Abs(ts.Seconds);
                            dr["ACTUALDURATION"] = days + ":" + hours.ToString("00") + ":" + mins.ToString("00") + ":" + secs.ToString("00");
                            if (days > 0 || hours > 0 || mins > 0 || secs >= 0)
                            {
                                dr["STATUS"] = "Completed Normally";
                            }
                        }
                    }
                    else
                    {
                        dr["ACTUALDURATION"] = DBNull.Value;
                        dr["STATUS"] = DBNull.Value;
                    }
                    dt.Rows.Add(dr);
                }
            }

            driver.Quit();



            return dt;
        }
        public static DataTable Consumeredge_TableauExtract_new(IWebDriver driver, DataTable dt, String ReportName, int appid)// newly addedqwweer
        {
            //string URL = @"http://tableauedge.jnj.com/#/site/NAAI/projects/97/workbooks";
            //driver.Navigate().GoToUrl(URL);
            //AutoItX3 autoIT = new AutoItX3();
            //autoIT.Send("{F6}");
            //Task.Delay(1000).Wait();

            //autoIT.Send("http://tableauedge.jnj.com/#/site/NAAI/projects/97/workbooks");
            //Task.Delay(2000).Wait();

            //autoIT.Send("{ENTER}");
            //Task.Delay(20000).Wait();

            driver.FindElement(By.XPath("html/body/div/div/div/div[1]/div/div/div[1]/div/a[2]")).Click();
            Task.Delay(3000).Wait();

            driver.FindElement(By.XPath("html/body/div/div/div/div[1]/div/div/div[2]/div/div[3]/div/div/div/span/div/div/div/span[4]/div[1]/a")).Click();
            Task.Delay(5000).Wait();

            //driver.FindElement(By.XPath("/html/body/div/div/div/div[1]/div/div/div[1]/div[2]/a[2]/span[1]")).Click();
            //Task.Delay(2000);

            driver.FindElement(By.XPath("html/body/div/div/div/div[1]/div/div/div[1]/div[2]/a[2]/span[1]")).Click();
            Task.Delay(5000).Wait();

            //String end_time = driver.FindElement(By.XPath("html/body/div/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div/div/span/div/div[2]/div/div[1]/div/div/div[4]/span/span[2]")).Text;
            String Filename = "";

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

            string reportNamePath = "html/body/div/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div/div/span/div/div[2]/div/div[1]/div/div/div[2]/span/span[2]";
            //js.ExecuteScript("document.getElementByXPath('//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div/div/span/div/div[2]/div/div[1]/div/div/div[2]/span/span[2]').Text");
            try
            {
                if (CheckElementAvailable(driver, reportNamePath))
                {
                    // js.ExecuteScript("document.getElementByXPath('//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div/div/span/div/div[2]/div/div[1]/div/div/div[2]/span/span[2]').Text");
                    Filename = driver.FindElement(By.XPath("html/body/div/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div/div/span/div/div[2]/div/div[1]/div/div/div[2]/span/span[2]")).Text;
                }
                else
                {
                    reportNamePath = "//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div/div/span/div/div[2]/div/div[1]/div/div/div[2]/span/span[2]";
                    if (CheckElementAvailable(driver, reportNamePath))
                    {
                        Filename = driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div/div/span/div/div[2]/div/div[1]/div/div/div[2]/span/span[2]")).Text;
                    }
                }

            }
            catch (Exception)
            {
                Filename = driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div/div/span/div/div[2]/div/div[1]/div/div/div[2]/span/span[2]")).Text;
            }
            if (Filename == ReportName)
            {
                DataRow dr = dt.NewRow();
                dr = dt.NewRow();
                dr["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                dr["SRC_DATA"] = "Tableaue Extractor";
                dr["APPID"] = appid;
                dr["JOB"] = "NAAI- Returns Dashboard";

                dr["JOBGROUP"] = "NAAI- Returns Dashboard";
                String EndTime = driver.FindElement(By.XPath("html/body/div/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div/div/span/div/div[2]/div/div[1]/div/div/div[4]/span/span[2]")).Text;
                if (EndTime.Length > 0)
                {

                    DateTime Quotabymonthendtime = ConvertDate(EndTime);
                    DateTime Quotabymonthstarttime = DateTime.Today.AddHours(4);
                    if (Quotabymonthendtime.Date >= DateTime.Today.Date)
                    {
                        TimeSpan ts;
                        int hours = 0, mins = 0, days = 0, secs = 0;

                        dr["ACTUALSTARTTIME"] = Quotabymonthstarttime;
                        //   dr["ACTUALDURATION"] = days + ":" + hour + ":" + min + ":" + sec;
                        // dr["ACTUALSTARTTIME"] = Convert.ToDateTime(starttime);
                        ts = Quotabymonthendtime - Quotabymonthstarttime;
                        hours = Math.Abs(ts.Hours);
                        days = Math.Abs(ts.Days);
                        mins = Math.Abs(ts.Minutes);
                        secs = Math.Abs(ts.Seconds);
                        dr["ACTUALDURATION"] = days + ":" + hours.ToString("00") + ":" + mins.ToString("00") + ":" + secs.ToString("00");
                        if (days > 0 || hours > 0 || mins > 0 || secs >= 0)
                        {
                            dr["STATUS"] = "Completed Normally";
                        }
                    }
                    else
                    {
                        dr["ACTUALDURATION"] = DBNull.Value;
                        dr["STATUS"] = DBNull.Value;
                    }
                }
                dt.Rows.Add(dr);


            }

            return dt;
        }
        static DateTime ConvertDate(string datetimestamp)
        {
            TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
            TimeZone localZone = TimeZone.CurrentTimeZone;
            int index = datetimestamp.IndexOf(',', datetimestamp.IndexOf(',') + 1);
            string datepart = datetimestamp.Substring(0, index);
            string format = "MMM d, yyyy";

            DateTime date = DateTime.ParseExact(datepart, format, new CultureInfo("en-US"));
            string time = datetimestamp.Substring(index + 1);
            datetimestamp = date.ToString("yyyy-MM-dd") + " " + time;
            date = Convert.ToDateTime(datetimestamp);
            string test = localZone.StandardName.ToString().Trim().ToLower();
            if (localZone.StandardName.ToString().Trim().ToLower().Contains("india"))
            {

                date = TimeZoneInfo.ConvertTimeFromUtc(date, easternZone);
            }

            return date;
        }
        static DateTime TimeInSTZ() // Singapore Time Zone
        {

            TimeZoneInfo singaporeZone = TimeZoneInfo.FindSystemTimeZoneById("Singapore Standard Time");
            TimeZone localZone = TimeZone.CurrentTimeZone;
            DateTime currentTime = DateTime.UtcNow;
            if (!localZone.StandardName.ToString().Trim().ToLower().Contains("Singapore"))
            {

                currentTime = TimeZoneInfo.ConvertTimeFromUtc(currentTime, singaporeZone);
            }

            return currentTime;
        } 
        public static DataTable ConsumerEdgeTableauReport(OpenQA.Selenium.IWebDriver driver, int appid, DataTable dt)
        {
            string userName = "", password = "", url = "";
            string[] arr = new string[20];
            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {
                
                con.Open();
                SqlCommand getappids = new SqlCommand("select [NTID], CAST([PASSWORD] AS VARCHAR(MAX)) as [Password], [AGENTS], [ENVIRONMENT],[URL] from [dbo].[CONN_DETAILS] where APPID=" + appid + "and Source ='Tableau Second Portal'", con);
                SqlDataReader dr = getappids.ExecuteReader();
                while (dr.Read())
                {
                    userName = dr["NTID"].ToString();
                    password = dr["Password"].ToString();
                    url = @dr["URL"].ToString();

                }
                //string URL = "http://" + username + ":" + password + "@" + "isis.na.jnj.com/isisrptdeliverystatus/ReportDeliveryStatus.aspx";
            }
            DataRow datarow = dt.NewRow();

            TimeSpan ts;
            int hours = 0, mins = 0, days = 0, secs = 0;
            
            driver.Navigate().GoToUrl(url);
            Task.Delay(5000).Wait();
            driver.FindElement(By.Name("username")).SendKeys(userName);

            driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[2]/span/form/div[1]/div[2]/div/div/input")).SendKeys(password);
            // Task.Delay(10000);
            driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[2]/span/form/button")).Click();
            Task.Delay(8000).Wait();
            driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[2]/div/span/div/div/div/div/div[1]/span/span/div/span")).Click();
            Task.Delay(5000).Wait();
            driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[2]/div/span/div/div/div/div/div[1]/span/div/div/div/div/div/div/a[2]")).Click();
            Task.Delay(5000).Wait();
            driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[1]/div/a[4]")).Click();
            Task.Delay(3000).Wait();
            driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[1]/span/span[3]/div[1]/div")).Click();
            Task.Delay(3000).Wait();

            string[] ReportName = new string[5];
            ReportName[0] = "Quota By Month";
            ReportName[1] = "Integrated Weekly (BP)";
            ReportName[2] = "Certified US Integrated Weekly";
            ReportName[3] = "Integrated Weekly (PROD)";
            ReportName[4] = "Returns DB Extract";// Newly added
            string EndTime;
            int numberOfReports = Convert.ToInt32(driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[1]/div/a[4]/span[2]")).Text); 
            IWebElement div = driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div"));
            List<IWebElement> tablerow = div.FindElements(By.TagName("div")).ToList();
            int i = 1;

            if (tablerow.Count > 0)
            {
                foreach (IWebElement row in tablerow)
                {
                    string report = driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[3]/div/div[1]/span/a")).Text;
                    if (driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[3]/div/div[1]/span/a")).Text.ToString() == ReportName[0])
                    //*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[3]/div/div[1]/span/a
                    {
                        datarow = dt.NewRow();
                        datarow["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                        datarow["SRC_DATA"] = "Tableaue Extractor";
                        datarow["APPID"] = appid;
                        datarow["JOB"] = "Edge- " + ReportName[0];

                        datarow["JOBGROUP"] = "Edge- " + ReportName[0];
                        EndTime = driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[9]/span/span[2]")).Text;
                        if (EndTime.Length > 0)
                        {

                            DateTime Quotabymonthendtime = ConvertDate(EndTime);
                            DateTime Quotabymonthstarttime = DateTime.Today.AddHours(4);
                            if (Quotabymonthendtime.Date >= DateTime.Today.Date)
                            {
                                datarow["ACTUALSTARTTIME"] = Quotabymonthstarttime;
                                //   dr["ACTUALDURATION"] = days + ":" + hour + ":" + min + ":" + sec;
                                // dr["ACTUALSTARTTIME"] = Convert.ToDateTime(starttime);
                                ts = Quotabymonthendtime - Quotabymonthstarttime;
                                hours = Math.Abs(ts.Hours);
                                days = Math.Abs(ts.Days);
                                mins = Math.Abs(ts.Minutes);
                                secs = Math.Abs(ts.Seconds);
                                datarow["ACTUALDURATION"] = days + ":" + hours.ToString("00") + ":" + mins.ToString("00") + ":" + secs.ToString("00");
                                if (days > 0 || hours > 0 || mins > 0 || secs > 0)
                                {
                                    datarow["STATUS"] = "Completed Normally";
                                }
                            }
                            else
                            {
                                datarow["ACTUALDURATION"] = DBNull.Value;
                                datarow["STATUS"] = DBNull.Value;
                            }
                        }
                        dt.Rows.Add(datarow);
                    }

                    if (DateTime.Today.ToString("dddd").ToLower() == "tuesday")
                    {

                        string report1 = driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[3]/div/div[1]/span/a")).Text;
                        if (driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[3]/div/div[1]/span/a")).Text == ReportName[1])
                        {

                            datarow = dt.NewRow();

                            datarow["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                            datarow["SRC_DATA"] = "Tableaue Extractor";
                            datarow["APPID"] = appid;
                            datarow["JOB"] = "EDGE - " + ReportName[1];
                            datarow["JOBGROUP"] = "EDGE - " + ReportName[1];
                            EndTime = driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[9]/span/span[2]")).Text;
                            if (EndTime.Length > 0)
                            {

                                DateTime endtime4 = ConvertDate(EndTime);
                                DateTime starttime4 = DateTime.Today.AddHours(4);
                                if (endtime4.Date >= DateTime.Today.Date)
                                {
                                    datarow["ACTUALSTARTTIME"] = starttime4;
                                    //   dr["ACTUALDURATION"] = days + ":" + hour + ":" + min + ":" + sec;
                                    // dr["ACTUALSTARTTIME"] = Convert.ToDateTime(starttime);
                                    ts = endtime4 - starttime4;
                                    hours = Math.Abs(ts.Hours);
                                    days = Math.Abs(ts.Days);
                                    mins = Math.Abs(ts.Minutes);
                                    secs = Math.Abs(ts.Seconds);
                                    datarow["ACTUALDURATION"] = days + ":" + hours.ToString("00") + ":" + mins.ToString("00") + ":" + secs.ToString("00");
                                    if (days > 0 || hours > 0 || mins > 0 || secs > 0)
                                    {
                                        datarow["STATUS"] = "Completed Normally";
                                    }
                                }
                                else
                                {
                                    datarow["ACTUALDURATION"] = DBNull.Value;
                                    datarow["STATUS"] = DBNull.Value;
                                }

                            }
                            //dr["EndTime"] = "";
                            dt.Rows.Add(datarow);

                        }
                        string repotr = driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[3]/div/div[1]/span/a")).Text;
                        //if (driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[3]/div/div[1]/span/a")).Text == ReportName[3])
                        //{

                        //    dr = dt.NewRow();

                        //    dr["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                        //    dr["SRC_DATA"] = "Tableaue Extractor";
                        //    dr["APPID"] = appid;
                        //    dr["JOB"] = "EDGE - " + ReportName[2];
                        //    dr["JOBGROUP"] = "EDGE - " + ReportName[2];
                        //    EndTime = driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[9]/span/span[2]")).Text;
                        //    if (EndTime.Length > 0)
                        //    {
                        //        DateTime endtime4 = ConvertDate(EndTime);
                        //        DateTime starttime4 = DateTime.Today.AddHours(4);
                        //        if (endtime4.Date >= DateTime.Today.Date)
                        //        {
                        //            dr["ACTUALSTARTTIME"] = starttime4;
                        //            //   dr["ACTUALDURATION"] = days + ":" + hour + ":" + min + ":" + sec;
                        //            // dr["ACTUALSTARTTIME"] = Convert.ToDateTime(starttime);
                        //            ts = endtime4 - starttime4;
                        //            hours = Math.Abs(ts.Hours);
                        //            days = Math.Abs(ts.Days);
                        //            mins = Math.Abs(ts.Minutes);
                        //            secs = Math.Abs(ts.Seconds);
                        //            dr["ACTUALDURATION"] = days + ":" + hours.ToString("00") + ":" + mins.ToString("00") + ":" + secs.ToString("00");
                        //            if (days > 0 || hours > 0 || mins > 0 || secs > 0)
                        //            {
                        //                dr["STATUS"] = "Completed Normally";
                        //            }
                        //        }
                        //        else
                        //        {
                        //            dr["ACTUALDURATION"] = DBNull.Value;
                        //            dr["STATUS"] = DBNull.Value;
                        //        }
                        //    }

                        //    //dr["EndTime"] = "";
                        //    dt.Rows.Add(dr);

                        //}
                        if (driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[3]/div/div[1]/span/a")).Text == ReportName[2])
                        {

                            datarow = dt.NewRow();

                            datarow["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                            datarow["SRC_DATA"] = "Tableaue Extractor";
                            datarow["APPID"] = appid;
                            datarow["JOB"] = "EDGE - " + ReportName[2];
                            datarow["JOBGROUP"] = "EDGE - " + ReportName[2];
                            //dr["JOB"] = "Edge- " + ReportName[2];
                            //dr["JOBGROUP"] ="Edge- " + ReportName[2];
                            EndTime = driver.FindElement(By.XPath("//*[@id='ng-app']/div/div/div[1]/div/div/div[2]/div[1]/div/div/div[3]/div/div/span/div/div[2]/div/div[1]/div/div[" + i + "]/div[9]/span/span[2]")).Text;
                            if (EndTime.Length > 0)
                            {

                                DateTime endtime4 = ConvertDate(EndTime);
                                DateTime starttime4 = DateTime.Today.AddHours(4);
                                if (endtime4.Date >= DateTime.Today.Date)
                                {
                                    datarow["ACTUALSTARTTIME"] = starttime4;
                                    //   dr["ACTUALDURATION"] = days + ":" + hour + ":" + min + ":" + sec;
                                    // dr["ACTUALSTARTTIME"] = Convert.ToDateTime(starttime);
                                    ts = endtime4 - starttime4;
                                    hours = Math.Abs(ts.Hours);
                                    days = Math.Abs(ts.Days);
                                    mins = Math.Abs(ts.Minutes);
                                    secs = Math.Abs(ts.Seconds);
                                    datarow["ACTUALDURATION"] = days + ":" + hours.ToString("00") + ":" + mins.ToString("00") + ":" + secs.ToString("00");
                                    if (days > 0 || hours > 0 || mins > 0 || secs > 0)
                                    {
                                        datarow["STATUS"] = "Completed Normally";
                                    }
                                }
                                else
                                {
                                    datarow["ACTUALDURATION"] = DBNull.Value;
                                    datarow["STATUS"] = DBNull.Value;
                                }
                            }

                            //dr["EndTime"] = "";
                            dt.Rows.Add(datarow);

                        }
                    }
                    i++;
                    if (i > numberOfReports)
                    { 
                        break;
                    }
                }

            }
            dt = Consumeredge_TableauExtract_new(driver, dt, ReportName[4], appid);// newly added
            return dt;
        }
        public static void QlikSense(int appid, DataTable dt)
        {
            String[] filters = new String[50];
            string ntid = "", password = "", url = "";
            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {
                con.Open();
                SqlCommand getappids = new SqlCommand("select [NTID], CAST([PASSWORD] AS VARCHAR(MAX)) as [Password], [AGENTS], [ENVIRONMENT],[URL] from [dbo].[CONN_DETAILS] where APPID=" + appid + "and Source ='QlikSense'", con);
                SqlDataReader dr = getappids.ExecuteReader();
                while (dr.Read())
                {
                    ntid = dr["NTID"].ToString();
                    password = dr["Password"].ToString();
                    url = dr["URL"].ToString();

                }
                con.Close();
            }
            using (SqlConnection databasecon = new SqlConnection(sqlconnstring))
            {
                databasecon.Open();
                string sqlcommand = "";
                if (appid != 8)
                {
                    sqlcommand = "Select START_JOB from SLT_APP_CONFIGURATION where appid=" + appid;
                }
                else
                {
                    sqlcommand = "Select START_JOB from SLT_APP_CONFIGURATION where category='Qliksense- Extractor'  and  appid=" + appid;
                }
                SqlCommand getjobnames = new SqlCommand(sqlcommand, databasecon);
                SqlDataReader dr = getjobnames.ExecuteReader();
                int i = 0;
                while (dr.Read())
                {
                    filters[i] = dr["START_JOB"].ToString();
                    i++;
                }
                databasecon.Close();
            }

            DateTime productionDate = DateTime.Today.AddDays(-1);
            if (appid == 8)
            {
                if (DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
                {
                    productionDate = DateTime.Now.AddDays(-1);
                }
                if (DateTime.Now.DayOfWeek == DayOfWeek.Sunday)
                {
                    productionDate = DateTime.Now.AddDays(-2);
                }
                if (DateTime.Now.DayOfWeek == DayOfWeek.Monday)
                {
                    productionDate = DateTime.Now.AddDays(-3);
                }
                if (DateTime.Now.DayOfWeek == DayOfWeek.Tuesday)
                {
                    productionDate = DateTime.Now.AddDays(-4);
                }
                if (DateTime.Now.DayOfWeek == DayOfWeek.Wednesday)
                {
                    productionDate = DateTime.Now.AddDays(-5);
                }
                if (DateTime.Now.DayOfWeek == DayOfWeek.Thursday)
                {
                    productionDate = DateTime.Now.AddDays(-6);
                }
                if (DateTime.Now.DayOfWeek == DayOfWeek.Friday)
                {
                    productionDate = DateTime.Now.AddDays(-7);
                }
            }
            //if (appid == 8)
            //{
            //    ntid = "msharm54";
            //    password = "Thanks4health";
            //    url = "https://iqlikqa.jnj.com/qmc/";

            //}

            int hours = 0, mins = 0, days = 0, secs = 0;
            OpenQA.Selenium.IWebDriver driver = null;


            driver = new OpenQA.Selenium.Firefox.FirefoxDriver();
            driver.Navigate().GoToUrl(url);
            Task.Delay(10000).Wait();

            AutoItX3 autoIT = new AutoItX3();
            autoIT.Send(ntid);
            Task.Delay(1000).Wait();

            autoIT.Send("{TAB}");

            autoIT.Send(password);
            Task.Delay(1000).Wait();

            driver.FindElement(By.XPath("/html/body/div/div[2]/div[2]/form/a")).Click();

            Task.Delay(40000).Wait();

            driver.FindElement(By.XPath("//*[@id='qmc.stage']/div[1]/qmc-home-tiles/a[1]/div[1]")).Click();
            Task.Delay(5000).Wait();

            //driver.FindElement(By.XPath("//*[@id='col:0']/div/qmc-button/button/i")).Click();
            //Task.Delay(5000).Wait();


            //     string filtersstring="";



            //Microsoft.Office.Interop.Excel.Workbook excelWorkbook = ExcelApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            //string currentSheet = "Sheet1";
            //Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheets.get_Item(currentSheet);
            //var cell = (Microsoft.Office.Interop.Excel.Range)excelWorksheet.Cells[10, 2];

            //string path = @"C:\Users\lgupta3\Desktop\Automation SLT\SiteNavigation_Backup_fullcode\bin\Debug\IQlik Job Names.xlsx";

            //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(path);
            //Microsoft.Office.Interop.Excel.Worksheet excelSheet = wb.ActiveSheet;

            ////Read the first cell
            //if (appid == Convert.ToInt32(excelSheet.Cells[2, 1].Value.ToString()))
            //{            
            //filtersstring = excelSheet.Cells[2, 2].Value.ToString();
            //}
            //if (appid == Convert.ToInt32(excelSheet.Cells[3, 1].Value.ToString()))
            //{
            //    filtersstring = excelSheet.Cells[3, 2].Value.ToString();
            //}
            //if (appid == Convert.ToInt32(excelSheet.Cells[4, 1].Value.ToString()))
            //{
            //    filtersstring = excelSheet.Cells[4, 2].Value.ToString();
            //}

            //wb.Close();

            //if (filtersstring != "")
            //{
            //    filters = filtersstring.Split(',');
            //    for (int i = 0; i < filters.Length; i++)
            //    {
            //        filters[i] = filters[i].Trim();
            //    }

            //}

            foreach (var filter in filters)
            {
                if (filter != null)
                {
                    driver.FindElement(By.XPath("//*[@id='qmc']/qmc-framework/qmc-breadcrumb")).Click();

                    driver.FindElement(By.XPath("//*[@id='col:0']/div/qmc-button/button/i")).Click();
                    Task.Delay(1000).Wait();
                    driver.FindElement(By.XPath("html/body/qmc-table-column-filter-popover/qmc-popover/div/div[2]/div/div/qmc-table-column-filter-text/div/div/input")).Clear();
                    // driver.FindElement(By.XPath("html/body/qmc-table-column-filter-popover/qmc-popover/div/div[2]/div/div/qmc-table-column-filter-text/div/div/input")).SendKeys(filter);
                    autoIT.Send(filter);
                    Task.Delay(2000).Wait();
                    IWebElement table = driver.FindElement(By.ClassName("qmc-table-rows"));
                    List<IWebElement> tablerow = table.FindElements(By.TagName("tr")).ToList();

                    int i = 1;
                    foreach (var tr in tablerow)
                    {
                        DataRow datarow = dt.NewRow();


                        if (i > 15 & tablerow.Count > i - 3)
                        {
                            try
                            {
                                int j = i + 3;

                                IWebElement webelement = driver.FindElement(By.XPath("html/body/div/qmc-framework/div/div/qmc-table/qmc-table-rows/div/table/tbody/tr[" + j + "]/td/div"));

                                Actions actions = new Actions(driver);

                                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", webelement);
                                //Thread.sleep(500); //not sure why the sleep was needed, but it was needed or it wouldnt work :(
                                driver.FindElement(By.XPath("html/body/div/qmc-framework/div/div/qmc-table/qmc-table-rows/div/table/tbody/tr[" + j + "]/td/div")).Click();

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex);
                            }
                        }

                        driver.FindElement(By.XPath("html/body/div/qmc-framework/div/div/qmc-table/qmc-table-rows/div/table/tbody/tr[" + i + "]/td/div")).Click();
                        string jobname = driver.FindElement(By.XPath("html/body/div/qmc-framework/div/div/qmc-table/qmc-table-rows/div/table/tbody/tr[" + i + "]/td/div")).Text.ToString();
                        driver.FindElement(By.XPath("//*[@id='qmc.stageheader.task']/div[2]")).Click();
                        string status = driver.FindElement(By.XPath("html/body/div/qmc-framework/div/div/qmc-table/qmc-table-rows/div/table/tbody/tr[" + i + "]/td[5]/div/span")).Text.ToString();
                        string start_time_data;
                        driver.FindElement(By.XPath("//*/table[contains(@class, 'qmc-table-rows')]/tbody/tr[" + i + "]/td[5]/i")).Click();
                        string start_time_path, end_time_path;
                        Task.Delay(2000).Wait();
                        if (status == "Success")
                        {
                            start_time_path = "html/body/qmc-task-log-dialog/qmc-popover/div/div[2]/qmc-task-execution-info/div[2]/div/div[5]/div[1]";
                            end_time_path = "html/body/qmc-task-log-dialog/qmc-popover/div/div[2]/qmc-task-execution-info/div[2]/div/div[1]/div[1]";
                        }
                        else
                        {

                            start_time_path = "html/body/qmc-task-log-dialog/qmc-popover/div/div[2]/qmc-task-execution-info/div[2]/div/div[7]/div[1]";
                            end_time_path = "html/body/qmc-task-log-dialog/qmc-popover/div/div[2]/qmc-task-execution-info/div[2]/div/div[1]/div[1]";
                        }
                        if (CheckElementAvailable(driver, end_time_path))
                        {
                            bool starttimeelementpresent = CheckElementAvailable(driver, start_time_path);
                            if (starttimeelementpresent == true)
                            {

                                start_time_data = driver.FindElement(By.XPath(start_time_path)).Text;
                            }
                            else
                            {
                                start_time_path = "html/body/qmc-task-log-dialog/qmc-popover/div/div[2]/qmc-task-execution-info/div[2]/div/div[6]/div[1]";
                                start_time_data = driver.FindElement(By.XPath(start_time_path)).Text;
                            }
                            start_time_data = start_time_data.Substring(0, start_time_data.Length - 4);
                            DateTime starttime = Convert.ToDateTime(start_time_data);
                            TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                            DateTime starteasternTime = TimeZoneInfo.ConvertTimeFromUtc(starttime, easternZone);
                            datarow["ACTUALSTARTTIME"] = starteasternTime;
                            string end_time_data = driver.FindElement(By.XPath(end_time_path)).Text;
                            end_time_data = end_time_data.Substring(0, end_time_data.Length - 4);
                            DateTime endtime = Convert.ToDateTime(end_time_data);
                            DateTime endeasterntime = TimeZoneInfo.ConvertTimeFromUtc(endtime, easternZone);
                            TimeSpan ts = starteasternTime - endeasterntime;
                            hours = Math.Abs(ts.Hours);
                            days = Math.Abs(ts.Days);
                            mins = Math.Abs(ts.Minutes);
                            secs = Math.Abs(ts.Seconds);
                            string duration = days + ":" + hours.ToString("00") + ":" + mins.ToString("00") + ":" + secs.ToString("00");
                            datarow["ACTUALDURATION"] = duration;

                        }
                        if (status == "Success")
                        {
                            datarow["STATUS"] = "Completed Normally";

                        }
                        else
                        {
                            datarow["STATUS"] = status;
                        }
                        datarow["PRODUCTION_DATE"] = productionDate;
                        datarow["SRC_DATA"] = "QlikSense";
                        datarow["APPID"] = appid;
                        datarow["JOB"] = jobname;
                        datarow["JOBGROUP"] = jobname;
                        dt.Rows.Add(datarow);
                        i++;


                    }
                }

            }
            BulkUpload(dt);
            driver.Quit();
        }
        public static bool CheckElementAvailable(OpenQA.Selenium.IWebDriver driver, string elementpath)
        {
            Boolean present = false;
            try
            {
                driver.FindElement(By.XPath(elementpath));
                present = true;
            }
            catch
            {
                present = false;
            }
            return present;
        }
        public static void GlobalSales(int appid, DataTable dt)
        {
            DateTime productionDate = DateTime.Today;
            string username = "", password = "", server = "";
            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {
                con.Open();
                SqlCommand getappids = new SqlCommand("select [NTID], CAST([PASSWORD] AS VARCHAR(MAX)) as [Password], [AGENTS], [ENVIRONMENT],[URL] from [dbo].[CONN_DETAILS] where APPID=" + appid + " and Source='Redshift'", con);
                SqlDataReader dr = getappids.ExecuteReader();
                while (dr.Read())
                {
                    username = dr["NTID"].ToString();
                    password = dr["Password"].ToString();
                    server = dr["URL"].ToString();

                }
                con.Close();
            }
        // Server, e.g. "examplecluster.xyz.us-west-2.redshift.amazonaws.com"
         
        // Port
         string port = "5439";
        // MasterUserName
         
        // MasterUserPassword
       
        // DBName, e.g. ""
         string DBName = "globalsales_prod";
        // Create the ODBC connection string.
        /*Redshift ODBC Driver - 64 bits                
        string connString = "Driver={Amazon Redshift (x64)};" +
            String.Format("Server={0};Database={1};" +
            "UID={2};PWD={3};Port={4};SSL=true;Sslmode=Require",
            server, DBName, masterUsername,
            masterUserPassword, port);
        */
        //Redshift ODBC Driver - 32 bits
         string connString = "Driver={Amazon Redshift (x86)};" +
            String.Format("Server={0};Database={1};" +
            "UID={2};PWD={3};Port={4};SSL=true;Sslmode=Require",
            server, DBName, username,
            password, port);
             OdbcConnection conn = new OdbcConnection(connString);
             DataTable reportlisttable = new DataTable();
             string backendReportQuery = "Select etl_region_cd, CONVERT_TIMEZONE('UTC+1','EST',load_start_time) as load_start_time,CONVERT_TIMEZONE('UTC+1','EST',load_end_time) as load_end_time , latest_data_dt from globalsales.tl_gbl_load_status;";
             conn.Open();
             DataSet ds = new DataSet();
             
             //Excecuting the query to fect the TO, CC, BCC, FROM, Imagename, PDF attachment details for all the tableau report.
             OdbcDataAdapter da = new OdbcDataAdapter(backendReportQuery, conn);
             da.Fill(ds);
             reportlisttable = ds.Tables[0];
             conn.Close();
             foreach (DataRow row in reportlisttable.Rows)
             {
                 int hours=0, days=0, mins=0, secs=0;
                 DataRow datarow = dt.NewRow();
                
                 datarow["SRC_DATA"] = "Redshift";
                 datarow["APPID"] = appid;
                 datarow["JOB"] = row["etl_region_cd"];
                 datarow["JOBGROUP"] = row["etl_region_cd"];
                 datarow["PRODUCTION_DATE"] = productionDate;
                 DateTime starttime = Convert.ToDateTime(row["load_start_time"]);
                 DateTime endtime = Convert.ToDateTime(row["load_end_time"]);
                 if (starttime.Date == DateTime.Today.Date)
                 {
                      datarow["ACTUALSTARTTIME"] = row["load_start_time"];
                     TimeSpan ts = starttime - endtime;
                     hours = Math.Abs(ts.Hours);
                     days = Math.Abs(ts.Days);
                     mins = Math.Abs(ts.Minutes);
                     secs = Math.Abs(ts.Seconds);
                     string duration = days + ":" + hours.ToString("00") + ":" + mins.ToString("00") + ":" + secs.ToString("00");
                     datarow["ACTUALDURATION"] = duration;


                     DateTime lastestdate = Convert.ToDateTime(row["latest_data_dt"]);
                     TimeSpan time = lastestdate - productionDate.AddDays(-1);
                     days = Math.Abs(time.Days);

                     if (days == 0 & starttime < endtime)
                     {
                         datarow["STATUS"] = "Completed Normally";
                     }
                     else
                     {
                         if (days == 0 & starttime > endtime)
                             datarow["STATUS"] = "WIP";
                         else
                         {
                             datarow["STATUS"] = "Completed Abnormally";
                         }
                     }
                 }
                 else {
                     
                     datarow["STATUS"] = "";
                 }
                 dt.Rows.Add(datarow);
             }


             BulkUpload(dt);
        }
        static void Tac(int appid, DataTable dt)
        {
            string userName = "", password = "", url = "";
            string[] arr = new string[20];
            using (SqlConnection con = new SqlConnection(sqlconnstring))
            {

                con.Open();
                SqlCommand getappids = new SqlCommand("select [NTID], CAST([PASSWORD] AS VARCHAR(MAX)) as [Password], [AGENTS], [ENVIRONMENT],[URL] from [dbo].[CONN_DETAILS] where APPID=" + appid + "and Source ='TAC'", con);
                SqlDataReader sdr = getappids.ExecuteReader();
                while (sdr.Read())
                {
                    userName = sdr["NTID"].ToString();
                    password = sdr["Password"].ToString();
                    url = @sdr["URL"].ToString();

                }
                con.Close();
                userName = "ssakthi2@its.jnj.com";
                password = "Sep@2017";
                url = @"https://awsaunsgp1220.jnj.com:8443/talend";
            }
            OpenQA.Selenium.IWebDriver driver = null;

            driver = new OpenQA.Selenium.Chrome.ChromeDriver();
            driver.Navigate().GoToUrl(url);
            Task.Delay(10000).Wait();
            driver.FindElement(By.Id("idLoginInput")).SendKeys(userName);
            driver.FindElement(By.Id("idLoginPasswordInput")).SendKeys(password);
            driver.FindElement(By.Id("idLoginButton")).Click();
            Task.Delay(5000).Wait();
            string forcedLogout = "html/body/div/div/div/div[2]/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/em/button";
            if (CheckElementAvailable(driver, forcedLogout))
            {
                driver.FindElement(By.XPath(forcedLogout)).Click();
                Task.Delay(2000).Wait();
                driver.FindElement(By.Id("idLoginButton")).Click();
                Task.Delay(30000).Wait();
            }
            else
            {
                Task.Delay(25000).Wait();
            }
            driver.FindElement(By.XPath("//*[@id='x-auto-17__!!!menu.executionTasks.element!!!']/span[2]/span")).Click();
            Task.Delay(5000).Wait();
            //List<IWebElement> tables;

            //IWebElement dataDiv = driver.FindElement(By.Id("x-auto-320-gp-groupid-0-bd"));
            //tables = dataDiv.FindElements(By.TagName("Table")).ToList();
            string numberOfJobsString = driver.FindElement(By.XPath("html/body/div/div/div[2]/div/div[3]/div/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div[2]/div/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/div")).Text;
            Char delimiter = ' ';
            String[] substrings = numberOfJobsString.Split(delimiter);
            int numberOfJobs = Convert.ToInt32(substrings[substrings.Length - 1]);
            int numberOfElements = 0;
            bool elementsOnPage = true;
            for (int i = 0; i < numberOfJobs; i++)
            {
                numberOfElements++;
                if (numberOfElements > 10)
                {
                    numberOfElements = 1;
                    driver.FindElement(By.XPath("html/body/div/div/div[2]/div/div[3]/div/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div[2]/div/table/tbody/tr/td/table/tbody/tr/td[8]/table/tbody/tr[2]/td[2]/em/button")).Click();
                    Task.Delay(5000).Wait();
                }
                string jobNamePath = "html/body/div/div/div[2]/div/div[3]/div/div[4]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/div[2]/div[2]/table/tbody/tr/td[6]/div/span";
                string reportDetailsButton = "";
                if (CheckElementAvailable(driver, jobNamePath))
                {
                    jobNamePath = "html/body/div/div/div[2]/div/div[3]/div/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/div[2]/div[2]/table/tbody/tr/td[6]/div/span";
                    reportDetailsButton = "html/body/div/div/div[2]/div/div[3]/div/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/div[2]/div[" + numberOfElements + "]/table/tbody/tr/td[10]/div/div/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/em/button";
                }
                else
                {
                    jobNamePath = "html/body/div/div/div[2]/div/div[3]/div/div[4]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/div[2]/div[2]/table/tbody/tr/td[6]/div/span";
                    reportDetailsButton = "html/body/div/div/div[2]/div/div[4]/div/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/div[2]/div[" + numberOfElements + "]/table/tbody/tr/td[10]/div/div/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/em/button";
                }
                string jobName = driver.FindElement(By.XPath(jobNamePath)).Text;
                driver.FindElement(By.XPath(reportDetailsButton)).Click();
                Task.Delay(5000).Wait();

                //html/body/div/div/div[2]/div/div[3]/div/div[4]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/div[2]/div[1 ]/table/tbody/tr/td[10]/div/div/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/em/button
               
                if (elementsOnPage)
                {
                    driver.FindElement(By.XPath("html/body/div/div/div[2]/div/div[3]/div/div[4]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[3]/div/table/tbody/tr/td/table/tbody/tr/td[11]/div/input")).Click();
                    driver.FindElement(By.XPath("html/body/div/div/div[2]/div/div[3]/div/div[4]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[3]/div/table/tbody/tr/td/table/tbody/tr/td[11]/div/input")).Clear();
                    driver.FindElement(By.XPath("html/body/div/div/div[2]/div/div[3]/div/div[4]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[3]/div/table/tbody/tr/td/table/tbody/tr/td[11]/div/input")).SendKeys("1");
                    AutoItX3 autoIT = new AutoItX3();
                    autoIT.Send("{ENTER}");
                    elementsOnPage = false;
                }
                Task.Delay(3000).Wait();
                string jobStartTimePath = "html/body/div/div/div[2]/div[1]/div[3]/div/div[4]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div/div[2]/div/div[1]/table/tbody/tr/td[9]/div/span";
                string jobEndTimePath = "html/body/div/div/div[2]/div[1]/div[3]/div/div[4]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div/div[2]/div/div[1]/table/tbody/tr/td[10]/div/span";
                if (CheckElementAvailable(driver, jobStartTimePath))
                {
                    string jobStartTime = driver.FindElement(By.XPath(jobStartTimePath)).Text.ToString();
                    string jobEndTime = driver.FindElement(By.XPath(jobEndTimePath)).Text.ToString();
                    string status = driver.FindElement(By.XPath("html/body/div/div/div[2]/div[1]/div[3]/div/div[4]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div/div[2]/div/div[1]/table/tbody/tr/td[3]/div/span")).Text.ToString();
                    if (status.ToLower() == "ok")
                    {
                        status = "Completed Normally";
                    }
                    else
                    {
                        status = "Completed Abnomally";
                    }
                    DataRow datarow = dt.NewRow();
                    datarow["ACTUALSTARTTIME"] = jobStartTime;
                    datarow["STATUS"] = status;
                    datarow["PRODUCTION_DATE"] = DateTime.Today.AddDays(-1);
                    datarow["SRC_DATA"] = "TAC";
                    datarow["APPID"] = appid;
                    datarow["JOB"] = jobName;
                    datarow["JOBGROUP"] = jobName;
                    datarow["SRC_DATA"] = "TAC";
                    dt.Rows.Add(datarow);
                }
                driver.FindElement(By.XPath("//*[@id='x-auto-17__!!!menu.executionTasks.element!!!']/span[2]/span")).Click();
                Task.Delay(3000).Wait();
                ////*[@id="x-auto-1128"]/div/span
                //*[@id="x-auto-1073"]/div


            }
            driver.FindElement(By.XPath("html/body/div/div/div[2]/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/em/button")).Click();
            Task.Delay(2000).Wait();
            driver.Quit();
        }
    }
}
