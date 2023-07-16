using System.Collections.Generic;
using System.Web.Http;
using System.Data;
using System.Data.SqlClient;
using PunchSystem.Models;
using System;

namespace PunchSystem.Controllers
{
    public class PunchController : ApiController
    {
        // 資料庫連線
        SqlConnection connString = new SqlConnection();
        string conSource = System.Web.Configuration.WebConfigurationManager.AppSettings["Source"]; //主機
        string conDBName = System.Web.Configuration.WebConfigurationManager.AppSettings["DBName"]; //資料庫
        string conUId = System.Web.Configuration.WebConfigurationManager.AppSettings["Id"];  //登入資料庫帳號
        string conUPwd = System.Web.Configuration.WebConfigurationManager.AppSettings["Pwd"];//登入資料庫密碼
        DateTime today = DateTime.Today;

        // GET: api/Punch/?employeeNum={employeeNum 打卡功能
        public string Punch(string employeeNum)
        {
            Dictionary<string, string> res = new Dictionary<string, string>();
            try
            { 
                connString.ConnectionString = 
                    @"Data Source=" + conSource + ";Initial Catalog=" + conDBName + ";Persist Security Info=True;user ID=" + conUId + ";Password=" + conUPwd + " ";
                connString.Open();

                // 檢查是否有打過卡
                string sqlstr = string.Format("select count(employeeNumber) from ClockinData where employeeNumber='{0}' AND left(clockIn, '10')={1}", employeeNum, today);
                SqlCommand checkData = new SqlCommand(sqlstr, connString);
                Int32 count = Convert.ToInt32(checkData.ExecuteScalar());
            
                if (count < 0)
                { // 今天無打卡紀錄，新增上班打卡紀錄
                    SqlCommand cmdPunchIn = new SqlCommand("insert into ClockinData([employeeNumber],[clockIn]) "
                        + "values('" + employeeNum + "' ,getdate())", connString);
                    cmdPunchIn.ExecuteNonQuery();
                }
                else
                { // 已有打卡紀錄，故新增下班打卡紀錄
                    SqlCommand cmdPunchOut = new SqlCommand("update ClockinData "
                        + "set values employeeNumber='" + employeeNum + "' , clockOut=getdate())", connString);
                    cmdPunchOut.ExecuteNonQuery();
                }
                connString.Close();
                res.Add("success", "true");
                res.Add("message", "打卡成功");
                return res.ToString();
            }
            catch (Exception ex)
            {
                connString.Close();
                res.Add("success", "false");
                res.Add("message", ex.Message);
                return res.ToString();
            }
        }

        // POST: api/makeupPunchIn/?employeeNum={employeeNum}&dateInfo={dateInfo} 補打上班卡功能
        public string makeupPunchIn(string employeeNum, string dateInfo)
        {
            Dictionary<string, string> res = new Dictionary<string, string>();
            try
            {
                connString.ConnectionString =
                    @"Data Source=" + conSource + ";Initial Catalog=" + conDBName + ";Persist Security Info=True;user ID=" + conUId + ";Password=" + conUPwd + " ";
                connString.Open();

                SqlCommand cmdMakeupPunchIn = new SqlCommand("update ClockinData "
                        + "set values clockIn=getdate())"
                        + "where employeeNumber='" + employeeNum + " AND left(clockOut, '10')='" + dateInfo + "'", connString);
                cmdMakeupPunchIn.ExecuteNonQuery();
                connString.Close();
                res.Add("success", "true");
                res.Add("message", "補打卡成功");
                return res.ToString();
            }
            catch (Exception ex)
            {
                connString.Close();
                res.Add("success", "false");
                res.Add("message", ex.Message);
                return res.ToString();
            }

        }

        // POST: api/makeupPunchOut/?employeeNum={employeeNum}&dateInfo={dateInfo} 補打下班卡功能
        public string makeupPunchOut(string employeeNum, string dateInfo)
        {
            Dictionary<string, string> res = new Dictionary<string, string>();
            try
            {
                connString.ConnectionString =
                    @"Data Source=" + conSource + ";Initial Catalog=" + conDBName + ";Persist Security Info=True;user ID=" + conUId + ";Password=" + conUPwd + " ";
                connString.Open();

                SqlCommand cmdMakeupPunchOut = new SqlCommand("update ClockinData "
                        + "set values clockOut=getdate())"
                        + "where employeeNumber='" + employeeNum + " AND left(clockIn, '10')='" + dateInfo + "'", connString);
                cmdMakeupPunchOut.ExecuteNonQuery();
                connString.Close();
                res.Add("success", "true");
                res.Add("message", "補打卡成功");
                return res.ToString();
            }
            catch (Exception ex)
            {
                connString.Close();
                res.Add("success", "false");
                res.Add("message", ex.Message);
                return res.ToString();
            }

        }
        
        // POST: api/theDayData/ 所有員工當日資訊
        public IEnumerable<WholeDates> theDayData()
        {
            Dictionary<string, string> res = new Dictionary<string, string>();
            try
            {
                connString.ConnectionString =
                    @"Data Source=" + conSource + ";Initial Catalog=" + conDBName + ";Persist Security Info=True;user ID=" + conUId + ";Password=" + conUPwd + " ";
                connString.Open();

                string sqlstr = string.Format("select * from ClockinData where left(clockIn, 10) = {0} or left(clockOut, 10) = {0}", today);
                SqlCommand cmdDtas = new SqlCommand(sqlstr, connString);
                SqlDataReader dr = cmdDtas.ExecuteReader();
                DataSet ds = new DataSet();
                List<WholeDates> list = new List<WholeDates>();

                if (dr.Read())
                {
                    dr.Dispose();
                    SqlDataAdapter adapter = new SqlDataAdapter(cmdDtas);
                    adapter.Fill(ds);
                    
                    foreach (DataRow r in ds.Tables[0].Rows)
                    {
                        DateTime tCheckIn = (DateTime)r[1];
                        DateTime tCheckOut = (DateTime)r[2];
                        // 計算時間差
                        TimeSpan timeDiff = tCheckOut - tCheckIn;
                        // 以小時表示並顯示到小數點第2位
                        double hours = timeDiff.TotalHours - 1.5;
                        string formattedHours = hours.ToString("F2");
                        
                        list.Add(new WholeDates { employeeNumber = r[0].ToString(), clockIn = r[1].ToString(), clockOut = r[2].ToString(), restTime = "1.5", totalTime = formattedHours });
                    }
                }
                dr.Close();
                connString.Close();

                res.Add("success", "true");
                res.Add("message", "所有員工當日資訊");
                return list;
            }
            catch (Exception ex)
            {
                return new List<WholeDates>();
            }

        }

        // POST: api/theDayData/date={date} 指定日期當天所有員工資訊
        public IEnumerable<WholeDates> someDayData(string date) 
        {
            Dictionary<string, string> res = new Dictionary<string, string>();
            try
            {
                connString.ConnectionString =
                    @"Data Source=" + conSource + ";Initial Catalog=" + conDBName + ";Persist Security Info=True;user ID=" + conUId + ";Password=" + conUPwd + " ";
                connString.Open();

                string sqlstr = string.Format("select * from ClockinData where left(clockIn, 10) = {0} or left(clockOut, 10) = {0}", date);
                SqlCommand cmdDtas = new SqlCommand(sqlstr, connString);
                SqlDataReader dr = cmdDtas.ExecuteReader();
                DataSet ds = new DataSet();
                List<WholeDates> list = new List<WholeDates>();

                if (dr.Read())
                {
                    dr.Dispose();
                    SqlDataAdapter adapter = new SqlDataAdapter(cmdDtas);
                    adapter.Fill(ds);
                    
                    foreach (DataRow r in ds.Tables[0].Rows)
                    {
                        DateTime tCheckIn = (DateTime)r[1];
                        DateTime tCheckOut = (DateTime)r[2];
                        // 計算時間差
                        TimeSpan timeDiff = tCheckOut - tCheckIn;
                        // 以小時表示並顯示到小數點第2位
                        double hours = timeDiff.TotalHours - 1.5;
                        string formattedHours = hours.ToString("F2");

                        list.Add(new WholeDates { employeeNumber = r[0].ToString(), clockIn = r[1].ToString(), clockOut = r[2].ToString(), restTime = "1.5", totalTime = formattedHours });
                    }
                }
                dr.Close();
                connString.Close();

                return list;
            }
            catch (Exception ex)
            {
                return new List<WholeDates>();
            }

        }

        // POST: api/unPunchOutList/?sDate={sDate}&eDate={eDate} 指定日期區間未打下班卡的員工清單
        public IEnumerable<CheckInData> unPunchOutList(string sDate, string eDate)
        {
            Dictionary<string, string> res = new Dictionary<string, string>();
            try
            {
                connString.ConnectionString =
                    @"Data Source=" + conSource + ";Initial Catalog=" + conDBName + ";Persist Security Info=True;user ID=" + conUId + ";Password=" + conUPwd + " ";
                connString.Open();
                
                string sqlstr = string.Format("select * from ClockinData where left(chechIn, 10) between '{0}' and '{1}' AND chechOut is null", sDate, eDate);
                SqlCommand cmdDtas = new SqlCommand(sqlstr, connString);
                SqlDataReader dr = cmdDtas.ExecuteReader();
                DataSet ds = new DataSet();
                List<CheckInData> list = new List<CheckInData>();
                if (dr.Read())
                {
                    dr.Dispose();
                    SqlDataAdapter adapter = new SqlDataAdapter(cmdDtas);
                    adapter.Fill(ds);

                    foreach (DataRow r in ds.Tables[0].Rows)
                    {
                        list.Add(new CheckInData { employeeNumber = r[0].ToString(), clockIn = r[1].ToString(), clockOut = r[2].ToString() });
                    }
                }

                dr.Close();
                connString.Close();

                return list;
            }
            catch (Exception ex)
            {
                return new List<CheckInData>();
            }

        }

        // POST: api/earlyPunchData 指定日期，當天前五名最早打卡上班的員工
        public IEnumerable<CheckInData> earlyPunchData(string date)
        {
            Dictionary<string, string> res = new Dictionary<string, string>();
            try
            {
                connString.ConnectionString =
                    @"Data Source=" + conSource + ";Initial Catalog=" + conDBName + ";Persist Security Info=True;user ID=" + conUId + ";Password=" + conUPwd + " ";
                connString.Open();
                
                string sqlstr = string.Format("select TOP 5 * from ClockinData where left(chechIn, 10) = '{0}' order by checkIn", date);
                SqlCommand cmdDtas = new SqlCommand(sqlstr, connString);
                SqlDataReader dr = cmdDtas.ExecuteReader();
                DataSet ds = new DataSet();
                List<CheckInData> list = new List<CheckInData>();
                if (dr.Read())
                {
                    dr.Dispose();
                    SqlDataAdapter adapter = new SqlDataAdapter(cmdDtas);
                    adapter.Fill(ds);

                    foreach (DataRow r in ds.Tables[0].Rows)
                    {
                        list.Add(new CheckInData { employeeNumber = r[0].ToString(), clockIn = r[1].ToString(), clockOut = r[2].ToString() });
                    }
                }

                return list;
            }
            catch (Exception ex)
            {
                return new List<CheckInData>();
            }

        }
    }
}
