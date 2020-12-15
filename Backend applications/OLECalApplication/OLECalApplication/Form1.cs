using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace OLECalApplication
{
    public partial class Form1 : Form
    {
        i_facility_tsalEntities1 db = new i_facility_tsalEntities1();
        string reportpath = ConfigurationManager.AppSettings["path"];
        string path1 = "";
        public Form1()
        {
            //InitializeComponent();
            var nwDetails1 = db.tblnetworkdetailsforddls.Where(m => m.IsDeleted == 0 && m.NPFDDLID == 10).FirstOrDefault(); // INCremental DDL
             path1 = nwDetails1.Path;
            string username1 = nwDetails1.UserName;
            string password1 = nwDetails1.Password;
            string domainname1 = nwDetails1.DomainName;
            path1 = @"C:\TVS_batery\Source";
            try
            {
                //InitializeComponent();
                InitializeComponent(path1, username1, password1, domainname1);
            }
            catch (Exception e)
            {
                //MessageBox.Show("Path Error: " + e);
                IntoFile("Authentiction Error |" + e.ToString());
            }

            //Timer MyTimer = new Timer();
            ////MyTimer.Interval = (20 * 1000); // 20 seconds
            //MyTimer.Interval = (60 * 1000); // 1 minute
            //MyTimer.Tick += new EventHandler(MyTimer_Tick);
            //MyTimer.Start();
        }

        private void MyTimer_Tick(object sender, EventArgs e)
        {
            string correctedDate = GetCorrectedDate();
            try
            {
                InsertOLEtable(correctedDate);
            }
            catch (Exception)
            {

            }
        }

        private string GetCorrectedDate()
        {
            string CorrectedDate = "";
          
            tbldaytiming StartTime1 = db.tbldaytimings.Where(m=>m.IsDeleted == 0).FirstOrDefault();
            TimeSpan Start = StartTime1.StartTime;
            if (Start <= DateTime.Now.TimeOfDay)
            {
                CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
            }
            else
            {
                CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            }
            return CorrectedDate;
        }

        private void fileSystemWatcher1_Created(object sender, FileSystemEventArgs e)
        {
            try
            {

                string rootFolderPath = e.FullPath;
              
                string filename = Path.GetFileName(rootFolderPath);
                string fileExt = Path.GetExtension(rootFolderPath);

                if ((fileExt == ".xlsx"))
                {

                    try
                    {

                        InsertHrDetails(path1);
                        
                    }
                    catch (Exception ex)
                    {
                        IntoFile(ex.ToString());
                    }
                    Timer MyTimer = new Timer();
                    //MyTimer.Interval = (20 * 1000); // 20 seconds
                    MyTimer.Interval = (60 * 1000); // 1 minute
                    MyTimer.Tick += new EventHandler(MyTimer_Tick);
                    MyTimer.Start();

                }
                else
                {
                    // file Error
                    IntoFile("File Format Error ");
                }
            }
            catch (Exception exc)
            {
                IntoFile("IN FileWatcher Section: " + exc);
            }

        }

        

        public void InsertHrDetails(string path)
        {
            try
            {
                DataSet ds = new DataSet();
                DirectoryInfo di = new DirectoryInfo(path);
                FileInfo[] subFiles = di.GetFiles();
                if (subFiles.Length > 0)
                {
                    DataTable dt = new DataTable();
                    string filepath = path + "//" + subFiles[0];
                    //Thread.Sleep(1 * 60 * 1000);
                    dt = GetDataTableFromExcel(filepath);

                    //DataView dataview = dt.DefaultView;
                    dt.DefaultView.Sort = "EID";
                    //dataview.Sort = Convert.ToString(dt.Rows[0]["OpId"]);
                    //DataTable dt1 = dataview.ToTable();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string opid = null;
                        opid = Convert.ToString(dt.Rows[i][1]);
                        string CorrectedDate = null;
                        CorrectedDate = Convert.ToString(dt.Rows[i][6]);
                        string stratDate = Convert.ToDateTime( CorrectedDate).ToString("yyyy-MM-dd");
                        string startTime = Convert.ToString(dt.Rows[i][7]);
                        if (startTime == "0" || startTime == "")
                            startTime = "00:00";
                        string StartDateTime = stratDate + " " + startTime;
                        string endTime = Convert.ToString(dt.Rows[i][8]);
                        
                        string endDate = CorrectedDate;
                       

                        if (endTime == "0" || endTime == "")
                            endTime = "00:00";
                        endDate = GetEndCorrectedDate(endTime,Convert.ToDateTime(stratDate));
                        string endDateTime = endDate + " " + endTime;
                        string TotaldurInHrs = Convert.ToString(dt.Rows[i][10]);
                        double durIn = TimeSpan.Parse(TotaldurInHrs).TotalMinutes;

                        string DurationinMin = Convert.ToDateTime(endDateTime).Subtract(Convert.ToDateTime(StartDateTime)).TotalMinutes.ToString();
                        string durInMin = DurationinMin;
                       
                        DateTime CreatedOn = DateTime.Now;
                        int CreatedBy = 1;

                        if (!string.IsNullOrEmpty(opid))
                        {
                            try
                            {
                                using (MsqlConnection mc1 = new MsqlConnection())
                                {
                                    mc1.open();
                                    SqlCommand cmd2 = null;
                                    //MessageBox.Show("WO & OPNo & PartNo Values: " + WorkOrder + " " + OpNo + " " + partNo);
                                    //cmd2 = new MySqlCommand("INSERT INTO tblmachinedetails (InsertedBy,InsertedOn,IsDeleted,MachineType, MachineInvNo, IPAddress, ControllerType,MachineModel,MachineMake,ModelType,MachineDispName,IsParameters,ShopNo,IsPCB) VALUES( '" + dat + "'," + 0 + "," + 2 + ",'" + ds.Tables[0].Rows[i][0].ToString() + "','" + ds.Tables[0].Rows[i][1].ToString() + "','" + ds.Tables[0].Rows[i][2].ToString() + "','" + ds.Tables[0].Rows[i][3].ToString() + "','" + ds.Tables[0].Rows[i][4].ToString() + "','" + ds.Tables[0].Rows[i][5].ToString() + "','" + ds.Tables[0].Rows[i][6].ToString() + "','" + ds.Tables[0].Rows[i][7].ToString() + "','" + ds.Tables[0].Rows[i][8].ToString() + "','" + ds.Tables[0].Rows[i][9].ToString() + "')", mc1.msqlConnection);
                                    cmd2 = new SqlCommand("INSERT INTO " + MsqlConnection.DBSchemaName + ".tblhrdetails(opid,StartTime,EndTime,CorrectedDate,DurationInMin,CreatedOn,CreatedBy)VALUES(" + opid + ",'" + StartDateTime + "','" + endDateTime + "','" + CorrectedDate + "','" + durInMin + "','" + CreatedOn + "'," + CreatedBy + ");", mc1.msqlConnection);
                                    cmd2.ExecuteNonQuery();
                                    mc1.close();
                                }
                            }

                            catch (Exception e)
                            {
                                //continue;
                                IntoFile(e.ToString());
                            }
                        }
                    }
                }
            }

            catch (Exception ex)
            {

            }
        }

        public async void InsertOLEtable(string CorrectedDate)
        {
            int Shift = 0;
            DateTime Shiftstart = DateTime.Now;
            DateTime ShiftEnd = DateTime.Now;
            string createdOn1 = null;
            string createdOn2 = null;
            DataTable dtshift = new DataTable();
            String queryshift = "SELECT ShiftName,StartTime,EndTime FROM shift_master WHERE IsDeleted = 0";
            MsqlConnection mcp = new MsqlConnection();
            mcp.open();
            using (SqlDataAdapter dashift = new SqlDataAdapter(queryshift, mcp.msqlConnection))
            {
                dashift.Fill(dtshift);
            }
            mcp.close();

            String[] msgtime = System.DateTime.Now.TimeOfDay.ToString().Split(':');
            TimeSpan msgstime = System.DateTime.Now.TimeOfDay;
            //TimeSpan msgstime = new TimeSpan(Convert.ToInt32(msgtime[0]), Convert.ToInt32(msgtime[1]), Convert.ToInt32(msgtime[2]));
            TimeSpan s1t1 = new TimeSpan(0, 0, 0), s1t2 = new TimeSpan(0, 0, 0), s2t1 = new TimeSpan(0, 0, 0), s2t2 = new TimeSpan(0, 0, 0);
            TimeSpan s3t1 = new TimeSpan(0, 0, 0), s3t2 = new TimeSpan(0, 0, 0), s3t3 = new TimeSpan(0, 0, 0), s3t4 = new TimeSpan(23, 59, 59);
            for (int k = 0; k < dtshift.Rows.Count; k++)
            {
                if (dtshift.Rows[k][0].ToString().Contains("1"))
                {
                    String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                    s1t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                    String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
                    s1t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                }
                else if (dtshift.Rows[k][0].ToString().Contains("2"))
                {
                    String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                    s2t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                    String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
                    s2t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                }
                else if (dtshift.Rows[k][0].ToString().Contains("3"))
                {
                    String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                    s3t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                    String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
                    s3t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                }
            }
            //CorrectedDate = System.DateTime.Now.ToString("yyyy-MM-dd");
            if (msgstime >= s1t1 && msgstime < s1t2)
            {
                createdOn1 = CorrectedDate + " " + s1t1;
                createdOn2 = CorrectedDate + " " + s1t2;
                Shiftstart = Convert.ToDateTime(createdOn1);
                ShiftEnd = Convert.ToDateTime(createdOn2);
                Shift = 1;
            }
            else if (msgstime >= s2t1 && msgstime < s2t2)
            {
                createdOn1 = CorrectedDate + " " + s2t1;
                createdOn2 = CorrectedDate + " " + s2t2;
                Shiftstart = Convert.ToDateTime(createdOn1);
                ShiftEnd = Convert.ToDateTime(createdOn2);
                Shift = 2;
            }
            else if ((msgstime >= s3t1 && msgstime <= s3t4) || (msgstime >= s3t3 && msgstime < s3t2))
            {
                createdOn1 = CorrectedDate + " " + s3t1;
                createdOn2 = CorrectedDate + " " + s3t2;
                Shiftstart = Convert.ToDateTime(createdOn1);
                ShiftEnd = Convert.ToDateTime(createdOn2);
                Shift = 3;
            }
            int hmidur = 0;
            int diff1 = 0;
            int oTDet = 0;
            double green, red, yellow, blue, setup = 0, scrap = 0, NOP = 0, OperatingTime = 0, DownTimeBreakdown = 0, ROALossess = 0, AvailableTime = 0, SettingTime = 0, PlannedDownTime = 0, UnPlannedDownTime = 0;
            double SummationOfSCTvsPP = 0, MinorLosses = 0, ROPLosses = 0;
            double ScrapQtyTime = 0, ReWOTime = 0, ROQLosses = 0;

            var machinesdet = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.IsNormalWC == 1).ToList();
            foreach (var macrow in machinesdet)
            {
               // macrow.MachineID = 987;
                var lossdet = db.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.CorrectedDate == CorrectedDate && m.StartTime >= Shiftstart && m.EndTime <= ShiftEnd && m.ColorCode == "yellow" && m.MachineID == macrow.MachineID).Sum(m => m.DurationInSec);

                blue = await GetOPIDleBreakDown(CorrectedDate, macrow.MachineID, "blue");
                green = await GetOPIDleBreakDown(CorrectedDate, macrow.MachineID, "green");
                try
                {
                    //Availability
                    SettingTime = await GetSettingTime(CorrectedDate, macrow.MachineID);
                    if (SettingTime < 0)
                    {
                        SettingTime = 0;
                    }
                    ROALossess = await GetDownTimeLosses(CorrectedDate, macrow.MachineID, "ROA");
                    if (ROALossess < 0)
                    {
                        ROALossess = 0;
                    }

                    //Performance
                    SummationOfSCTvsPP = await GetSummationOfSCTvsPP(CorrectedDate, macrow.MachineID);
                    if (SummationOfSCTvsPP <= 0)
                    {
                        SummationOfSCTvsPP = 0;
                    }

                    //ROPLosses = GetDownTimeLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "ROP");
                }
                catch (Exception e)
                {

                }

                //Quality
                try
                {
                    ScrapQtyTime = await GetScrapQtyTimeOfWO(CorrectedDate, macrow.MachineID);
                    if (ScrapQtyTime < 0)
                    {
                        ScrapQtyTime = 0;
                    }
                    ReWOTime = await GetScrapQtyTimeOfRWO(CorrectedDate, macrow.MachineID);
                    if (ReWOTime < 0)
                    {
                        ReWOTime = 0;
                    }
                }
                catch (Exception e)
                {

                }

                var hmidet = db.tbllivehmiscreens.Where(m => m.CorrectedDate == CorrectedDate /*&& m.Date >= Shiftstart && m.Time <= ShiftEnd*/ && m.MachineID == macrow.MachineID).ToList();
                if(hmidet.Count!=0)
                {
                    foreach (var hmirow in hmidet)
                    {
                        string[] opid = hmirow.OperatorDet.Split(',');
                        foreach (var oprow in opid)
                        {
                            int OperatorId = Convert.ToInt32(oprow);
                            var hrdet = db.tblhrdetails.Where(m => m.CorrectedDate == CorrectedDate && m.Isdeleted == 0 && m.opid == OperatorId).Sum(m => m.DurationInMin);
                            if (hrdet != 0)
                            {
                                var oledet = db.tblolecaldetails.Where(m => m.opid == OperatorId).OrderByDescending(m => m.oleid).FirstOrDefault();
                                if (oledet != null)
                                {
                                    if (oledet.shift != Shift)
                                    {
                                        oTDet = (int)hrdet;
                                    }
                                }

                                tblolecaldetail obj = new tblolecaldetail();
                                obj.CorrectedDate = CorrectedDate;
                                obj.CreatedBy = 1;
                                obj.CreatedOn = DateTime.Now;
                                obj.Green = green;
                                obj.Blue = blue;
                                obj.ROALossess = ROALossess;
                                obj.ReWOTime = ReWOTime;
                                obj.ScrapQtyTime = ScrapQtyTime;
                                obj.SettingTime = SettingTime;
                                obj.SummationOfSCTvsPP = SummationOfSCTvsPP;
                                obj.Isdeleted = 0;
                                obj.lossDuration = lossdet / 60;
                                obj.opid = OperatorId;
                                obj.MachineId = macrow.MachineID;
                                obj.shift = Shift;
                                obj.OTTime = Convert.ToString(oTDet);

                                obj.opWorkingDuration = hrdet;
                                db.tblolecaldetails.Add(obj);
                                db.SaveChanges();
                            }
                        }
                    }

                }

            }

        }

        public async Task<double> GetScrapQtyTimeOfRWO(string UsedDateForExcel, int MachineID)
        {
            double SQT = 0;
            using (i_facility_tsalEntities1 dbhmi = new i_facility_tsalEntities1())
            {
                var PartsData = dbhmi.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0) && m.isWorkOrder == 1).ToList();
                foreach (var row in PartsData)
                {
                    string partno = row.PartNo;
                    string operationno = row.OperationNo;
                    int scrapQty = Convert.ToInt32(row.Rej_Qty);
                    int DeliveredQty = Convert.ToInt32(row.Delivered_Qty);
                    DateTime startTime = Convert.ToDateTime(row.Date);
                    DateTime endTime = Convert.ToDateTime(row.Time);
                    Double WODuration = await GetGreen(UsedDateForExcel, startTime, endTime, MachineID);

                    //Double WODuration = endTime.Subtract(startTime).TotalMinutes;
                    //For Availability Loss
                    //double Settingtime = GetSetupForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID);
                    //double green = GetOT(UsedDateForExcel, startTime, endTime, MachineID);
                    //double DownTime = GetDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "ROA");
                    //double BreakdownTime = GetBreakDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID);
                    //double AL = DownTime + BreakdownTime + Settingtime;

                    //For Performance Loss
                    //double downtimeROP = GetDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "ROP");
                    //double minorlossWO = GetMinorLossForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "yellow");
                    //double PL = downtimeROP + minorlossWO;

                    SQT += (WODuration / 60);
                }
            }
            return await Task.FromResult<double>(SQT);
        }

        public async Task<double> GetScrapQtyTimeOfWO(string UsedDateForExcel, int MachineID)
        {
            double SQT = 0;
            using (i_facility_tsalEntities1 dbhmi = new i_facility_tsalEntities1())
            {
                var PartsData = dbhmi.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0) && m.isWorkOrder == 0).ToList();
                foreach (var row in PartsData)
                {
                    string partno = row.PartNo;
                    string operationno = row.OperationNo;
                    int scrapQty = 0;
                    int DeliveredQty = 0;
                    string scrapQtyString = Convert.ToString(row.Rej_Qty);
                    string DeliveredQtyString = Convert.ToString(row.Delivered_Qty);
                    string x = scrapQtyString;
                    int value;
                    if (int.TryParse(x, out value))
                    {
                        scrapQty = value;
                    }
                    x = DeliveredQtyString;
                    if (int.TryParse(x, out value))
                    {
                        DeliveredQty = value;
                    }

                    DateTime startTime = Convert.ToDateTime(row.Date);
                    DateTime endTime = Convert.ToDateTime(row.Time);
                    //Double WODuration = endTimeTemp.Subtract(startTime).TotalMinutes;
                    Double WODuration = await GetGreen(UsedDateForExcel, startTime, endTime, MachineID);

                    if ((scrapQty + DeliveredQty) == 0)
                    {
                        SQT += 0;
                    }
                    else
                    {
                        SQT += ((WODuration / 60) / (scrapQty + DeliveredQty)) * scrapQty;
                    }
                }
            }
            return await Task.FromResult<double>(SQT);
        }

        public async Task<double> GetGreen(string UsedDateForExcel, DateTime StartTime, DateTime EndTime, int MachineID)
        {
            double settingTime = 0;

            DataTable lossesData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {

                mc.open();
                //String query1 = "SELECT Sum(DurationInSec) From tblmode WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and ColorCode = 'green'"
                //    + " and ( StartTime >= '" + WOstarttimeDate + "' and EndTime <= '" + WOendtimeDate + "' )";

                String query1 = "SELECT StartTime,EndTime,ModeID From " + MsqlConnection.DBSchemaName + ".tblmode WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and ColorCode = 'green'  and"
                   + "( StartTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
                   + " ( StartTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";
                

                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(lossesData);
                mc.close();
                //if (lossesData.Rows.Count > 0)
                //{
                //    //settingTime = Convert.ToDouble(lossesData.Rows[0][0]);
                //    settingTime = 0;
                //}

                for (int i = 0; i < lossesData.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][1])))
                    {
                        DateTime LStartDate = Convert.ToDateTime(lossesData.Rows[i][0]);
                        DateTime LEndDate = Convert.ToDateTime(lossesData.Rows[i][1]);
                        double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

                        //Get Duration Based on start & end Time.

                        if (LStartDate < StartTime)
                        {
                            double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
                            IndividualDur -= StartDurationExtra;
                        }
                        if (LEndDate > EndTime)
                        {
                            double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
                            IndividualDur -= EndDurationExtra;
                        }
                        settingTime += IndividualDur;
                    }
                }
            }
            return await Task.FromResult<double>(settingTime);
        }

        public async Task<double> GetSummationOfSCTvsPP(string UsedDateForExcel, int MachineID)
        {
            double SummationofTime = 0;
            // UsedDateForExcel = "2018-12-01";

            #region OLD 2017-02-10
            //var PartsData = db.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)).ToList();
            //if (PartsData.Count == 0)
            //{
            //    //return -1;
            //}
            //foreach (var row in PartsData)
            //{
            //    string partno = row.PartNo;
            //    string operationno = row.OperationNo;
            //    int totalpartproduced = Convert.ToInt32(row.Delivered_Qty) + Convert.ToInt32(row.Rej_Qty);
            //    Double stdCuttingTime = 0;
            //    var stdcuttingTimeData = db.tblmasterparts_st_sw.Where(m => m.IsDeleted == 0 && m.OpNo == operationno && m.PartNo == partno).FirstOrDefault();
            //    if (stdcuttingTimeData != null)
            //    {
            //        string stdcuttingvalString = Convert.ToString(stdcuttingTimeData.StdCuttingTime);
            //        Double stdcuttingval = 0;
            //        if (double.TryParse(stdcuttingvalString, out stdcuttingval))
            //        {
            //            stdcuttingval = stdcuttingval;
            //        }

            //        string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
            //        if (Unit == "Hrs")
            //        {
            //            stdCuttingTime = stdcuttingval * 60;
            //        }
            //        else //Unit is Minutes
            //        {
            //            stdCuttingTime = stdcuttingval;
            //        }
            //    }
            //    SummationofTime += stdCuttingTime * totalpartproduced;
            //}

            ////To Extract MultiWorkOrder Cutting Time
            //PartsData = db.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.IsMultiWO == 1 && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)).ToList();
            //if (PartsData.Count == 0)
            //{
            //    return SummationofTime;
            //}
            //foreach (var row in PartsData)
            //{
            //    int HMIID = row.HMIID;

            //    var DataInMultiwoSelection = db.tbl_multiwoselection.Where(m => m.HMIID == HMIID).ToList();
            //    foreach (var rowData in DataInMultiwoSelection)
            //    {
            //        string partno = rowData.PartNo;
            //        string operationno = rowData.OperationNo;
            //        int totalpartproduced = Convert.ToInt32(rowData.DeliveredQty) + Convert.ToInt32(rowData.ScrapQty);
            //        int stdCuttingTime = 0;
            //        var stdcuttingTimeData = db.tblmasterparts_st_sw.Where(m => m.IsDeleted == 0 && m.OpNo == operationno && m.PartNo == partno).FirstOrDefault();
            //        if (stdcuttingTimeData != null)
            //        {
            //            int stdcuttingval = Convert.ToInt32(stdcuttingTimeData.StdCuttingTime);
            //            string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
            //            if (Unit == "Hrs")
            //            {
            //                stdCuttingTime = stdcuttingval * 60;
            //            }
            //            else //Unit is Minutes
            //            {
            //                stdCuttingTime = stdcuttingval;
            //            }
            //        }
            //        SummationofTime += stdCuttingTime * totalpartproduced;
            //    }
            //}

            #endregion

            #region OLD 2017-02-10
            //List<string> OccuredWOs = new List<string>();
            ////To Extract Single WorkOrder Cutting Time
            //using (i_facility_tsalEntities1 dbhmi = new i_facility_tsalEntities1())
            //{
            //    var PartsDataAll = dbhmi.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.IsMultiWO == 0 && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)).OrderByDescending(m => m.HMIID).ToList();
            //    if (PartsDataAll.Count == 0)
            //    {
            //        //return SummationofTime;
            //    }
            //    foreach (var row in PartsDataAll)
            //    {
            //        string partNo = row.PartNo;
            //        string woNo = row.Work_Order_No;
            //        string opNo = row.OperationNo;

            //        string occuredwo = partNo + "," + woNo + "," + opNo;
            //        if (!OccuredWOs.Contains(occuredwo))
            //        {
            //            OccuredWOs.Add(occuredwo);
            //            var PartsData = dbhmi.tblhmiscreens.
            //                Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.IsMultiWO == 0
            //                    && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)
            //                    && m.Work_Order_No == woNo && m.PartNo == partNo && m.OperationNo == opNo).
            //                    OrderByDescending(m => m.HMIID).ToList();

            //            int totalpartproduced = 0;
            //            int ProcessQty = 0, DeliveredQty = 0;
            //            //Decide to select deliveredQty & ProcessedQty lastest(from HMI or tblmultiWOselection)

            //            #region new code

            //            //here 1st get latest of delivered and processed among row in tblHMIScreen & tblmulitwoselection
            //            int isHMIFirst = 2; //default NO History for that wo,pn,on

            //            var mulitwoData = dbhmi.tbl_multiwoselection.Where(m => m.WorkOrder == woNo && m.PartNo == partNo && m.OperationNo == opNo).OrderByDescending(m => m.MultiWOID).Take(1).ToList();
            //            //var hmiData = db.tblhmiscreens.Where(m => m.Work_Order_No == WONo && m.PartNo == Part && m.OperationNo == Operation && m.isWorkInProgress == 0).OrderByDescending(m => m.HMIID).Take(1).ToList();

            //            //Note: we are in this loop => hmiscreen table data is Available

            //            if (mulitwoData.Count > 0)
            //            {
            //                isHMIFirst = 1;
            //            }
            //            else if (PartsData.Count > 0)
            //            {
            //                isHMIFirst = 0;
            //            }
            //            else if (PartsData.Count > 0 && mulitwoData.Count > 0) //we both Dates now check for greatest amongst
            //            {
            //                int hmiIDFromMulitWO = row.HMIID;
            //                DateTime multiwoDateTime = Convert.ToDateTime(from r in db.tblhmiscreens
            //                                                              where r.HMIID == hmiIDFromMulitWO
            //                                                              select r.Time
            //                                                              );
            //                DateTime hmiDateTime = Convert.ToDateTime(row.Time);

            //                if (Convert.ToInt32(multiwoDateTime.Subtract(hmiDateTime).TotalSeconds) > 0)
            //                {
            //                    isHMIFirst = 1; // multiwoDateTime is greater than hmitable datetime
            //                }
            //                else
            //                {
            //                    isHMIFirst = 0;
            //                }
            //            }
            //            if (isHMIFirst == 1)
            //            {
            //                string delivString = Convert.ToString(mulitwoData[0].DeliveredQty);
            //                int.TryParse(delivString, out DeliveredQty);
            //                string processString = Convert.ToString(mulitwoData[0].ProcessQty);
            //                int.TryParse(processString, out ProcessQty);

            //            }
            //            else if (isHMIFirst == 0)//Take Data from HMI
            //            {
            //                string delivString = Convert.ToString(PartsData[0].Delivered_Qty);
            //                int.TryParse(delivString, out DeliveredQty);
            //                string processString = Convert.ToString(PartsData[0].ProcessQty);
            //                int.TryParse(processString, out ProcessQty);
            //            }

            //            #endregion

            //            //totalpartproduced = DeliveredQty + ProcessQty;
            //            totalpartproduced = DeliveredQty;

            //            #region InnerLogic Common for both ways(HMI or tblmultiWOselection)

            //            double stdCuttingTime = 0;
            //            var stdcuttingTimeData = db.tblmasterparts_st_sw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
            //            if (stdcuttingTimeData != null)
            //            {
            //                double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
            //                string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
            //                if (Unit == "Hrs")
            //                {
            //                    stdCuttingTime = stdcuttingval * 60;
            //                }
            //                else //Unit is Minutes
            //                {
            //                    stdCuttingTime = stdcuttingval;
            //                }
            //            }
            //            #endregion

            //            SummationofTime += stdCuttingTime * totalpartproduced;
            //        }
            //    }
            //}
            ////To Extract Multi WorkOrder Cutting Time
            //using (i_facility_tsalEntities1 dbhmi = new i_facility_tsalEntities1())
            //{
            //    var PartsDataAll = dbhmi.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.IsMultiWO == 1 && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)).ToList();
            //    if (PartsDataAll.Count == 0)
            //    {
            //        //return SummationofTime;
            //    }
            //    foreach (var row in PartsDataAll)
            //    {
            //        string partNo = row.PartNo;
            //        string woNo = row.Work_Order_No;
            //        string opNo = row.OperationNo;

            //        string occuredwo = partNo + "," + woNo + "," + opNo;
            //        if (!OccuredWOs.Contains(occuredwo))
            //        {
            //            OccuredWOs.Add(occuredwo);
            //            var PartsData = dbhmi.tblhmiscreens.
            //                Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.IsMultiWO == 0
            //                    && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)
            //                    && m.Work_Order_No == woNo && m.PartNo == partNo && m.OperationNo == opNo).
            //                    OrderByDescending(m => m.HMIID).ToList();

            //            int totalpartproduced = 0;
            //            int ProcessQty = 0, DeliveredQty = 0;
            //            //Decide to select deliveredQty & ProcessedQty lastest(from HMI or tblmultiWOselection)

            //            #region new code

            //            //here 1st get latest of delivered and processed among row in tblHMIScreen & tblmulitwoselection
            //            int isHMIFirst = 2; //default NO History for that wo,pn,on

            //            var mulitwoData = dbhmi.tbl_multiwoselection.Where(m => m.WorkOrder == woNo && m.PartNo == partNo && m.OperationNo == opNo).OrderByDescending(m => m.MultiWOID).Take(1).ToList();
            //            //var hmiData = db.tblhmiscreens.Where(m => m.Work_Order_No == WONo && m.PartNo == Part && m.OperationNo == Operation && m.isWorkInProgress == 0).OrderByDescending(m => m.HMIID).Take(1).ToList();

            //            //Note: we are in this loop => hmiscreen table data is Available

            //            if (mulitwoData.Count > 0)
            //            {
            //                isHMIFirst = 1;
            //            }
            //            else if (PartsData.Count > 0)
            //            {
            //                isHMIFirst = 0;
            //            }
            //            else if (PartsData.Count > 0 && mulitwoData.Count > 0) //we have both Dates now check for greatest amongst
            //            {
            //                int hmiIDFromMulitWO = row.HMIID;
            //                DateTime multiwoDateTime = Convert.ToDateTime(from r in db.tblhmiscreens
            //                                                              where r.HMIID == hmiIDFromMulitWO
            //                                                              select r.Time
            //                                                              );
            //                DateTime hmiDateTime = Convert.ToDateTime(row.Time);

            //                if (Convert.ToInt32(multiwoDateTime.Subtract(hmiDateTime).TotalSeconds) > 0)
            //                {
            //                    isHMIFirst = 1; // multiwoDateTime is greater than hmitable datetime
            //                }
            //                else
            //                {
            //                    isHMIFirst = 0;
            //                }
            //            }

            //            if (isHMIFirst == 1)
            //            {
            //                string delivString = Convert.ToString(mulitwoData[0].DeliveredQty);
            //                int.TryParse(delivString, out DeliveredQty);
            //                string processString = Convert.ToString(mulitwoData[0].ProcessQty);
            //                int.TryParse(processString, out ProcessQty);
            //            }
            //            else if (isHMIFirst == 0) //Take Data from HMI
            //            {
            //                string delivString = Convert.ToString(PartsData[0].Delivered_Qty);
            //                int.TryParse(delivString, out DeliveredQty);
            //                string processString = Convert.ToString(PartsData[0].ProcessQty);
            //                int.TryParse(processString, out ProcessQty);
            //            }

            //            #endregion

            //            //totalpartproduced = DeliveredQty + ProcessQty;
            //            totalpartproduced = DeliveredQty;
            //            #region InnerLogic Common for both ways(HMI or tblmultiWOselection)

            //            double stdCuttingTime = 0;
            //            var stdcuttingTimeData = db.tblmasterparts_st_sw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
            //            if (stdcuttingTimeData != null)
            //            {
            //                double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
            //                string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
            //                if (Unit == "Hrs")
            //                {
            //                    stdCuttingTime = stdcuttingval * 60;
            //                }
            //                else //Unit is Minutes
            //                {
            //                    stdCuttingTime = stdcuttingval;
            //                }
            //            }
            //            #endregion

            //            SummationofTime += stdCuttingTime * totalpartproduced;
            //        }
            //    }
            //}
            #endregion

            //new Code 2017-03-08
            using (i_facility_tsalEntities1 dbhmi = new i_facility_tsalEntities1())
            {
                var PartsDataAll = dbhmi.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)).OrderByDescending(m => m.PartNo).ThenByDescending(m => m.OperationNo).ToList();
                if (PartsDataAll.Count == 0)
                {
                    //return SummationofTime;
                }
                foreach (var row in PartsDataAll)
                {
                    if (row.IsMultiWO == 0)
                    {
                        string partNo = row.PartNo;
                        string woNo = row.Work_Order_No;
                        string opNo = row.OperationNo;
                        int DeliveredQty = 0;
                        DeliveredQty = Convert.ToInt32(row.Delivered_Qty);
                        #region InnerLogic Common for both ways(HMI or tblmultiWOselection)
                        double stdCuttingTime = 0;
                        var stdcuttingTimeData = db.tblmasterparts_st_sw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
                        if (stdcuttingTimeData != null)
                        {
                            double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
                            string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
                            if (Unit == "Hrs")
                            {
                                stdCuttingTime = stdcuttingval * 60;
                            }
                            else if (Unit == "Sec") //Unit is Minutes
                            {
                                stdCuttingTime = stdcuttingval / 60;
                            }
                            else
                            {
                                stdCuttingTime = stdcuttingval;
                            }
                            // no need of else , its already in minutes
                        }
                        #endregion
                        //MessageBox.Show("CuttingTime " + stdCuttingTime +" DeliveredQty " +DeliveredQty );
                        SummationofTime += stdCuttingTime * DeliveredQty;
                        //MessageBox.Show("Single" + SummationofTime);
                    }
                    else
                    {
                        int hmiid = row.HMIID;
                        var multiWOData = dbhmi.tbl_multiwoselection.Where(m => m.HMIID == hmiid).ToList();
                        foreach (var rowMulti in multiWOData)
                        {
                            string partNo = rowMulti.PartNo;
                            string opNo = rowMulti.OperationNo;
                            int DeliveredQty = 0;
                            DeliveredQty = Convert.ToInt32(rowMulti.DeliveredQty);
                            #region
                            double stdCuttingTime = 0;
                            var stdcuttingTimeData = db.tblmasterparts_st_sw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
                            if (stdcuttingTimeData != null)
                            {
                                double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
                                string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
                                if (Unit == "Hrs")
                                {
                                    stdCuttingTime = stdcuttingval * 60;
                                }
                                else if (Unit == "Sec") //Unit is Minutes
                                {
                                    stdCuttingTime = stdcuttingval / 60;
                                }
                                else
                                {
                                    stdCuttingTime = stdcuttingval;
                                }

                            }
                            #endregion
                            //MessageBox.Show("CuttingTime " + stdCuttingTime + " DeliveredQty " + DeliveredQty);
                            SummationofTime += stdCuttingTime * DeliveredQty;
                            //MessageBox.Show("Multi" + SummationofTime);
                        }
                    }
                    //MessageBox.Show("" + SummationofTime);
                }
            }
            return await Task.FromResult<double>(SummationofTime);
        }

        public async Task<double> GetDownTimeLosses(string UsedDateForExcel, int MachineID, string contribute)
        {
            double LossTime = 0;
            //string contribute = "ROA";
            //getting all ROA sublevels ids. Only those of IDLE.

            using (i_facility_tsalEntities1 dbLoss = new i_facility_tsalEntities1())
            {
                var SettingIDs = dbLoss.tbllossescodes.Where(m => m.ContributeTo == contribute && (m.MessageType != "PM" || m.MessageType != "BREAKDOWN")).Select(m => m.LossCodeID).ToList();

                var SettingData = dbLoss.tbllossofentries.Where(m => SettingIDs.Contains(m.MessageCodeID) && m.MachineID == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();

                var LossDuration = dbLoss.tblmodes.Where(m => m.MachineID == MachineID && m.CorrectedDate == UsedDateForExcel && m.IsCompleted == 1 && m.DurationInSec > 120 && m.ColorCode == "YELLOW").Sum(m => m.DurationInSec);

                //foreach (var row in SettingData)
                //{
                //    DateTime startTime = Convert.ToDateTime(row.StartDateTime);
                //    DateTime endTime = Convert.ToDateTime(row.EndDateTime);
                //    LossTime += endTime.Subtract(startTime).TotalMinutes;
                //}
                try
                {
                    LossTime = (int)LossDuration;
                }
                catch { }
            }
            return await Task.FromResult<double>(LossTime);
        }

        public async Task<double> GetSettingTime(string UsedDateForExcel, int MachineID)
        {
            double settingTime = 0;
            int setupid = 0;
            string settingString = "Setup";
            var setupiddata = db.tbllossescodes.Where(m => m.MessageType.Contains(settingString)).FirstOrDefault();
            if (setupiddata != null)
            {
                setupid = setupiddata.LossCodeID;
            }
            else
            {
                //Session["Error"] = "Unable to get Setup's ID";
                return -1;
            }
            //getting all setup's sublevels ids.
            using (i_facility_tsalEntities1 dbLoss = new i_facility_tsalEntities1())
            {
                var SettingIDs = dbLoss.tbllossescodes.Where(m => m.LossCodesLevel1ID == setupid || m.LossCodesLevel2ID == setupid).Select(m => m.LossCodeID).ToList();


                //settingTime = (from row in db.tbllivelossofenties
                //               where  row.CorrectedDate == UsedDateForExcel && row.MachineID == MachineID );


                var SettingData = dbLoss.tbllossofentries.Where(m => SettingIDs.Contains(m.MessageCodeID) && m.MachineID == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
                foreach (var row in SettingData)
                {
                    DateTime startTime = Convert.ToDateTime(row.StartDateTime);
                    DateTime endTime = Convert.ToDateTime(row.EndDateTime);
                    settingTime += endTime.Subtract(startTime).TotalMinutes;
                }
            }
            return await Task.FromResult<double>(settingTime);
        }

        public async Task<double> GetOPIDleBreakDown(string CorrectedDate, int MachineID, string Colour)
        {
            DateTime currentdate = Convert.ToDateTime(CorrectedDate);
            string datetime = currentdate.ToString("yyyy-MM-dd");

            double count = 0;
            //MsqlConnection mc = new MsqlConnection();
            //mc.open();
            ////operating
            //mc.open();
            //String query1 = "SELECT count(ID) From tbldailyprodstatus WHERE CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND ColorCode='" + Colour + "'";
            //SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
            //DataTable OP = new DataTable();
            //da1.Fill(OP);
            //mc.close();
            //if (OP.Rows.Count != 0)
            //{
            //    count[0] = Convert.ToInt32(OP.Rows[0][0]);
            //}

            using (i_facility_tsalEntities1 dbLoss = new i_facility_tsalEntities1())
            {
                var blah = dbLoss.tblmodes.Where(m => m.MachineID == MachineID && m.CorrectedDate == CorrectedDate && m.ColorCode == Colour).Sum(m => m.DurationInSec);
                count = await Task.FromResult<double>(Convert.ToDouble(blah));
            }
            return count;
        }



        public  DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            DataTable tbl = new DataTable();
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {

                    try
                    {
                        using (var stream = System.IO.File.OpenRead(path))
                        {
                            pck.Load(stream);
                        }
                        var ws = pck.Workbook.Worksheets.First();

                        foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                        {
                            tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                        }
                        var startRow = hasHeader ? 2 : 1;
                        for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                        {
                            //18 Columns in out excel fixed.
                            var wsRow = ws.Cells[rowNum, 1, rowNum, 18];
                            DataRow row = tbl.Rows.Add();
                            foreach (var cell in wsRow)
                            {
                                row[cell.Start.Column - 1] = cell.Text;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        using (Form1 obj = new Form1())
                        {
                            obj.IntoFile("Reading excel Data " + e);
                        }
                    }

                }
            }
            catch (Exception exc)
            {

                IntoFile("GetDataTableFromExcel section: " + exc);
            }
            return tbl;
        }

        public void IntoFile(string Msg)
        {
            string path1 = AppDomain.CurrentDomain.BaseDirectory;
            string appPath = Application.StartupPath + @"\OLELogFile.txt";
            using (StreamWriter writer = new StreamWriter(appPath, true)) //true => Append Text
            {
                writer.WriteLine(System.DateTime.Now + ":  " + Msg);
            }
        }

        private string GetEndCorrectedDate(string EndTime,DateTime StartDate)
        {
            string CorrectedDate = "";

            
            TimeSpan EndTime1 = TimeSpan.Parse(EndTime);
            TimeSpan  Start = TimeSpan.Parse("06:00:00");
            if (Start < EndTime1)
            {
                CorrectedDate = StartDate.ToString("yyyy-MM-dd");
            }
            else
            {
                CorrectedDate = StartDate.AddDays(1).ToString("yyyy-MM-dd");
            }
            return CorrectedDate;
        }
    }
}
