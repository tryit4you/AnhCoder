using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MMS.AnalyzeFile.Analyze
{
    public class LP
    {
        public string MeterSerial { get; set; }
        public string P_C { get; set; }
        public string P_T { get; set; }
        public string Q_C { get; set; }
        public string Q_T { get; set; }

        public DateTime TimeData { get; set; }

        public LP(string v0, string v1, string v2, string v3, string v4, DateTime v5)
        {
            this.MeterSerial = v0;
            this.P_C = v1;
            this.P_T = v2;
            this.Q_C = v3;
            this.Q_T = v4;
            this.TimeData = v5;
        }

    }

    public class DataTUTI
    {
        public string MeterSerial;
        public double CT_TU;
        public double CT_MAU;
        public double VT_TU;
        public double VT_MAU;

        public DataTUTI()
        {

        }

        public DataTUTI(string v0, double v1, double v2, double v3, double v4)
        {
            this.MeterSerial = v0;
            this.CT_TU = v1;
            this.CT_MAU = v2;
            this.VT_TU = v3;
            this.VT_MAU = v4;

        }

    }


    public class DataPQmax
    {
        public double pmax;
        public double pmax1;
        public double pmax2;
        public double pmax3;

        public DateTime timepmax;
        public DateTime timepmax1;
        public DateTime timepmax2;
        public DateTime timepmax3;
        public string Loaidulieu;
        public string MA_DIEMDO;
        public int dateid;
        public int monthid;
        public int yearid;



        public string strSerial;

        public DataPQmax()
        {

        }

        public DataPQmax(string inputLoaidulieu, double inputpmax, double inputpmax1, double inputpmax2, double inputpmax3, DateTime inputp, DateTime inputp1, DateTime inputp2, DateTime inputp3, string serial, string Ma_diemdo)
        {
            pmax = inputpmax;
            pmax1 = inputpmax1;
            pmax2 = inputpmax2;
            pmax3 = inputpmax3;
            timepmax = inputp;
            timepmax1 = inputp1;
            timepmax2 = inputp2;
            timepmax3 = inputp3;
            strSerial = serial;
            Loaidulieu = inputLoaidulieu;
            MA_DIEMDO = Ma_diemdo;
        }
        public DataPQmax(DataPQmax obj)
        {
            pmax = obj.pmax;
            pmax1 = obj.pmax1;
            pmax2 = obj.pmax2;
            pmax3 = obj.pmax3;
            timepmax = obj.timepmax;
            timepmax1 = obj.timepmax1;
            timepmax2 = obj.timepmax2;
            timepmax3 = obj.timepmax3;
            strSerial = obj.strSerial;
            Loaidulieu = obj.Loaidulieu;
            MA_DIEMDO = obj.MA_DIEMDO;
            dateid = obj.dateid;
            monthid = obj.monthid;
            yearid = obj.yearid;
        }
    }

    public class DataChiso
    {
        public double chiso;
        public double chiso1;
        public double chiso2;
        public double chiso3;
        public int dateid;
        public int monthid;
        public int yearid;
        public double Qtong;
        public string strSerial;
        public string Loaidulieu;
        public DateTime Thoigian;
        public string MA_DIEMDO;
        public int strSeq;
        public DataChiso()
        {

        }
        public DataChiso(DataChiso obj)
        {
            chiso = obj.chiso;
            chiso1 = obj.chiso1;
            chiso2 = obj.chiso2;
            chiso3 = obj.chiso3;
            strSerial = obj.strSerial;
            Qtong = obj.Qtong;
            Loaidulieu = obj.Loaidulieu;
            Thoigian = obj.Thoigian;
            MA_DIEMDO = obj.MA_DIEMDO;
            strSeq = obj.strSeq;
            dateid = obj.dateid;
            monthid = obj.monthid;
            yearid = obj.yearid;
        }

        public DataChiso(double inputchiso, double inputchiso1, double inputchiso2, double inputchiso3, double Q, string serial, string loaidulieu, DateTime dtThoigian, string strMA_DIEMDO, int strSQ)
        {
            chiso = inputchiso;
            chiso1 = inputchiso1;
            chiso2 = inputchiso2;
            chiso3 = inputchiso3;
            strSerial = serial;
            Qtong = Q;
            Loaidulieu = loaidulieu;
            Thoigian = dtThoigian;
            MA_DIEMDO = strMA_DIEMDO;
            strSeq = strSQ;
        }
    }

    public struct DataRepeater
    {
        public double chiso;
        public string strSerial;
        public string Loaidulieu;
        public DateTime Thoigian;
        public int dateid;
        public int monthid;
        public int yearid;
        public DataRepeater(double inputchiso, double inputchiso1, double inputchiso2, double inputchiso3, double Q, string serial, string loaidulieu, DateTime dtThoigian, int idateid, int imonthid, int iyearid)
        {
            chiso = inputchiso;
            strSerial = serial;
            Loaidulieu = loaidulieu;
            Thoigian = dtThoigian;
            dateid = idateid;
            monthid = imonthid;
            yearid = iyearid;
        }
    }
    public struct DataEvent
    {
        public string maevent;
        public string strSerial;
        public string Loaidulieu;
        public DateTime Thoigian;

        public DataEvent(string imaevent, string serial, string loaidulieu, DateTime dtThoigian)
        {
            maevent = imaevent;
            strSerial = serial;
            Loaidulieu = loaidulieu;
            Thoigian = dtThoigian;
        }
    }

    public struct DataLoadprofile
    {
        public double P;
        public string strSerial;
        public string Loaidulieu;
        public DateTime Thoigian;

        public DataLoadprofile(double inP, string serial, string loaidulieu, DateTime dtThoigian)
        {
            P = inP;
            strSerial = serial;
            Loaidulieu = loaidulieu;
            Thoigian = dtThoigian;
        }
    }

    public class DataBLock
    {
        public string MA_DIEMDO = null;
        public int? SeqID = null;
        public string SERIALID = null;
        public DateTime? NGAYGIO = null;
        public DateTime? START_DATE = null;
        public DateTime? END_DATE = null;
        public DateTime? START_BILL = null;
        public DateTime? END_BILL = null;
        public int? MONTHID = null;
        public int? YEARID = null;
        public int? DATEID = null;
        public int? MINUTEID = null;
        public int? HOURID = null;

        public double? TU_TU = null;
        public double? TU_MAU = null;
        public double? TI_TU = null;
        public double? TI_MAU = null;
        public double? TU = null;
        public double? TI = null;

        public double? HSN = null;
        public double? P_V_A = null;
        public double? P_V_B = null;
        public double? P_V_C = null;
        public double? P_A_A = null;
        public double? P_A_B = null;
        public double? P_A_C = null;
        public double? P_AP_T = null;
        public double? P_AP_A = null;
        public double? P_AP_B = null;
        public double? P_AP_C = null;
        public double? P_RP_T = null;
        public double? P_RP_A = null;
        public double? P_RP_B = null;
        public double? P_RP_C = null;
        public double? COSFI = null;
        public double? P_IMPORTKWH = null; // Chỉ sô biểu tổng 
        public double? P_IMPBT = null; // Chí số biểu 1 
        public double? P_IMPCD = null; // CHỉ số biểu 2
        public double? P_IMPTD = null; // Chỉ số biểu 3 
        public double? P_EXPORTKWH = null;
        public double? P_EXPBT = null;
        public double? P_EXPTD = null;
        public double? P_EXPCD = null;
        public double? P_C1 = null;
        public double? P_C2 = null;
        public double? PMAX_IMPORTKWH = null;
        public DateTime? PMAX_DATE = null;
        public double? PMAX_IMPORTKWHNHAN = null;
        public DateTime? PMAX_DATENHAN = null;
        public double? LOAD_P_IMPORTKW = null;
        public double? LOAD_P_EXPORTKW = null;
        public double? LOAD_P_C1 = null;
        public double? LOAD_P_C2 = null;
        public double? LOAD_P_CONG_CS = null;
        public double? LOAD_P_TRU_CS = null;
        public double? LOAD_Q_TRU_CS = null;
        public double? LOAD_Q_CONG_CS = null;
        public DateTime? P_NGAY = null;
        public double? CHOT_HCNCS = null;
        public double? CHOT_HCNCSB1 = null;
        public double? CHOT_HCNCSB2 = null;
        public double? CHOT_HCNCSB3 = null;
        public double? CHOT_HCGCS = null;
        public double? CHOT_HCGCSB1 = null;
        public double? CHOT_HCGCSB2 = null;
        public double? CHOT_HCGCSB3 = null;
        public double? CHOT_VCGCS = null;
        public double? CHOT_VCGCSB1 = null;
        public double? CHOT_VCGCSB2 = null;
        public double? CHOT_VCGCSB3 = null;
        public double? CHOT_VCNCS = null;
        public double? CHOT_VCNCSB1 = null;
        public double? CHOT_VCNCSB2 = null;
        public double? CHOT_VCNCSB3 = null;
        public DateTime? P_CLOCK_METER = null;


    }


    public class AnalyzeDCU
    {
        public bool isruning = false;
        public RichTextBox rtbLog;
        public object Analyze { get; private set; }

        public void do_analyze(string path, RichTextBox rtbLog)
        {
            try
            {
                string err = "";
                dynamic obj = new ExpandoObject();
                string error = "";
                string[] files = System.IO.Directory.GetFiles(path, "*.res");
                int result = -1;
                foreach (string filename in files)
                {
                    string values = System.IO.File.ReadAllText(filename);
                    char[] delimiters = new char[] { '\r', '\n' };
                    string[] value_item = values.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < value_item.Length; i++)
                    {
                        string[] row_item = value_item[i].Split(new char[] { '\t' });
                        if (row_item.Length != 8)
                            continue;
                        try
                        {

                            obj.P_MA_DIEMDO = row_item[0].ToString();
                            obj.P_SERIALID = row_item[1].ToString();
                            obj.P_READING_PATH = row_item[6].ToString();
                            obj.P_DATETIME = Convert.ToDateTime(row_item[7].ToString());

                            string[] item_para = row_item[2].Split(new char[] { '_' });
                            if (item_para.Count() > 3)
                            {
                                obj.P_IMPORTKWH = double.Parse(item_para[0].ToString().Replace(",", "."));
                                obj.P_EXPORTKWH = double.Parse(item_para[1].ToString().Replace(",", "."));
                                obj.P_VOLPHASE = double.Parse(item_para[3].ToString().Replace(",", "."));
                                obj.P_CURPHASE = double.Parse(item_para[2].ToString().Replace(",", "."));
                                obj.P_ACTIVE_POWER_PHASE = 0;
                            }
                            else
                            {
                                obj.P_IMPORTKWH = double.Parse(row_item[2].ToString().Replace(",", "."));
                                obj.P_EXPORTKWH = 0;
                                obj.P_VOLPHASE = 0;
                                obj.P_CURPHASE = 0;
                                obj.P_ACTIVE_POWER_PHASE = 0;
                            }

                            //result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_CHISO_1PHA, ref obj);
                            result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_CHISO_1PHA, ref obj, ref err);
                            if (err != "")
                            {
                                MessageBox.Show(err);
                            }

                            if (result != -1)
                            {
                                System.IO.File.Delete(filename);
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }
            }
            catch (Exception ex)
            { }
        }


        public void do_analyze_huuhong_1pha_block(string path, RichTextBox rtbLog)
        {
            try
            {

                string sourcePath = @"D:\EVNHES_DCU";
                System.IO.Directory.CreateDirectory(@"D:\EVNHES_DCU\RES_ERR\" + DateTime.Now.ToString("ddMMyy"));
                string targetPath = @"D:\EVNHES_DCU\RES_ERR\" + DateTime.Now.ToString("ddMMyy");
                // string oradb = OracleConnStringnew("10.9.0.214", "1521", "EVNHES", "EVNHES", "EVNHES");
                string oradb = General.connectString;
                //  OracleConnection conn = new OracleConnection(oradb);  // C#

                string LoaiDulieu = "";
                string[] value_item = new string[10];
                string[] value_filename = new string[10];
                dynamic obj = new ExpandoObject();
                dynamic objHIS = new ExpandoObject();
                dynamic objCUR = new ExpandoObject();
                dynamic objLoad = new ExpandoObject();
                string error = "";
                string[] files = System.IO.Directory.GetFiles(path, "*.res");

                int filescount = files.Count();
                int result = -1;
                string sourceFile;
                string destFile;
                string ketqua = "";
                string values = "";
                string tenfile = "";
                int icount = 0;


                foreach (string filename in files)
                {

                    sourceFile = System.IO.Path.Combine(sourcePath, filename);
                    destFile = System.IO.Path.Combine(targetPath);
                    try
                    {
                        rtbLog.Clear();
                    }
                    catch
                    {
                    }

                    try
                    {
                        values = System.IO.File.ReadAllText(filename);
                        if (values == "")
                        {
                            System.IO.File.Delete(filename);
                            return;
                        }

                        char[] delimiters = new char[] { '\r', '\n' };
                        string strNgaygio = "";
                        string strTemp = "";
                        value_item[icount] = values.Trim();
                        tenfile = Path.GetFileName(filename);

                        General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                        File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, LoaiDulieu.ToString() + tenfile), true);
                        LoaiDulieu = "";
                        System.IO.File.Delete(filename);

                        if (result != -1)
                        {
                            General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                            System.IO.File.Delete(filename);
                        }
                        icount = icount + 1;
                        if (icount == 10)
                            break;
                    }
                    catch
                    {
                        return;
                    }

                }


                List<DataPQmax> PQMax = new List<DataPQmax>();
                List<DataChiso> ListChiso = new List<DataChiso>();
                List<DataChiso> ListChisoGio = new List<DataChiso>();

                List<DataLoadprofile> ListLoadprofile = new List<DataLoadprofile>();
                List<DataEvent> ListEvent = new List<DataEvent>();
                List<DataRepeater> ListRepeater = new List<DataRepeater>();
                // List insert
                List<DataChiso> ListChisoQ = new List<DataChiso>();
                List<DataChiso> ListChisoHC = new List<DataChiso>();

                if (filescount != 0)
                {
                    analyzeHUUHONG_NEW(value_item, ref ListRepeater, ref PQMax, ref ListChiso, ref ListChisoGio, ref ListLoadprofile, ref ListEvent, ref obj);
                    DataChiso objChiso = new DataChiso();
                    DataLoadprofileCk objLoadDCU = new DataLoadprofileCk();

                    #region Luu bulkinsert 
                    /*
                 string LstSerial = "";
                 for (int iicount = ListChiso.Count - 1; iicount >= 0; iicount--)
                 {

                     LstSerial = LstSerial + ListChiso[iicount].strSerial + ",";

                 }

                 OracleConnection conn2 = new OracleConnection(oradb);  // C#


                 DataTable _listmadiemdo = GetListMeter_1phareSerial(LstSerial, conn);
                 int sequenid = 0;
                 for (int iicount = ListChiso.Count - 1; iicount >= 0; iicount--)
                 {

                     var res = from row in _listmadiemdo.AsEnumerable()
                               where row.Field<string>("SERIALID") == ListChiso[iicount].strSerial
                               select row.Field<string>("MA_DIEMDO");
                     //int.TryParse((System.DateTime.Now.Day.ToString() + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Year.ToString()+ System.DateTime.Now.Millisecond.ToString()), out sequenid);
                     int.TryParse(( System.DateTime.Now.Year.ToString() + System.DateTime.Now.Millisecond.ToString()), out sequenid);

                     if (res.Count() > 0)
                     {
                         ListChiso[iicount].MA_DIEMDO = res.ToList()[0].ToString();
                         ListChiso[iicount].strSeq = sequenid+ iicount;



                     }
                     else
                     {
                         ListChiso[iicount].MA_DIEMDO ="PD00CHUAMA";
                         ListChiso[iicount].strSeq = sequenid;

                     }

                     if (ListChiso[iicount].Loaidulieu == "041B" || ListChiso[iicount].Loaidulieu == "0114")
                     {
                         DataChiso datacs = new DataChiso(ListChiso[iicount]);

                         ListChisoHC.Add(datacs);

                     }
                     else
                     {
                         if (ListChiso[iicount].Loaidulieu == "081B")
                         {
                             DataChiso datacs = new DataChiso(ListChiso[iicount]);
                           //  datacs. = ListChisoQ[iicount].();
                             ListChisoQ.Add(datacs);
                         }
                     }



                 }



                 insert_ImisChisoDCU_bulk(conn2, ListChisoHC);

                  LstSerial = "";
                 for (int iicount = PQMax.Count - 1; iicount >= 0; iicount--)
                 {

                     LstSerial = LstSerial + PQMax[iicount].strSerial + ",";

                 }

               _listmadiemdo = GetListMeter_1phareSerial(LstSerial, conn);

                 for (int iicount = PQMax.Count - 1; iicount >= 0; iicount--)
                 {

                     var res = from row in _listmadiemdo.AsEnumerable()
                               where row.Field<string>("SERIALID") == PQMax[iicount].strSerial
                               select row.Field<string>("MA_DIEMDO");
                     //int.TryParse((System.DateTime.Now.Day.ToString() + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Year.ToString()+ System.DateTime.Now.Millisecond.ToString()), out sequenid);
                    // int.TryParse((System.DateTime.Now.Year.ToString() + System.DateTime.Now.Millisecond.ToString()), out sequenid);

                     if (res.Count() > 0)
                     {
                         PQMax[iicount].MA_DIEMDO = res.ToList()[0].ToString();
                        // PQMax[iicount].strSeq = sequenid + iicount;



                     }
                     else
                     {
                         PQMax[iicount].MA_DIEMDO = "PD00CHUAMA";
                        // ListChiso[iicount].strSeq = sequenid;

                     }

                 }

                 insert_ImisPMax_bulk(conn2, PQMax);

                 */
                    #endregion





                    #region inssert loadprofile 
                    OracleConnection conn = new OracleConnection(oradb);  // C#
                    if (conn != null && conn.State == ConnectionState.Closed)
                        conn.Open();
                    if (ListLoadprofile.Count() > 0)
                    {


                        //  var item = ListLoadprofile.FirstOrDefault(o => o.strSerial == "000017082450");
                        List<DataLoadprofile> Listdataloadprofile;

                        var query = (from t in ListLoadprofile
                                     group t by new { t.strSerial, t.Thoigian }
                                 into grp
                                     select new
                                     {
                                         grp.Key.strSerial,
                                         grp.Key.Thoigian,
                                     }).ToList();


                        if (conn != null && conn.State == ConnectionState.Closed)
                            conn.Open();

                        foreach (var iitem in query)
                        {
                            Listdataloadprofile = (from records in ListLoadprofile
                                                   where records.strSerial == iitem.strSerial
                                                   && records.Thoigian == iitem.Thoigian
                                                   select records).ToList();
                            foreach (var tempobj in Listdataloadprofile)
                            {
                                switch (tempobj.Loaidulieu)
                                {

                                    case "0127":
                                        objLoadDCU.P_IMPORT = 0;
                                        objLoadDCU.P_IMPORT = Convert.ToDouble(tempobj.P);
                                        break;
                                    case "0227":
                                        objLoadDCU.P_EXPORT = 0;
                                        objLoadDCU.P_EXPORT = Convert.ToDouble(tempobj.P);
                                        break;
                                    case "0427":
                                        objLoadDCU.Q_IMPORT = 0;
                                        objLoadDCU.Q_IMPORT = Convert.ToDouble(tempobj.P);
                                        break;
                                    case "0827":
                                        objLoadDCU.Q_EXPORT = 0;
                                        objLoadDCU.Q_EXPORT = Convert.ToDouble(tempobj.P);
                                        break;
                                    default:

                                        break;

                                }
                                //objLoadDCU.P_IMPORTKW_CS = tempobj.P
                                //     objLoadDCU.P_EXPORTKW_CS
                            }
                            objLoadDCU.strSerial = iitem.strSerial;
                            objLoadDCU.Thoigian = iitem.Thoigian;



                            try
                            {
                                //  connection.Open();

                                OracleCommand cmd = new OracleCommand("PKG_DCU_MDMS.INSERT_LOADPRO_HUUHONG", conn);
                                cmd.CommandType = CommandType.StoredProcedure;

                                cmd.Parameters.Add("@P_SERIALID", OracleDbType.NVarchar2).Value = objLoadDCU.strSerial;
                                cmd.Parameters.Add("@P_STARTDATE", OracleDbType.Date).Value = objLoadDCU.Thoigian;
                                cmd.Parameters.Add("@P_IMPORTKW", OracleDbType.Double).Value = objLoadDCU.P_IMPORT;
                                cmd.Parameters.Add("@P_EXPORTKW", OracleDbType.Double).Value = objLoadDCU.P_EXPORT;
                                cmd.Parameters.Add("@P_C1", OracleDbType.Double).Value = objLoadDCU.Q_IMPORT;
                                cmd.Parameters.Add("@P_C2", OracleDbType.Double).Value = objLoadDCU.Q_EXPORT;
                                cmd.ExecuteNonQuery();

                                objLoadDCU.P_IMPORT = 0;
                                objLoadDCU.P_EXPORT = 0;
                                objLoadDCU.Q_IMPORT = 0;
                                objLoadDCU.Q_EXPORT = 0;

                            }
                            catch (Exception ex)
                            {
                                objLoadDCU.P_IMPORT = 0;
                                objLoadDCU.P_EXPORT = 0;
                                objLoadDCU.Q_IMPORT = 0;
                                objLoadDCU.Q_EXPORT = 0;
                                //    MessageBox.Show(ex.Message);
                            }


                            //      result = OracleDataAccess1.execute_Loadprofile_byObject(ref error, General.connectionstr, PKG_Library.MDMSDCU_INSERT_LOADPRO_DCU,  objLoadDCU, ref error);
                            //  result = 1;
                        }





                        conn.Close();
                        // conn.Dispose();

                    }
                    #endregion

                    #region Insertchiso

                    if (ListChiso.Count > 0)
                    {
                        var query = (from t in ListChiso
                                     group t by new { t.strSerial, t.Thoigian }
                               into grp
                                     select new
                                     {
                                         grp.Key.strSerial,
                                         grp.Key.Thoigian,
                                     }).ToList();




                        List<DataChiso> Listchiso;
                        // INSERT_CHISO_HUUHONG
                        if (conn != null && conn.State == ConnectionState.Closed)
                            conn.Open();
                        OracleTransaction txn = conn.BeginTransaction(IsolationLevel.ReadCommitted);



                        foreach (var iitem in query)
                        {
                            Listchiso = (from records in ListChiso
                                         where records.strSerial == iitem.strSerial
                                         select records).ToList();
                            //  && records.Thoigian == iitem.Thoigian
                            foreach (var tempobj in Listchiso)
                            {
                                switch (tempobj.Loaidulieu)
                                {

                                    case "0114":
                                        objChiso.chiso = Convert.ToDouble(tempobj.chiso);
                                        objChiso.chiso1 = Convert.ToDouble(tempobj.chiso1);
                                        objChiso.chiso2 = Convert.ToDouble(tempobj.chiso2);
                                        objChiso.chiso3 = Convert.ToDouble(tempobj.chiso3);
                                        objChiso.Loaidulieu = tempobj.Loaidulieu;
                                        break;
                                    case "0214":
                                        objChiso.Qtong = Convert.ToDouble(tempobj.chiso);
                                        break;
                                    case "041B":
                                        objChiso.chiso = Convert.ToDouble(tempobj.chiso);
                                        objChiso.chiso1 = Convert.ToDouble(tempobj.chiso1);
                                        objChiso.chiso2 = Convert.ToDouble(tempobj.chiso2);
                                        objChiso.chiso3 = Convert.ToDouble(tempobj.chiso3);
                                        objChiso.Loaidulieu = tempobj.Loaidulieu;
                                        break;
                                    case "081B":
                                        objChiso.Qtong = Convert.ToDouble(tempobj.Qtong);
                                        objChiso.Loaidulieu = tempobj.Loaidulieu;
                                        LoaiDulieu = "QTong";
                                        break;
                                    default:

                                        break;

                                }

                            }
                            objChiso.strSerial = iitem.strSerial;
                            LoaiDulieu = LoaiDulieu + "_" + iitem.strSerial;
                            objChiso.Thoigian = iitem.Thoigian;

                            try
                            {
                                OracleCommand cmd = new OracleCommand("PKG_DCU_MDMS.INSERT_CHISO_HUUHONG", conn);
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.Add("@P_SERIALID", OracleDbType.NVarchar2).Value = objChiso.strSerial;
                                cmd.Parameters.Add("@P_DATE", OracleDbType.Date).Value = objChiso.Thoigian;
                                cmd.Parameters.Add("@P_CHISO", OracleDbType.Double).Value = objChiso.chiso;
                                cmd.Parameters.Add("@P_CHISOA", OracleDbType.Double).Value = objChiso.chiso1;
                                cmd.Parameters.Add("@P_CHISOB", OracleDbType.Double).Value = objChiso.chiso2;
                                cmd.Parameters.Add("@P_CHISOC", OracleDbType.Double).Value = objChiso.chiso3;
                                cmd.Parameters.Add("@P_C1", OracleDbType.Double).Value = objChiso.Qtong;
                                cmd.Parameters.Add("@P_C2", OracleDbType.Double).Value = 0;
                                cmd.Parameters.Add("@P_LOAICS", OracleDbType.NVarchar2).Value = objChiso.Loaidulieu;
                                cmd.ExecuteNonQuery();
                                objChiso.chiso = 0;
                                objChiso.chiso1 = 0;
                                objChiso.chiso2 = 0;
                                objChiso.chiso3 = 0;
                                objChiso.Qtong = 0;

                            }

                            catch (Exception ex)
                            {
                                objChiso.chiso = 0;
                                objChiso.chiso1 = 0;
                                objChiso.chiso2 = 0;
                                objChiso.chiso3 = 0;
                                objChiso.Qtong = 0;
                            }

                        }

                        //   Insert_List_ChiSo_1Phase(ListChiso);
                        txn.Commit();

                        #region test


                        //try
                        //{
                        //    var cmd2 = conn.CreateCommand();

                        //    cmd2.CommandText = "pkgInsertChiso.ProcInsertChisoArray";
                        //    cmd2.CommandType = CommandType.StoredProcedure;

                        //    OracleParameter param1 = new OracleParameter();
                        //    param1.Direction = ParameterDirection.Input;
                        //    param1.CollectionType = OracleCollectionType.PLSQLAssociativeArray;
                        //    param1.Value = ListChiso.ToArray();
                        //    param1.Size = ListChiso.Count;
                        //    param1.DbType = DbType.Object;
                        //    cmd2.ExecuteNonQuery();
                        //    conn.Close();

                        //}
                        //catch (Exception ex)
                        //{
                        //   MessageBox.Show(ex.Message);
                        //}

                        #endregion


                        conn.Close();



                    }

                    #endregion

                    #region InsertPQMax
                    // INSERT_PQ MAX
                    if (PQMax.Count > 0)
                    {

                        if (conn != null && conn.State == ConnectionState.Closed)
                            conn.Open();
                        foreach (var tempobj in PQMax)
                        {
                            try
                            {
                                OracleCommand cmd = new OracleCommand("PKG_DCU_MDMS.INSERT_PQMAX_HUUHONG", conn);
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.Add("@P_SERIALID", OracleDbType.NVarchar2).Value = tempobj.strSerial;
                                cmd.Parameters.Add("@P_DATE", OracleDbType.Date).Value = tempobj.timepmax;
                                cmd.Parameters.Add("@P_PMAX", OracleDbType.Double).Value = tempobj.pmax;
                                cmd.Parameters.Add("@P_LOAIDL", OracleDbType.Double).Value = tempobj.Loaidulieu;
                                cmd.ExecuteNonQuery();

                            }
                            catch (Exception ex)
                            {
                                //  MessageBox.Show(ex.Message);
                            }

                        }
                        conn.Close();

                    }


                    #endregion

                    #region InsertEvent
                    if (ListEvent.Count > 0)
                    {

                        if (conn != null && conn.State == ConnectionState.Closed)
                            conn.Open();
                        foreach (var tempobj in ListEvent)
                        {
                            try
                            {
                                OracleCommand cmd = new OracleCommand("PKG_DCU_MDMS.INSERT_EVENT_HUUHONG", conn);
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.Add("@p_SERIALID", OracleDbType.NVarchar2).Value = tempobj.strSerial;
                                cmd.Parameters.Add("@p_EVENT", OracleDbType.NVarchar2).Value = tempobj.maevent;
                                cmd.Parameters.Add("@p_DATE", OracleDbType.Date).Value = tempobj.Thoigian;
                                cmd.ExecuteNonQuery();

                            }
                            catch (Exception ex)
                            {
                                //  MessageBox.Show(ex.Message);
                            }

                        }
                        conn.Close();

                    }
                    #endregion




                }

            }
            catch (Exception ex)
            {
                General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex.Message), Color.Red);

            }
        }


        private bool insert_ImisChisoDCU_bulk(OracleConnection conn, List<DataChiso> objCUR)
        {
            bool returnValue = false;
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                string query = @"insert into EVNHES.DCU_CHISO_TEST(MA_DIEMDO,SERIALID,IMPORTKWH,NGAYGIO,DATEID,MONTH,YEAR) 
                                                         values (:MA_DIEMDO,:SERIALID,:IMPORTKWH,:NGAYGIO,:DATEID,:MONTHID,:YEARID)";


                using (var command = conn.CreateCommand())
                {
                    command.CommandText = query;
                    command.CommandType = CommandType.Text;
                    command.BindByName = true;
                    // In order to use ArrayBinding, the ArrayBindCount property
                    // of OracleCommand object must be set to the number of records to be inserted
                    command.ArrayBindCount = objCUR.Count();
                    //  command.Parameters.Add(":ID", OracleDbType.Varchar2, objCUR.Select(c => c.SeqID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MA_DIEMDO", OracleDbType.Varchar2, objCUR.Select(c => c.MA_DIEMDO).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":SERIALID", OracleDbType.Varchar2, objCUR.Select(c => c.strSerial).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":IMPORTKWH", OracleDbType.Double, objCUR.Select(c => c.chiso).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":NGAYGIO", OracleDbType.Date, objCUR.Select(c => c.Thoigian).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MONTHID", OracleDbType.Int32, objCUR.Select(c => c.monthid).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":YEARID", OracleDbType.Int32, objCUR.Select(c => c.yearid).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":DATEID", OracleDbType.Int32, objCUR.Select(c => c.dateid).ToArray(), ParameterDirection.Input);
                    int result = command.ExecuteNonQuery();
                    if (result == objCUR.Count)
                        returnValue = true;
                }

            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                // conn.Close();
                return false;
            }

            //  conn.Close();
            return true;

        }


        private bool insert_ImisPMax_bulk(OracleConnection conn, List<DataPQmax> objCUR)
        {
            bool returnValue = false;
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();
                //  string query = @" INSERT INTO IMIS_MAXDEMAND_TEST(MA_DIEMDO,SERIALID,IMPORTKW,IMPORTKW_TIME,PDATE,PMONTH,PYEAR) VALUES(:P_MA_DIEMDO,:P_SERIALID,:P_IMPORTKW;:P_IMPORTKW_TIME,:DATEID,:MONTHID,:YEARID)";

                string query = @" INSERT INTO IMIS_MAXDEMAND_TEST(MA_DIEMDO,SERIALID,IMPORTKW,IMPORTKW_TIME,PDATE,PMONTH,PYEAR) VALUES(:P_MA_DIEMDO,:P_SERIALID,:P_IMPORTKW,:P_IMPORTKW_TIME,:DATEID,:MONTHID,:YEARID)";

                using (var command = conn.CreateCommand())
                {
                    command.CommandText = query;
                    command.CommandType = CommandType.Text;
                    command.BindByName = true;
                    // In order to use ArrayBinding, the ArrayBindCount property
                    // of OracleCommand object must be set to the number of records to be inserted
                    command.ArrayBindCount = objCUR.Count();
                    //  command.Parameters.Add(":ID", OracleDbType.Varchar2, objCUR.Select(c => c.SeqID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":P_MA_DIEMDO", OracleDbType.Varchar2, objCUR.Select(c => c.MA_DIEMDO).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":P_SERIALID", OracleDbType.Varchar2, objCUR.Select(c => c.strSerial).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":P_IMPORTKW", OracleDbType.Double, objCUR.Select(c => c.pmax).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":P_IMPORTKW_TIME", OracleDbType.Date, objCUR.Select(c => c.timepmax).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":DATEID", OracleDbType.Int32, objCUR.Select(c => c.dateid).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MONTHID", OracleDbType.Int32, objCUR.Select(c => c.monthid).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":YEARID", OracleDbType.Int32, objCUR.Select(c => c.yearid).ToArray(), ParameterDirection.Input);

                    int result = command.ExecuteNonQuery();
                    if (result == objCUR.Count)
                        returnValue = true;
                }

            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                // conn.Close();
                return false;
            }

            //  conn.Close();
            return true;

        }
        public static bool Insert_List_ChiSo_1Phase(List<DataChiso> chiso, OracleConnection conn)
        {
            bool returnValue = false;
            string oradb = OracleConnStringnew("10.9.0.214", "1521", "EVNHES", "EVNHES", "EVNHES");
            try
            {
                //  OracleTransaction txn = conn.BeginTransaction(IsolationLevel.ReadCommitted);
                //   Oracle.ManagedDataAccess.Client.OracleConnection conn = new Oracle.ManagedDataAccess.Client.OracleConnection(oradb);

                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();
                OracleTransaction txn = conn.BeginTransaction();

                Oracle.ManagedDataAccess.Client.OracleCommand cmd = conn.CreateCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "PKG_DCU_MDMS.INSERT_CHISO_HUUHONG_TEST";
                cmd.ArrayBindCount = chiso.Count();
                cmd.BindByName = true;
                cmd.Parameters.Add("P_SERIALID", OracleDbType.Varchar2, chiso.Select(x => x.strSerial).ToArray(), ParameterDirection.Input);
                cmd.Parameters.Add("P_DATE", OracleDbType.Date, chiso.Select(x => x.Thoigian).ToArray(), ParameterDirection.Input);
                cmd.Parameters.Add("P_CHISO", OracleDbType.Varchar2, chiso.Select(x => x.chiso).ToArray(), ParameterDirection.Input);
                cmd.Parameters.Add("P_CHISOA", OracleDbType.Decimal, chiso.Select(x => x.chiso1).ToArray(), ParameterDirection.Input);
                cmd.Parameters.Add("P_CHISOB", OracleDbType.Decimal, chiso.Select(x => x.chiso2).ToArray(), ParameterDirection.Input);
                cmd.Parameters.Add("P_CHISOC", OracleDbType.Decimal, chiso.Select(x => x.chiso3).ToArray(), ParameterDirection.Input);
                cmd.Parameters.Add("P_C1", OracleDbType.Decimal, chiso.Select(x => x.Qtong).ToArray(), ParameterDirection.Input);
                cmd.Parameters.Add("P_C2", OracleDbType.Decimal, chiso.Select(x => 0).ToArray(), ParameterDirection.Input);
                cmd.Parameters.Add("P_LOAICS", OracleDbType.NVarchar2, chiso.Select(x => x.Loaidulieu).ToArray(), ParameterDirection.Input);


                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                txn.Commit();
                conn.Close();
                returnValue = true;

            }
            catch (Exception ex)
            {
                returnValue = false;

            }
            return returnValue;
        }

        public void do_analyze_huuhong_1pha(string path, RichTextBox rtbLog)
        {
            try
            {
                // HCM D:\EVNHES_DCU

                // string sourcePath = @"C:\Users\Administrator\Desktop\DCU_2152018";
                string sourcePath = @"D:\EVNHES_DCU";
                System.IO.Directory.CreateDirectory(@"D:\EVNHES_DCU\RES_ERR\" + DateTime.Now.ToString("ddMMyy"));

                string targetPath = @"D:\EVNHES_DCU\RES_ERR\" + DateTime.Now.ToString("ddMMyy");
                //  string oradb = OracleConnStringnew("10.9.0.214", "1521", "EVNHES", "EVNHES", "EVNHES");
                string oradb = General.connectString;
                OracleConnection conn = new OracleConnection(oradb);  // C#

                string LoaiDulieu = "";

                dynamic obj = new ExpandoObject();
                dynamic objHIS = new ExpandoObject();
                dynamic objCUR = new ExpandoObject();
                dynamic objLoad = new ExpandoObject();




                string error = "";

                string[] files = System.IO.Directory.GetFiles(path, "*.res");

                int filescount = files.Count();

                int result = -1;
                foreach (string filename in files)
                {



                    string sourceFile = System.IO.Path.Combine(sourcePath, filename);
                    string destFile = System.IO.Path.Combine(targetPath);
                    string ketqua = "";
                    string values = "";
                    try
                    {
                        values = System.IO.File.ReadAllText(filename);
                        if (values == "")
                        {
                            System.IO.File.Delete(filename);
                            return;
                        }
                    }
                    catch
                    {
                        return;
                    }
                    char[] delimiters = new char[] { '\r', '\n' };
                    string strNgaygio = "";
                    string strTemp = "";
                    string[] value_item = values.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    string tenfile = "";
                    List<DataPQmax> PQMax = new List<DataPQmax>();
                    List<DataChiso> ListChiso = new List<DataChiso>();
                    List<DataLoadprofile> ListLoadprofile = new List<DataLoadprofile>();
                    List<DataEvent> ListEvent = new List<DataEvent>();
                    List<DataRepeater> ListRepet = new List<DataRepeater>();

                    if (1 == 1)
                    {
                        analyzeHUUHONG_NEW(value_item, ref ListRepet, ref PQMax, ref ListChiso, ref ListLoadprofile, ref ListEvent, ref obj);
                        DataChiso objChiso = new DataChiso();
                        DataLoadprofileCk objLoadDCU = new DataLoadprofileCk();


                        #region InsertRepeater
                        if (ListRepet.Count > 0)
                        {
                            var query = (from t in ListRepet
                                         group t by new { t.strSerial, t.Thoigian }
                                   into grp
                                         select new
                                         {
                                             grp.Key.strSerial,
                                             grp.Key.Thoigian,
                                         }).ToList();

                            List<DataRepeater> Listrepeate;
                            // INSERT_CHISO_HUUHONG
                            if (conn != null && conn.State == ConnectionState.Closed)
                                conn.Open();
                            OracleTransaction txn = conn.BeginTransaction(IsolationLevel.ReadCommitted);



                            foreach (var iitem in query)
                            {
                                Listrepeate = (from records in ListRepet
                                               where records.strSerial == iitem.strSerial
                                             && records.Thoigian == iitem.Thoigian
                                               select records).ToList();

                                foreach (var tempobj in Listrepeate)
                                {
                                    LoaiDulieu = "";
                                    switch (tempobj.Loaidulieu)
                                    {

                                        case "201C":
                                            objChiso.chiso = Convert.ToDouble(tempobj.chiso);
                                            objChiso.Loaidulieu = tempobj.Loaidulieu;
                                            break;
                                        default:

                                            break;

                                    }

                                }
                                objChiso.strSerial = iitem.strSerial;
                                LoaiDulieu = LoaiDulieu + "_" + iitem.strSerial;
                                objChiso.Thoigian = iitem.Thoigian;

                                try
                                {
                                    OracleCommand cmd = new OracleCommand("PKG_DCU_MDMS.INSERT_REPEATER_HUUHONG", conn);
                                    cmd.CommandType = CommandType.StoredProcedure;
                                    cmd.Parameters.Add("@P_SERIALID", OracleDbType.NVarchar2).Value = objChiso.strSerial;
                                    cmd.Parameters.Add("@P_DATE", OracleDbType.Date).Value = objChiso.Thoigian;
                                    cmd.Parameters.Add("@P_CHISO", OracleDbType.Double).Value = objChiso.chiso;
                                    cmd.Parameters.Add("@P_LOAICS", OracleDbType.NVarchar2).Value = objChiso.Loaidulieu;
                                    cmd.ExecuteNonQuery();
                                }

                                catch (Exception ex)
                                {
                                    // MessageBox.Show(ex.Message);
                                }

                            }

                            //   Insert_List_ChiSo_1Phase(ListChiso);
                            txn.Commit();

                            conn.Close();



                        }

                        #endregion


                        #region inssert loadprofile 

                        if (ListLoadprofile.Count() > 0)
                        {


                            //  var item = ListLoadprofile.FirstOrDefault(o => o.strSerial == "000017082450");
                            List<DataLoadprofile> Listdataloadprofile;

                            var query = (from t in ListLoadprofile
                                         group t by new { t.strSerial, t.Thoigian }
                                     into grp
                                         select new
                                         {
                                             grp.Key.strSerial,
                                             grp.Key.Thoigian,
                                         }).ToList();


                            if (conn != null && conn.State == ConnectionState.Closed)
                                conn.Open();

                            foreach (var iitem in query)
                            {
                                Listdataloadprofile = (from records in ListLoadprofile
                                                       where records.strSerial == iitem.strSerial
                                                       && records.Thoigian == iitem.Thoigian
                                                       select records).ToList();
                                foreach (var tempobj in Listdataloadprofile)
                                {
                                    switch (tempobj.Loaidulieu)
                                    {

                                        case "0127":
                                            objLoadDCU.P_IMPORT = Convert.ToDouble(tempobj.P);
                                            break;
                                        case "0227":
                                            objLoadDCU.P_EXPORT = Convert.ToDouble(tempobj.P);
                                            break;
                                        case "0427":
                                            objLoadDCU.Q_IMPORT = Convert.ToDouble(tempobj.P);
                                            break;
                                        case "0827":
                                            objLoadDCU.Q_EXPORT = Convert.ToDouble(tempobj.P);
                                            break;
                                        default:

                                            break;

                                    }
                                    //objLoadDCU.P_IMPORTKW_CS = tempobj.P
                                    //     objLoadDCU.P_EXPORTKW_CS
                                }
                                objLoadDCU.strSerial = iitem.strSerial;
                                objLoadDCU.Thoigian = iitem.Thoigian;



                                try
                                {
                                    //  connection.Open();

                                    OracleCommand cmd = new OracleCommand("PKG_DCU_MDMS.INSERT_LOADPRO_HUUHONG", conn);
                                    cmd.CommandType = CommandType.StoredProcedure;

                                    cmd.Parameters.Add("@P_SERIALID", OracleDbType.NVarchar2).Value = objLoadDCU.strSerial;
                                    cmd.Parameters.Add("@P_STARTDATE", OracleDbType.Date).Value = objLoadDCU.Thoigian;
                                    cmd.Parameters.Add("@P_IMPORTKW", OracleDbType.Double).Value = objLoadDCU.P_IMPORT;
                                    cmd.Parameters.Add("@P_EXPORTKW", OracleDbType.Double).Value = objLoadDCU.P_EXPORT;
                                    cmd.Parameters.Add("@P_C1", OracleDbType.Double).Value = objLoadDCU.Q_IMPORT;
                                    cmd.Parameters.Add("@P_C2", OracleDbType.Double).Value = objLoadDCU.Q_EXPORT;
                                    cmd.ExecuteNonQuery();


                                }
                                catch (Exception ex)
                                {
                                    //    MessageBox.Show(ex.Message);
                                }


                                //      result = OracleDataAccess1.execute_Loadprofile_byObject(ref error, General.connectionstr, PKG_Library.MDMSDCU_INSERT_LOADPRO_DCU,  objLoadDCU, ref error);
                                //  result = 1;
                            }





                            conn.Close();
                            // conn.Dispose();

                        }
                        #endregion

                        #region Insertchiso
                        if (ListChiso.Count > 0)
                        {
                            var query = (from t in ListChiso
                                         group t by new { t.strSerial, t.Thoigian }
                                   into grp
                                         select new
                                         {
                                             grp.Key.strSerial,
                                             grp.Key.Thoigian,
                                         }).ToList();

                            List<DataChiso> Listchiso;
                            // INSERT_CHISO_HUUHONG
                            if (conn != null && conn.State == ConnectionState.Closed)
                                conn.Open();
                            OracleTransaction txn = conn.BeginTransaction(IsolationLevel.ReadCommitted);



                            foreach (var iitem in query)
                            {
                                Listchiso = (from records in ListChiso
                                             where records.strSerial == iitem.strSerial
                                             && records.Thoigian == iitem.Thoigian
                                             select records).ToList();

                                foreach (var tempobj in Listchiso)
                                {
                                    LoaiDulieu = "";
                                    switch (tempobj.Loaidulieu)
                                    {

                                        case "0114":
                                            objChiso.chiso = Convert.ToDouble(tempobj.chiso);
                                            objChiso.chiso1 = Convert.ToDouble(tempobj.chiso1);
                                            objChiso.chiso2 = Convert.ToDouble(tempobj.chiso2);
                                            objChiso.chiso3 = Convert.ToDouble(tempobj.chiso3);
                                            objChiso.Loaidulieu = tempobj.Loaidulieu;
                                            break;
                                        //case "0214":
                                        //    objChiso.Qtong= Convert.ToDouble(tempobj.chiso);
                                        //    break;
                                        case "041B":
                                            objChiso.chiso = Convert.ToDouble(tempobj.chiso);
                                            objChiso.chiso1 = Convert.ToDouble(tempobj.chiso1);
                                            objChiso.chiso2 = Convert.ToDouble(tempobj.chiso2);
                                            objChiso.chiso3 = Convert.ToDouble(tempobj.chiso3);
                                            objChiso.Loaidulieu = tempobj.Loaidulieu;
                                            break;
                                        case "081B":
                                            objChiso.Qtong = Convert.ToDouble(tempobj.Qtong);
                                            objChiso.Loaidulieu = tempobj.Loaidulieu;
                                            LoaiDulieu = "QTong";
                                            break;
                                        default:

                                            break;

                                    }

                                }
                                objChiso.strSerial = iitem.strSerial;
                                LoaiDulieu = LoaiDulieu + "_" + iitem.strSerial;
                                objChiso.Thoigian = iitem.Thoigian;

                                try
                                {
                                    OracleCommand cmd = new OracleCommand("PKG_DCU_MDMS.INSERT_CHISO_HUUHONG", conn);
                                    cmd.CommandType = CommandType.StoredProcedure;
                                    cmd.Parameters.Add("@P_SERIALID", OracleDbType.NVarchar2).Value = objChiso.strSerial;
                                    cmd.Parameters.Add("@P_DATE", OracleDbType.Date).Value = objChiso.Thoigian;
                                    cmd.Parameters.Add("@P_CHISO", OracleDbType.Double).Value = objChiso.chiso;
                                    cmd.Parameters.Add("@P_CHISOA", OracleDbType.Double).Value = objChiso.chiso1;
                                    cmd.Parameters.Add("@P_CHISOB", OracleDbType.Double).Value = objChiso.chiso2;
                                    cmd.Parameters.Add("@P_CHISOC", OracleDbType.Double).Value = objChiso.chiso3;
                                    cmd.Parameters.Add("@P_C1", OracleDbType.Double).Value = objChiso.Qtong;
                                    cmd.Parameters.Add("@P_C2", OracleDbType.Double).Value = 0;
                                    cmd.Parameters.Add("@P_LOAICS", OracleDbType.NVarchar2).Value = objChiso.Loaidulieu;
                                    cmd.ExecuteNonQuery();


                                }

                                catch (Exception ex)
                                {
                                    // MessageBox.Show(ex.Message);
                                }

                            }

                            //   Insert_List_ChiSo_1Phase(ListChiso);
                            txn.Commit();

                            #region test


                            //try
                            //{
                            //    var cmd2 = conn.CreateCommand();

                            //    cmd2.CommandText = "pkgInsertChiso.ProcInsertChisoArray";
                            //    cmd2.CommandType = CommandType.StoredProcedure;

                            //    OracleParameter param1 = new OracleParameter();
                            //    param1.Direction = ParameterDirection.Input;
                            //    param1.CollectionType = OracleCollectionType.PLSQLAssociativeArray;
                            //    param1.Value = ListChiso.ToArray();
                            //    param1.Size = ListChiso.Count;
                            //    param1.DbType = DbType.Object;
                            //    cmd2.ExecuteNonQuery();
                            //    conn.Close();

                            //}
                            //catch (Exception ex)
                            //{
                            //   MessageBox.Show(ex.Message);
                            //}

                            #endregion


                            conn.Close();



                        }

                        #endregion

                        #region InsertPQMax
                        // INSERT_PQ MAX
                        if (PQMax.Count > 0)
                        {

                            if (conn != null && conn.State == ConnectionState.Closed)
                                conn.Open();
                            foreach (var tempobj in PQMax)
                            {
                                try
                                {

                                    if (conn != null && conn.State == ConnectionState.Closed)
                                        conn.Open();

                                    OracleCommand cmd = new OracleCommand("INSERT_IMIS_MAXDEMAND_HHMRF", conn);
                                    cmd.CommandType = CommandType.StoredProcedure;
                                    cmd.Parameters.Add("@P_IMPORTKW", OracleDbType.Double).Value = tempobj.pmax;
                                    cmd.Parameters.Add("@P_IMPORTKW_TIME", OracleDbType.Date).Value = tempobj.timepmax;
                                    cmd.Parameters.Add("@p_SERIALID", OracleDbType.NVarchar2).Value = tempobj.strSerial;
                                    cmd.ExecuteNonQuery();

                                }
                                catch (Exception ex)
                                {
                                    conn.Close();
                                    //  MessageBox.Show(ex.Message);
                                }

                            }
                            conn.Close();

                        }


                        #endregion

                        #region InsertEvent
                        if (ListEvent.Count > 0)
                        {

                            if (conn != null && conn.State == ConnectionState.Closed)
                                conn.Open();
                            foreach (var tempobj in ListEvent)
                            {
                                try
                                {
                                    OracleCommand cmd = new OracleCommand("PKG_DCU_MDMS.INSERT_EVENT_HUUHONG", conn);
                                    cmd.CommandType = CommandType.StoredProcedure;
                                    cmd.Parameters.Add("@p_SERIALID", OracleDbType.NVarchar2).Value = tempobj.strSerial;
                                    cmd.Parameters.Add("@p_EVENT", OracleDbType.NVarchar2).Value = tempobj.maevent;
                                    cmd.Parameters.Add("@p_DATE", OracleDbType.Date).Value = tempobj.Thoigian;
                                    cmd.ExecuteNonQuery();

                                }
                                catch (Exception ex)
                                {
                                    //  MessageBox.Show(ex.Message);
                                }

                            }
                            conn.Close();

                        }
                        #endregion

                        //     result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.INSERT_CHISO_1PHA_HUUHONG, ref obj, ref error);
                        //  tenfile = filename + System.DateTime.Now.Millisecond.ToString() + "_" + System.DateTime.Now.Date.ToString("ddMMyy") + "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString();
                        tenfile = Path.GetFileName(filename);

                        General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                        File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, LoaiDulieu.ToString() + tenfile), true);
                        LoaiDulieu = "";
                        System.IO.File.Delete(filename);

                        if (result != -1)
                        {
                            General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                            System.IO.File.Delete(filename);
                        }

                    }
                    else
                    {
                        for (int i = 2; i < value_item.Length; i++)
                        {
                            // obj.P_MA_DIEMDO = value_item[0].Split(new char[] { '\t' })[1];
                            //obj.P_SERIALID = value_item[2].Split(new char[] { '\t' })[1];
                            strNgaygio = value_item[1].Split(new char[] { '\t' })[1];
                            strTemp = strNgaygio.Substring(10, 2);
                            if (Convert.ToInt32(strTemp) < 30)
                            {
                                strTemp = "00";
                            }
                            else
                            {
                                strTemp = "30";
                            }

                            obj.P_NGAYGIO = DateTime.ParseExact(strNgaygio.Substring(0, 10) + strTemp + "00", "ddMMyyyyHHmmss", null);
                            //  obj.P_NGAYGIO = DateTime.Now;
                            obj.P_NGAY = DateTime.Now;

                            switch (value_item[0].Split(new char[] { '\t' })[1])
                            {
                                case "DCU_IFC":
                                    analyzeIFC(value_item, ref obj);
                                    break;
                                case "DCU_EMEC":
                                    analyzeEMEC(value_item, ref obj);
                                    result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_ROUTING_PATH, ref obj, ref error);
                                    break;
                                case "DTRFEXT":

                                    break;
                                case "HUUHONG_3G":

                                    obj.P_SERIALID = value_item[2].Split(new char[] { '\t' })[1];



                                    objCUR.P_MA_DIEMDO = obj.P_MA_DIEMDO;
                                    objHIS.P_MA_DIEMDO = obj.P_MA_DIEMDO;
                                    objLoad.P_MA_DIEMDO = obj.P_MA_DIEMDO;
                                    objCUR.P_NGAYGIO = obj.P_NGAYGIO;
                                    objHIS.P_NGAYGIO = obj.P_NGAYGIO;
                                    objLoad.P_NGAYGIO = obj.P_NGAYGIO;
                                    objCUR.P_NGAY = obj.P_NGAY;
                                    objHIS.P_NGAY = obj.P_NGAY;
                                    objLoad.P_NGAY = obj.P_NGAY;
                                    objCUR.P_SERIALID = obj.P_SERIALID;
                                    objHIS.P_SERIALID = obj.P_SERIALID;
                                    objLoad.P_SERIALID = obj.P_SERIALID;



                                    //  p_NGAYGIO              DATE,
                                    //p_CLOCK_METER DATE,
                                    //p_NGAY                 DATE,



                                    // To copy a folder's contents to a new location:
                                    // Create a new target folder, if necessary.
                                    if (!System.IO.Directory.Exists(targetPath))
                                    {
                                        System.IO.Directory.CreateDirectory(targetPath);
                                    }



                                    analyze3GHUUHONG_LoadProfile(value_item, ref objLoad, ref obj, ref objHIS, ref objCUR, ref ketqua);

                                    if (ketqua == "")
                                        result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_LOADPRO, ref objLoad, ref error);
                                    else
                                    {
                                        //System.IO.File.Copy(filename, destFile);
                                        tenfile = obj.P_MA_DIEMDO + "_" + System.DateTime.Now.Date.ToString("ddMMyy") + "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString();
                                        // tenfile = filename.Substring(sourceDir.Length + 1);

                                        File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, tenfile + ".res"), true);
                                        System.IO.File.Delete(filename);
                                    }
                                    if (result < 0)
                                    {
                                        //System.IO.File.Copy(filename, destFile);
                                        tenfile = obj.P_MA_DIEMDO + "_" + System.DateTime.Now.Date.ToString("ddMMyy") + "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString();
                                        // tenfile = filename.Substring(sourceDir.Length + 1);

                                        File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, tenfile + ".res"), true);
                                        System.IO.File.Delete(filename);
                                    }

                                    ketqua = "";
                                    error = "";
                                    analyze3GHUUHONG_TSVH(value_item, ref objLoad, ref obj, ref objHIS, ref objCUR, ref ketqua);
                                    if (ketqua == "")
                                        result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_CUR, ref objCUR, ref error);


                                    analyze3GHUUHONG_CSC(value_item, ref objLoad, ref obj, ref objHIS, ref objCUR, ref ketqua);
                                    objHIS.P_STARTBILL = new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, 1);
                                    objHIS.P_ENDBILL = new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, 1).AddMonths(1).AddDays(-1);
                                    if (ketqua == "")
                                        result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_HIS, ref objHIS, ref error);



                                    if (result != -1)
                                    {
                                        General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                                        System.IO.File.Delete(filename);
                                    }


                                    return;

                                default:
                                    break;
                            }
                            string err = "";
                            result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_CHISO_1PHA, ref obj, ref err);
                            if (err != "")
                            {
                                MessageBox.Show(err);
                            }
                            if (result != -1)
                            {
                                General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                                if (ketqua != "")
                                {
                                    System.IO.File.Copy(filename, destFile, true);
                                }

                                System.IO.File.Delete(filename);
                            }
                        }
                    }

                }
                conn.Dispose();
            }
            catch (Exception ex)
            {
                General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex.Message), Color.Red);

            }
        }



        /*
         *     cmd.Parameters.Add("@P_SERIALID", OracleDbType.NVarchar2).Value = objChiso.strSerial;
                                    cmd.Parameters.Add("@P_DATE", OracleDbType.Date).Value = objChiso.Thoigian;
                                    cmd.Parameters.Add("@P_CHISO", OracleDbType.Double).Value = objChiso.chiso;
                                    cmd.Parameters.Add("@P_CHISOA", OracleDbType.Double).Value = objChiso.chiso1;
                                    cmd.Parameters.Add("@P_CHISOB", OracleDbType.Double).Value = objChiso.chiso2;
                                    cmd.Parameters.Add("@P_CHISOC", OracleDbType.Double).Value = objChiso.chiso3;
                                    cmd.Parameters.Add("@P_C1", OracleDbType.Double).Value = objChiso.Qtong;
                                    cmd.Parameters.Add("@P_C2", OracleDbType.Double).Value = 0;
                                    cmd.Parameters.Add("@P_LOAICS", OracleDbType.NVarchar2).Value = objChiso.Loaidulieu;

         */


        public void do_analyze_move(string path, RichTextBox rtbLog)
        {
            try
            {
                Random rnd = new Random();

                int result = -1;
                int i = 0;

                // i = rnd.Next(1, 20);

                String directoryName = path;
                DirectoryInfo dirInfo = new DirectoryInfo(directoryName);

                if (dirInfo.Exists == false)
                    Directory.CreateDirectory(directoryName);

                List<string> MyMusicFiles = Directory
                                   .GetFiles(path, "*.res", SearchOption.AllDirectories).ToList();

                foreach (string file in MyMusicFiles)
                {
                    FileInfo mFile = new FileInfo(file);
                    directoryName = Path.GetDirectoryName(directoryName) + "\\Result";
                    dirInfo = new DirectoryInfo(directoryName + i.ToString());
                    // to remove name collisions
                    if (new FileInfo(dirInfo + "\\" + mFile.Name).Exists == false)
                    {


                        mFile.MoveTo(dirInfo + "\\" + mFile.Name);
                        i++;
                        if (i == General.channelCount)
                            i = 0;
                        //if (i == 9)
                        //{
                        //     i = rnd.Next(10, 15);
                        //}
                    }

                }




            }
            catch (Exception ex)
            {
                //  General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex.Message), Color.Red);

            }
        }



        public void do_analyze1(string path, RichTextBox rtbLog)
        {
            try
            {

                string sourcePath = @"C:\Users\MyPC\Desktop\DCU_2152018";
                string targetPath = @"C:\DU LIEU CONG TO DANG LAM\DCU\DCU Huu Hong\DCU_Anlyze\RES_ERR";


                dynamic obj = new ExpandoObject();
                dynamic objHIS = new ExpandoObject();
                dynamic objCUR = new ExpandoObject();
                dynamic objLoad = new ExpandoObject();


                string error = "";
                string[] files = System.IO.Directory.GetFiles(path, "*.res");
                int result = -1;
                foreach (string filename in files)
                {

                    string sourceFile = System.IO.Path.Combine(sourcePath, filename);
                    string destFile = System.IO.Path.Combine(targetPath);
                    string ketqua = "";
                    string values = System.IO.File.ReadAllText(filename);
                    char[] delimiters = new char[] { '\r', '\n' };
                    string strNgaygio = "";
                    string strTemp = "";
                    string[] value_item = values.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    string tenfile = "";
                    if (value_item[0].Contains("Vinasino"))
                    {
                        analyzeVinasino(value_item, filename);
                    }
                    else
                    {
                        for (int i = 2; i < value_item.Length; i++)
                        {
                            // obj.P_MA_DIEMDO = value_item[0].Split(new char[] { '\t' })[1];
                            //obj.P_SERIALID = value_item[2].Split(new char[] { '\t' })[1];
                            strNgaygio = value_item[1].Split(new char[] { '\t' })[1];
                            strTemp = strNgaygio.Substring(10, 2);
                            if (Convert.ToInt32(strTemp) < 30)
                            {
                                strTemp = "00";
                            }
                            else
                            {
                                strTemp = "30";
                            }

                            obj.P_NGAYGIO = DateTime.ParseExact(strNgaygio.Substring(0, 10) + strTemp + "00", "ddMMyyyyHHmmss", null);
                            //  obj.P_NGAYGIO = DateTime.Now;
                            obj.P_NGAY = DateTime.Now;

                            switch (value_item[0].Split(new char[] { '\t' })[1])
                            {
                                case "DCU_IFC":
                                    analyzeIFC(value_item, ref obj);
                                    break;
                                case "DCU_EMEC":
                                    analyzeEMEC(value_item, ref obj);
                                    result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_ROUTING_PATH, ref obj, ref error);
                                    break;
                                case "DTRFEXT":
                                    List<DataPQmax> DataPQmax = new List<DataPQmax>();
                                    List<DataChiso> DataChiso = new List<DataChiso>();
                                    List<DataLoadprofile> ListLoadprofile = new List<DataLoadprofile>();
                                    List<DataEvent> ListEvent = new List<DataEvent>();
                                    List<DataRepeater> ListRepeater = new List<DataRepeater>();
                                    analyzeHUUHONG_NEW(value_item, ref ListRepeater, ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListEvent, ref obj);
                                    result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.INSERT_CHISO_1PHA_HUUHONG, ref obj, ref error);
                                    tenfile = System.DateTime.Now.Millisecond.ToString() + "_" + System.DateTime.Now.Date.ToString("ddMMyy") + "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString();
                                    File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, tenfile + ".res"), true);
                                    break;
                                case "HUUHONG_3G":

                                    obj.P_SERIALID = value_item[2].Split(new char[] { '\t' })[1];
                                    objCUR.P_MA_DIEMDO = obj.P_MA_DIEMDO;
                                    objHIS.P_MA_DIEMDO = obj.P_MA_DIEMDO;
                                    objLoad.P_MA_DIEMDO = obj.P_MA_DIEMDO;
                                    objCUR.P_NGAYGIO = obj.P_NGAYGIO;
                                    objHIS.P_NGAYGIO = obj.P_NGAYGIO;
                                    objLoad.P_NGAYGIO = obj.P_NGAYGIO;
                                    objCUR.P_NGAY = obj.P_NGAY;
                                    objHIS.P_NGAY = obj.P_NGAY;
                                    objLoad.P_NGAY = obj.P_NGAY;
                                    objCUR.P_SERIALID = obj.P_SERIALID;
                                    objHIS.P_SERIALID = obj.P_SERIALID;
                                    objLoad.P_SERIALID = obj.P_SERIALID;

                                    //  p_NGAYGIO              DATE,
                                    //p_CLOCK_METER DATE,
                                    //p_NGAY                 DATE,



                                    // To copy a folder's contents to a new location:
                                    // Create a new target folder, if necessary.
                                    if (!System.IO.Directory.Exists(targetPath))
                                    {
                                        System.IO.Directory.CreateDirectory(targetPath);
                                    }



                                    analyze3GHUUHONG_LoadProfile(value_item, ref objLoad, ref obj, ref objHIS, ref objCUR, ref ketqua);

                                    if (ketqua == "")
                                        result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_LOADPRO, ref objLoad, ref error);
                                    else
                                    {
                                        //System.IO.File.Copy(filename, destFile);
                                        tenfile = obj.P_MA_DIEMDO + "_" + System.DateTime.Now.Date.ToString("ddMMyy") + "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString();
                                        // tenfile = filename.Substring(sourceDir.Length + 1);

                                        File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, tenfile + ".res"), true);
                                        System.IO.File.Delete(filename);
                                    }
                                    if (result < 0)
                                    {
                                        //System.IO.File.Copy(filename, destFile);
                                        tenfile = obj.P_MA_DIEMDO + "_" + System.DateTime.Now.Date.ToString("ddMMyy") + "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString();
                                        // tenfile = filename.Substring(sourceDir.Length + 1);

                                        File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, tenfile + ".res"), true);
                                        System.IO.File.Delete(filename);
                                    }

                                    ketqua = "";
                                    error = "";
                                    analyze3GHUUHONG_TSVH(value_item, ref objLoad, ref obj, ref objHIS, ref objCUR, ref ketqua);
                                    if (ketqua == "")
                                        result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_CUR, ref objCUR, ref error);


                                    analyze3GHUUHONG_CSC(value_item, ref objLoad, ref obj, ref objHIS, ref objCUR, ref ketqua);
                                    objHIS.P_STARTBILL = new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, 1);
                                    objHIS.P_ENDBILL = new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, 1).AddMonths(1).AddDays(-1);
                                    if (ketqua == "")
                                        result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_HIS, ref objHIS, ref error);



                                    if (result != -1)
                                    {
                                        General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                                        System.IO.File.Delete(filename);
                                    }


                                    return;

                                default:
                                    break;
                            }
                            string err = "";
                            result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_CHISO_1PHA, ref obj, ref err);
                            if (err != "")
                            {
                                // MessageBox.Show(err);
                            }
                            if (result != -1)
                            {
                                General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                                if (ketqua != "")
                                {
                                    System.IO.File.Copy(filename, destFile, true);
                                }

                                System.IO.File.Delete(filename);
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex.Message), Color.Red);

            }
        }


        public static byte[] StringToByteArray(string hex)
        {
            return Enumerable.Range(0, hex.Length)
                             .Where(x => x % 2 == 0)
                             .Select(x => Convert.ToByte(hex.Substring(x, 2), 16))
                             .ToArray();
        }

        static bool FileInUse(string path)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    return fs.CanWrite;
                }

            }
            catch (IOException ex)
            {
                return true;
            }
        }

        public static double ConvertArrayToInt2(byte[] b, int opt = 0)
        {
            int len = 2;
            byte[] bb = new byte[len];
            object o = null;
            switch (opt)
            {
                case 0: //2 byte
                    len = 2;
                    bb = new byte[len];
                    Array.Copy(b, 0, bb, 0, b.Length);
                    o = BitConverter.ToUInt16(bb, 0);
                    break;
                case 1: // 4 byte
                    len = 4;
                    bb = new byte[len];
                    Array.Copy(b, 0, bb, 0, b.Length);
                    o = BitConverter.ToUInt32(bb, 0);
                    break;
                case 2: //8 byte
                    len = 8;
                    bb = new byte[len];
                    Array.Copy(b, 0, bb, 0, b.Length);
                    o = BitConverter.ToUInt64(bb, 0);
                    break;

            }

            return double.Parse(o.ToString());
        }

        public string analyze3GHUUHONG_AutoReport_newprotocol(string value_item, ref DataBLock objCUR, ref string ketqua)
        {

            try
            {
                double doubleValue = 0;
                double douValue1 = 0;
                double douValue2 = 0;
                double douValue3 = 0;
                objCUR.SERIALID = "";
                DateTime dateValue = new DateTime(1000, 1, 1);
                DateTime datetimetemp = DateTime.Now;
                byte[] array = StringToByteArray(value_item);
                // Số Serial 19 20 21 22 23 24
                string _serialMetter = (array[23]).ToString("X2") + (array[22]).ToString("X2") + (array[21]).ToString("X2") + (array[20]).ToString("X2") + (array[19]).ToString("X2") + (array[18]).ToString("X2");
                // thơi gian công tơ 
                string _datetime = array[26].ToString("X2") + array[27].ToString("X2") + array[28].ToString("X2") + array[24].ToString("X2") + array[25].ToString("X2");
                string _datetimebill = "01" + array[27].ToString("X2") + array[28].ToString("X2") + array[24].ToString("X2") + array[25].ToString("X2");
                DateTime _timeChuky = DateTime.ParseExact(_datetime, "ddMMyymmHH", null);
                DateTime _timeBDChuky = DateTime.ParseExact(_datetime, "ddMMyymmHH", null).AddMinutes(-30);

                DateTime _timesKTbill = DateTime.ParseExact(_datetimebill.Substring(0, 6), "ddMMyy", null);
                DateTime _timeBDbill = DateTime.ParseExact(_datetimebill.Substring(0, 6), "ddMMyy", null).AddMonths(-1);
                objCUR.NGAYGIO = _timeChuky;
                objCUR.P_NGAY = _timeChuky;
                objCUR.START_DATE = _timeBDChuky;
                objCUR.END_DATE = _timeChuky;

                objCUR.START_BILL = _timeBDbill;
                objCUR.END_BILL = _timesKTbill;
                objCUR.SERIALID = _serialMetter;
                objCUR.MONTHID = Int32.Parse(array[27].ToString("X2"));
                objCUR.DATEID = Int32.Parse(array[26].ToString("X2"));
                objCUR.YEARID = Int32.Parse("20" + array[28].ToString("X2"));
                objCUR.MINUTEID = Int32.Parse(array[24].ToString("X2"));
                objCUR.HOURID = Int32.Parse(array[25].ToString("X2"));

                string _data = value_item.Substring(60, value_item.Length - 60);
                string commandType;
                string strTempdata;
                int _length = 0;
                while (_data.Length > 4)
                {

                    commandType = _data.Substring(0, 8);
                    _length = Convert.ToInt32(_data.Substring(8, 2), 16);

                    _length = _length * 2;
                    _length = _length & 127;
                    strTempdata = _data.Substring(10, _length);
                    _data = _data.Substring(_length + 10, _data.Length - (_length + 10));
                    doubleValue = 0;
                    douValue1 = 0;
                    douValue2 = 0;
                    douValue3 = 0;
                    dateValue = new DateTime(1000, 1, 1);
                    switch (commandType)
                    {


                        case "606309FF":
                            // Hệ số nhân trên mặt đồng hồ công tơ hay còn goi là đơn vị tính trên mặt đồng hồ công tơ 
                            //ED 03 00 04
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.HSN = doubleValue;
                            break;
                        case "010800FF":
                            // Chỉ số biểu tổng chiều nhận 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_IMPORTKWH = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "010801FF":
                            // Chỉ số biểu 1  
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_IMPBT = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "010802FF":
                            // Chỉ số biểu 2  
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_IMPCD = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "010803FF":
                            // Chỉ số biểu 3  
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_IMPTD = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        // Chỉ số chiều nhận  
                        case "020800FF":
                            // HCT 010801FF //Chỉ số hữu công chiều nhận (khách hàng phát lên lưới ) A- 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_EXPORTKWH = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "020801FF":
                            // HCT 010801FF //Chỉ số hữu công chiều nhận (khách hàng phát lên lưới )
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_EXPBT = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "020802FF":
                            // BIểu 2 010802FF // Điện năng hữu công 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_EXPCD = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "020803FF":
                            // Biểu 3 010803FF // Điện năng hữu công 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_EXPTD = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;





                        // Phân tích dữ liệu load profile 
                        //63 01 00 FF
                        case "630100FF":
                            // 630100FF // Loadprofile P+ P- Q+ Q-
                            analyzeValue_Loadprofile(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.LOAD_P_IMPORTKW = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                            objCUR.LOAD_P_EXPORTKW = douValue1 * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                            objCUR.LOAD_P_C1 = douValue2 * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                            objCUR.LOAD_P_C2 = douValue3 * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                            break;

                        case "030800FF":
                            // HCN 030700FF // Chỉ số vô công giao tổng 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_C1 = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 10000;
                            break;

                        case "040800FF":
                            // 040700FF // chỉ số vô công chiều chiều ngược biểu tổng 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_C2 = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 10000;
                            break;
                        //case "170700FF":
                        //    //  170700FF //// Công suất vô cong pha A chiêu giao (Công suât tức thời ) Q+ Pha A
                        //    //analyzeValue_newpro(strTempdata, ref doubleValue);
                        //    objCUR.P_IMPCD = doubleValue/10000;
                        //    break;
                        //case "2B0700FF":
                        //    //  2B0700FF //// Công suất vô cong pha B chiêu giao (Công suât tức thời ) Q+ Pha B //3F 07 00 FF
                        //    analyzeValue_newpro(strTempdata, ref doubleValue);
                        //    objCUR.P_IMPTD = doubleValue/10000;
                        //    break;
                        //case "3F0700FF":
                        //    //  3F0700FF //// Công suất vô cong pha C chiêu giao (Công suât tức thời ) Q+ Pha C //3F 07 00 FF
                        //    analyzeValue_newpro(strTempdata, ref doubleValue);
                        //    objCUR.P_IMPTD = doubleValue / 10000;
                        //    break;
                        case "01120014":
                            // Thời gian kết thúc ngược pha 
                            analyzeValue_EVENT(strTempdata, ref dateValue);
                            break;

                        case "01080001":
                            // Chỉ số chốt điện năng chiều nhận biểu tổng 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.CHOT_HCNCS = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                            break;

                        case "01080101":
                            // Chỉ số chốt điện năng chiều nhận  biểu 1 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.CHOT_HCNCSB1 = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                            break;
                        case "01080201":
                            // Chỉ số chốt điện năng vô công chiều nhận biểu 2
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.CHOT_HCNCSB2 = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);

                            break;
                        case "01080301":
                            // Chỉ số chốt điện năng vô công chiều nhận biểu 3 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.CHOT_HCNCSB3 = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                            break;

                        case "02080001":
                            // Chi số chốt hữu công chiều nhận biểu tổng (Khach hàng chuyển cho điệnlực ) 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.CHOT_HCNCS = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "02080101":
                            // Chi số chốt hữu công chiều nhận biểu 1
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.CHOT_HCNCSB1 = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "02080201":
                            // Chi số chốt hữu công chiều nhận biểu 2 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.CHOT_HCNCSB2 = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "02080301":
                            // Chi số chốt hữu công chiều nhận biểu 3 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.CHOT_HCNCSB3 = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "03080001":
                            // Chỉ số chốt chỉ số vô công chiều giao biểu tổng
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.CHOT_VCNCS = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "04080001":
                            // Chỉ số chốt điện năng vô công nhận tổng (Nhận của khách hàng )
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.CHOT_VCGCS = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 100;
                            break;
                        case "010600FF":
                            //    PMAX: 33323434 Pmax chiều giao  
                            analyzeValue_pmax_newPro(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3, ref dateValue);
                            objCUR.PMAX_IMPORTKWH = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                            objCUR.PMAX_DATE = dateValue;
                            break;
                        case "020600FF":
                            //    PMAX: 33323534 Pmax chiều nhận 
                            analyzeValue_pmax_newPro(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3, ref dateValue);
                            objCUR.PMAX_IMPORTKWHNHAN = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                            objCUR.PMAX_DATENHAN = dateValue;
                            break;

                        // Active power(*)	01 07 00 FF
                        //  Active power L1(*)	15 07 00 FF
                        // Active power L2(*)	29 07 00 FF
                        // Active power L3(*)	3D 07 00 FF
                        case "010700FF":
                            // HCT 33323433 // Công suấT tổng hữu công chiều thuận P+ 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_AP_T = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 10000; ;
                            break;
                        case "150700FF":
                            // HCT 33323433 // Công suất pha A
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_AP_A = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 10000;
                            break;
                        case "290700FF":
                            // HCT 290700FF // Công suất pha B
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_AP_B = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 10000;
                            break;
                        case "3D0700FF":
                            // HCT 33323433 // Công suất pha c
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_AP_C = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU) / 10000;
                            break;
                        //case "020700FF":
                        //    //  HCN: 020700FF -- Công suất hữu công ngược biểu tổng  P-
                        //    analyzeValue_newpro(strTempdata, ref doubleValue);
                        //    objCUR.P_RP_T = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                        //    break;
                        //case "160700FF":
                        //    //  HCN: 33323533 -- Công suất Hữu công ngược pha A   P- Pha A
                        //    analyzeValue_newpro(strTempdata, ref doubleValue);
                        //    objCUR.P_RP_A = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                        //    break;
                        //case "2A0700FF":
                        //    //  : 2A0700FF -- Công suất Hữu công ngược pha B   P- Pha B
                        //    analyzeValue_newpro(strTempdata, ref doubleValue);
                        //    objCUR.P_RP_B = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                        //    break;
                        //case "3E0700FF":
                        //    //  3E0700FF  -- Hữu công ngược pha C  P- Pha C 
                        //    analyzeValue_newpro(strTempdata, ref doubleValue);
                        //    objCUR.P_RP_C = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) * (objCUR.TI_TU / objCUR.TI_MAU);
                        //    break;
                        case "200700FF":
                            // DIENAP: Pha A 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_V_A = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) / 10;

                            break;
                        case "340700FF":
                            // DIENAP: Pha B 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_V_B = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) / 10;
                            break;
                        case "480700FF":
                            // DIENAP: Pha C 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_V_C = doubleValue * (objCUR.TU_TU / objCUR.TU_MAU) / 10;
                            break;
                        case "600A01FF":
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            // doubleValue 
                            // 0: không xác định  
                            // 1: Đúng thứ tự pha
                            // 2: Sai thứ tự pha

                            break;

                        case "1F0700FF":
                            //   DONGDIEN: Pha A
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_A_A = doubleValue * (objCUR.TI_TU / objCUR.TI_MAU) / 1000;
                            break;
                        case "330700FF":
                            //   DONGDIEN: Pha B
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_A_B = doubleValue * (objCUR.TI_TU / objCUR.TI_MAU) / 1000;
                            break;
                        case "470700FF":
                            //   DONGDIEN: Pha C
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.P_A_C = doubleValue * (objCUR.TI_TU / objCUR.TI_MAU) / 1000;
                            break;
                        case "0D0700FF":
                            // COS_PHI: 0D 07 00 FF -- 
                            analyzeValue_newpro(strTempdata, ref doubleValue); // Sai anh triết ơi 
                            objCUR.COSFI = doubleValue * (objCUR.TI_TU / objCUR.TI_MAU) / 1000;
                            break;
                        case "0E0700FF":
                            //  TAN_SO: Phải chia cho 100
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            break;
                        case "0C010004":
                            analyzeValue_GIO(strTempdata, ref datetimetemp);
                            objCUR.P_CLOCK_METER = datetimetemp;
                            break;
                        case "000403FF": // Tử VT 3C363337
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.TU_TU = doubleValue / 100;


                            break;

                        case "000406FF": // Mẫu số VT 3E363337
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.TU_MAU = doubleValue / 100;
                            break;
                        case "000402FF": // Tử sô : CT 000402FF
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.TI_TU = doubleValue / 100;
                            break;
                        case "000405FF": // Mẫu sô : CT 000405FF
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.TI_MAU = doubleValue / 100;
                            break;
                        case "000408FF ": //Tử/mẫu : VT 000408FF
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.TU = doubleValue / 100;
                            break;
                        case "000409FF  ": //Tử/mẫu : CT 000408FF của TI 
                            analyzeValue_newpro(strTempdata, ref doubleValue);
                            objCUR.TI = doubleValue / 100;
                            break;
                        //case "3F343337":
                        //    analyzeValue_tuti(strTempdata, ref doubleValue);
                        //    objCUR.TI_MAU = doubleValue;
                        //    break;
                        default:
                            break;
                    }
                }

                return "ok";
            }
            catch (Exception ex)
            {
                return "err";
            }



        }




        // 00 FF 07 02 : Goc pha 
        // 

        public string analyze3GHUUHONG_AutoReport(string value_item, ref DataBLock objCUR, ref string ketqua)
        {

            try
            {
                double doubleValue = 0;
                double douValue1 = 0;
                double douValue2 = 0;
                double douValue3 = 0;
                objCUR.SERIALID = "";
                DateTime dateValue = new DateTime(1000, 1, 1);
                DateTime datetimetemp = DateTime.Now;
                byte[] array = StringToByteArray(value_item);
                // Số Serial 19 20 21 22 23 24
                string _serialMetter = (array[23]).ToString("X2") + (array[22]).ToString("X2") + (array[21]).ToString("X2") + (array[20]).ToString("X2") + (array[19]).ToString("X2") + (array[18]).ToString("X2");
                // thơi gian công tơ 
                string _datetime = array[26].ToString("X2") + array[27].ToString("X2") + array[28].ToString("X2") + array[24].ToString("X2") + array[25].ToString("X2");
                string _datetimebill = "01" + array[27].ToString("X2") + array[28].ToString("X2") + array[24].ToString("X2") + array[25].ToString("X2");
                DateTime _timeChuky = DateTime.ParseExact(_datetime, "ddMMyymmHH", null);
                DateTime _timeBDChuky = DateTime.ParseExact(_datetime, "ddMMyymmHH", null).AddMinutes(-30);

                DateTime _timesKTbill = DateTime.ParseExact(_datetimebill.Substring(0, 6), "ddMMyy", null);
                DateTime _timeBDbill = DateTime.ParseExact(_datetimebill.Substring(0, 6), "ddMMyy", null).AddMonths(-1);
                objCUR.NGAYGIO = _timeChuky;
                objCUR.P_NGAY = _timeChuky;
                objCUR.START_DATE = _timeBDChuky;
                objCUR.END_DATE = _timeChuky;

                objCUR.START_BILL = _timeBDbill;
                objCUR.END_BILL = _timesKTbill;
                objCUR.SERIALID = _serialMetter;
                objCUR.MONTHID = Int32.Parse(array[27].ToString("X2"));
                objCUR.DATEID = Int32.Parse(array[26].ToString("X2"));
                objCUR.YEARID = Int32.Parse("20" + array[28].ToString("X2"));
                objCUR.MINUTEID = Int32.Parse(array[24].ToString("X2"));
                objCUR.HOURID = Int32.Parse(array[25].ToString("X2"));

                string _data = value_item.Substring(60, value_item.Length - 60);
                string commandType;
                string strTempdata;
                int _length = 0;
                while (_data.Length > 4)
                {

                    commandType = _data.Substring(0, 8);

                    _length = Convert.ToInt32(_data.Substring(8, 2), 16);
                    _length = _length * 2;
                    strTempdata = _data.Substring(10, _length);
                    _data = _data.Substring(_length + 10, _data.Length - (_length + 10));
                    doubleValue = 0;
                    douValue1 = 0;
                    douValue2 = 0;
                    douValue3 = 0;
                    dateValue = new DateTime(1000, 1, 1);
                    switch (commandType)
                    {

                        case "ED030004":
                            // Hệ số nhân trên mặt đồng hồ công tơ hay còn goi là đơn vị tính trên mặt đồng hồ công tơ 
                            //ED 03 00 04
                            analyzeValue_autoDonvitinh(strTempdata, ref doubleValue);
                            objCUR.HSN = doubleValue;
                            break;
                        case "01010014":
                            // Thời gian bắt đầu ngược pha
                            analyzeValue_EVENT(strTempdata, ref dateValue);
                            break;
                        case "01120014":
                            // Thời gian kết thúc ngược pha 
                            analyzeValue_EVENT(strTempdata, ref dateValue);
                            break;

                        case "01FF0300":
                            // chi số chốt vô công chiều giao 
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.CHOT_VCGCS = doubleValue;
                            objCUR.CHOT_VCGCSB1 = douValue1;
                            objCUR.CHOT_VCGCSB2 = douValue2;
                            objCUR.CHOT_VCGCSB3 = douValue3;
                            break;
                        case "01FF0400":
                            // Chỉ số chốt điện năng vô công chiều nhận 
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.CHOT_VCNCS = doubleValue;
                            objCUR.CHOT_VCNCSB1 = douValue1;
                            objCUR.CHOT_VCNCSB2 = douValue2;
                            objCUR.CHOT_VCNCSB3 = douValue3;
                            break;
                        case "01FF0100":
                            // Chi số chốt hữu công chiều giao
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.CHOT_HCGCS = doubleValue;
                            objCUR.CHOT_HCGCSB1 = douValue1;
                            objCUR.CHOT_HCGCSB2 = douValue2;
                            objCUR.CHOT_HCGCSB3 = douValue3;
                            break;
                        case "01FF0200":
                            // Chỉ số chốt điện năng hữu công chiều nhận
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.CHOT_HCNCS = doubleValue;
                            objCUR.CHOT_HCNCSB1 = douValue1;
                            objCUR.CHOT_HCNCSB2 = douValue2;
                            objCUR.CHOT_HCNCSB3 = douValue3;
                            break;
                        case "01FF0000":
                            // Chi số chốt hữu công tổng 
                            //analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            //objCUR.P_IMPORTKWH = doubleValue;
                            //objCUR.P_IMPBT = douValue1;
                            //objCUR.P_IMPCD = douValue2;
                            //objCUR.P_IMPTD = douValue3;
                            break;
                        case "01000006":
                            // Loadprofile 34333339
                            analyzeLoadProfile_auto(strTempdata, ref objCUR);
                            break;
                        case "00FF0101":
                            //    PMAX: 33323434 Pmax chiều giao  
                            analyzeValue_pmax_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3, ref dateValue);
                            objCUR.PMAX_IMPORTKWH = doubleValue;
                            objCUR.PMAX_DATE = dateValue;
                            break;
                        case "00FF0201":
                            //    PMAX: 33323534 Pmax chiều nhận 
                            analyzeValue_pmax_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3, ref dateValue);
                            objCUR.PMAX_IMPORTKWHNHAN = doubleValue;
                            objCUR.PMAX_DATENHAN = dateValue;
                            break;
                        case "00FF0100":
                            // HCT 33323433 // Điện năng hữu công 
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_IMPORTKWH = doubleValue;
                            objCUR.P_IMPBT = douValue1;
                            objCUR.P_IMPCD = douValue2;
                            objCUR.P_IMPTD = douValue3;
                            break;
                        case "00FF0200":
                            //  HCN: 33323533 -- 00 FF 02 00 
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_EXPORTKWH = doubleValue;
                            objCUR.P_EXPBT = douValue1;
                            objCUR.P_EXPCD = douValue2;
                            objCUR.P_EXPTD = douValue3;
                            break;
                        case "00000300":
                            //  VCT: 33323533 -- Lech anh triết ơi 
                            analyzeValue_QTong(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_C1 = doubleValue;
                            break;
                        case "00000400":
                            //  VCN: 33323733 00 00 04 00 -- Lệch anh triết ơi
                            analyzeValue_QTong(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_C2 = doubleValue;
                            break;
                        case "00FF0102":
                            // DIENAP: 33323435 
                            analyzeValue_Dienap_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_V_A = douValue1;
                            objCUR.P_V_B = douValue2;
                            objCUR.P_V_C = douValue3;
                            break;
                        case "00FF0202":
                            //   DONGDIEN: 33323535
                            analyzeValue_Dongdien_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_A_A = douValue1;
                            objCUR.P_A_B = douValue2;
                            objCUR.P_A_C = douValue3;
                            break;
                        case "00FF0302":
                            //  P_HC: 33323635 //Công suất huu công 
                            analyzeValue_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_AP_T = doubleValue;
                            objCUR.P_AP_A = douValue1;
                            objCUR.P_AP_B = douValue2;
                            objCUR.P_AP_C = douValue3;
                            break;
                        case "00FF0402":
                            // P_VC: 33323735 // công suất vô công 
                            analyzeValue_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_RP_T = doubleValue;
                            objCUR.P_RP_A = douValue1;
                            objCUR.P_RP_B = douValue2;
                            objCUR.P_RP_C = douValue3;
                            break;
                        case "00FF0602":
                            // COS_PHI: 33323935 -- Lệch anh triết ơi 
                            analyzeValue_cosfi_auto(strTempdata, ref doubleValue); // Sai anh triết ơi 
                            objCUR.COSFI = doubleValue;
                            break;
                        case "3533B335":
                            //  TAN_SO: 3533B335
                            break;
                        case "0C010004":
                            analyzeValue_GIO(strTempdata, ref datetimetemp);
                            objCUR.P_CLOCK_METER = datetimetemp;
                            break;
                        case "08030004": // Tử VT 3C363337
                            analyzeValue_tuti_auto(strTempdata, ref doubleValue);
                            objCUR.TU_TU = doubleValue;
                            break;

                        case "0A030004": // Mẫu số VT 3E363337
                            analyzeValue_tuti_auto(strTempdata, ref doubleValue);
                            objCUR.TU_MAU = doubleValue;
                            break;
                        case "09030004": // Tử sô : CT 3B363337
                            analyzeValue_tuti_auto(strTempdata, ref doubleValue);
                            objCUR.TI_TU = doubleValue;
                            break;
                        case "0B030004": // Tử sô : CT 3D363337
                            analyzeValue_tuti_auto(strTempdata, ref doubleValue);
                            objCUR.TI_MAU = doubleValue;
                            break;


                        //case "3F343337":
                        //    analyzeValue_tuti(strTempdata, ref doubleValue);
                        //    objCUR.TI_MAU = doubleValue;
                        //    break;
                        default:
                            break;
                    }
                }

                return "ok";
            }
            catch (Exception ex)
            {
                return "err";
            }



        }


        public string analyze3GHUUHONG_AutoReport(string value_item, ref dynamic objLoad, ref dynamic obj, ref dynamic objHIS, ref dynamic objCUR, ref string ketqua)
        {

            try
            {
                double doubleValue = 0;
                double douValue1 = 0;
                double douValue2 = 0;
                double douValue3 = 0;
                objCUR.SERIALID = "";
                DateTime dateValue = new DateTime(1000, 1, 1);
                DateTime datetimetemp = DateTime.Now;
                byte[] array = StringToByteArray(value_item);
                // Số Serial 19 20 21 22 23 24
                string _serialMetter = (array[23]).ToString("X2") + (array[22]).ToString("X2") + (array[21]).ToString("X2") + (array[20]).ToString("X2") + (array[19]).ToString("X2") + (array[18]).ToString("X2");
                // thơi gian công tơ 
                string _datetime = array[26].ToString("X2") + array[27].ToString("X2") + array[28].ToString("X2") + array[24].ToString("X2") + array[25].ToString("X2");
                DateTime _timeChuky = DateTime.ParseExact(_datetime, "ddMMyymmHH", null);
                objCUR.NGAYGIO = _timeChuky;
                objCUR.SERIALID = _serialMetter;

                string _data = value_item.Substring(60, value_item.Length - 60);
                string commandType;
                string strTempdata;
                int _length = 0;
                while (_data.Length > 4)
                {

                    commandType = _data.Substring(0, 8);

                    _length = Convert.ToInt32(_data.Substring(8, 2), 16);
                    _length = _length * 2;
                    strTempdata = _data.Substring(10, _length);
                    _data = _data.Substring(_length + 10, _data.Length - (_length + 10));
                    doubleValue = 0;
                    douValue1 = 0;
                    douValue2 = 0;
                    douValue3 = 0;
                    dateValue = new DateTime(1000, 1, 1);
                    switch (commandType)
                    {

                        case "ED030004":
                            // Hệ số nhân trên mặt đồng hồ công tơ hay còn goi là đơn vị tính trên mặt đồng hồ công tơ 
                            //ED 03 00 04
                            analyzeValue_autoDonvitinh(strTempdata, ref doubleValue);
                            objCUR.HSN = doubleValue;
                            break;
                        case "01010014":
                            // Thời gian bắt đầu ngược pha
                            analyzeValue_EVENT(strTempdata, ref dateValue);
                            break;
                        case "01120014":
                            // Thời gian kết thúc ngược pha 
                            analyzeValue_EVENT(strTempdata, ref dateValue);
                            break;

                        case "01FF0300":
                            // chi số chốt vô công chiều giao 
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.CHOT_VCGCS = doubleValue;
                            objCUR.CHOT_VCGCSB1 = douValue1;
                            objCUR.CHOT_VCGCSB2 = douValue2;
                            objCUR.CHOT_VCGCSB3 = douValue3;
                            break;
                        case "01FF0400":
                            // Chỉ số chốt điện năng vô công chiều nhận 
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.CHOT_VCNCS = doubleValue;
                            objCUR.CHOT_VCNCSB1 = douValue1;
                            objCUR.CHOT_VCNCSB2 = douValue2;
                            objCUR.CHOT_VCNCSB3 = douValue3;
                            break;
                        case "01FF0100":
                            // Chi số chốt hữu công chiều giao  : Hay còn gọi là chiều ngược : tiếng anh gọi là Reactive power bill
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.CHOT_HCGCS = doubleValue;
                            objCUR.CHOT_HCGCSB1 = douValue1;
                            objCUR.CHOT_HCGCSB2 = douValue2;
                            objCUR.CHOT_HCGCSB3 = douValue3;
                            break;
                        case "01FF0200":
                            // Chỉ số chốt điện năng hữu công chiều nhận
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.CHOT_HCNCS = doubleValue;
                            objCUR.CHOT_HCNCSB1 = douValue1;
                            objCUR.CHOT_HCNCSB2 = douValue2;
                            objCUR.CHOT_HCNCSB3 = douValue3;
                            break;
                        case "01FF0000":
                            // Chi số chốt hữu công tổng 
                            //analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            //objCUR.P_IMPORTKWH = doubleValue;
                            //objCUR.P_IMPBT = douValue1;
                            //objCUR.P_IMPCD = douValue2;
                            //objCUR.P_IMPTD = douValue3;
                            break;
                        case "01000006":
                            // Loadprofile 34333339
                            analyzeLoadProfile_auto(strTempdata, ref objCUR);
                            break;
                        case "00FF0101":
                            //    PMAX: 33323434 Pmax chiều giao  
                            analyzeValue_pmax_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3, ref dateValue);
                            objCUR.PMAX_IMPORTKWH = doubleValue;
                            objCUR.PMAX_DATE = dateValue;
                            break;
                        case "00FF0201":
                            //    PMAX: 33323534 Pmax chiều nhận 
                            analyzeValue_pmax_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3, ref dateValue);
                            objCUR.PMAX_IMPORTKWHNHAN = doubleValue;
                            objCUR.PMAX_DATENHAN = dateValue;
                            break;
                        case "00FF0100":
                            // HCT 33323433 // Điện năng hữu công 
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_IMPORTKWH = doubleValue;
                            objCUR.P_IMPBT = douValue1;
                            objCUR.P_IMPCD = douValue2;
                            objCUR.P_IMPTD = douValue3;
                            break;
                        case "00FF0200":  // Điện năng hữu công chiều giao 
                            //  HCN: 33323533 -- 00 FF 02 00 
                            analyzeValue_auto4Byte(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_EXPORTKWH = doubleValue;
                            objCUR.P_EXPBT = douValue1;
                            objCUR.P_EXPCD = douValue2;
                            objCUR.P_EXPTD = douValue3;
                            break;
                        case "00000300":
                            //  VCT: 33323533 -- Lech anh triết ơi 
                            analyzeValue_QTong(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_C1 = doubleValue;
                            break;
                        case "00000400":
                            //  VCN: 33323733 00 00 04 00 -- Lệch anh triết ơi
                            analyzeValue_QTong(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_C2 = doubleValue;
                            break;
                        case "00FF0102":
                            // DIENAP: 33323435 
                            analyzeValue_Dienap_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_V_A = douValue1;
                            objCUR.P_V_B = douValue2;
                            objCUR.P_V_C = douValue3;
                            break;
                        case "00FF0202":
                            //   DONGDIEN: 33323535
                            analyzeValue_Dongdien_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_A_A = douValue1;
                            objCUR.P_A_B = douValue2;
                            objCUR.P_A_C = douValue3;
                            break;
                        case "00FF0302":
                            //  P_HC: 33323635 //Công suất huu công 
                            analyzeValue_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_AP_T = doubleValue;
                            objCUR.P_AP_A = douValue1;
                            objCUR.P_AP_B = douValue2;
                            objCUR.P_AP_C = douValue3;
                            break;
                        case "00FF0402":
                            // P_VC: 33323735 // công suất vô công 
                            analyzeValue_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_RP_T = doubleValue;
                            objCUR.P_RP_A = douValue1;
                            objCUR.P_RP_B = douValue2;
                            objCUR.P_RP_C = douValue3;
                            break;
                        case "00FF0602":
                            // COS_PHI: 33323935 -- Lệch anh triết ơi 
                            analyzeValue_cosfi_auto(strTempdata, ref doubleValue); // Sai anh triết ơi 
                            objCUR.COSFI = doubleValue;
                            break;
                        case "3533B335":
                            //  TAN_SO: 3533B335
                            break;
                        case "0C010004":
                            analyzeValue_GIO(strTempdata, ref datetimetemp);
                            objCUR.P_CLOCK_METER = datetimetemp;
                            break;
                        case "08030004": // Tử VT 3C363337
                            analyzeValue_tuti_auto(strTempdata, ref doubleValue);
                            objCUR.TU_TU = doubleValue;
                            break;

                        case "0A030004": // Mẫu số VT 3E363337
                            analyzeValue_tuti_auto(strTempdata, ref doubleValue);
                            objCUR.TU_MAU = doubleValue;
                            break;
                        case "09030004": // Tử sô : CT 3B363337
                            analyzeValue_tuti_auto(strTempdata, ref doubleValue);
                            objCUR.TI_TU = doubleValue;
                            break;
                        case "0B030004": // Tử sô : CT 3D363337
                            analyzeValue_tuti_auto(strTempdata, ref doubleValue);
                            objCUR.TI_MAU = doubleValue;
                            break;


                        //case "3F343337":
                        //    analyzeValue_tuti(strTempdata, ref doubleValue);
                        //    objCUR.TI_MAU = doubleValue;
                        //    break;
                        default:
                            break;
                    }
                }

                return "ok";
            }
            catch (Exception ex)
            {
                return "err";
            }



        }

        public bool UploadBulkDataLP(List<LP> bulkData, OracleConnection cnn)
        {
            bool returnValue = false;
            try
            {
                if (cnn != null && cnn.State == ConnectionState.Closed)
                    cnn.Open();

                // var cmd = oracleConnection.CreateCommand();

                // cmd.CommandText = "SELECT NVL (A.ASSETID, ' '), NVL (Z.KIEU_CTO, ' '),A.SERIALNUM FROM A_ASSET A JOIN ZAG_MDMS_DIEMDO Z ON A.ASSETID = Z.OBJID";
                // cmd.CommandType = CommandType.Text;
                // OracleDataReader dr = cmd.ExecuteReader();
                // dr.Read();
                //string test = dr.GetString(0);




                string query = @"insert into EVNHES.LP (MeterSerial,P_C,P_T,Q_C,Q_T,TimeData) values (:MeterSerial,:P_C,:P_T,:Q_C,:Q_T,:TimeData)";

                using (var command = cnn.CreateCommand())
                {
                    command.CommandText = query;
                    command.CommandType = CommandType.Text;
                    command.BindByName = true;
                    // In order to use ArrayBinding, the ArrayBindCount property
                    // of OracleCommand object must be set to the number of records to be inserted
                    command.ArrayBindCount = bulkData.Count;
                    command.Parameters.Add(":MeterSerial", OracleDbType.Varchar2, bulkData.Select(c => c.MeterSerial).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":P_C", OracleDbType.Varchar2, bulkData.Select(c => c.P_C).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":P_T", OracleDbType.Varchar2, bulkData.Select(c => c.P_T).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":Q_C", OracleDbType.Varchar2, bulkData.Select(c => c.Q_C).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":Q_T", OracleDbType.Varchar2, bulkData.Select(c => c.Q_T).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":TimeData", OracleDbType.Date, bulkData.Select(c => c.TimeData).ToArray(), ParameterDirection.Input);
                    int result = command.ExecuteNonQuery();
                    if (result == bulkData.Count)
                        returnValue = true;
                }
            }
            catch (OracleException ex)
            {
                //Log error thrown
            }
            finally
            {
                cnn.Close();
            }
            return returnValue;
        }

        public bool UploadBulkDataVcumvals(List<LP> bulkData, OracleConnection cnn)
        {
            bool returnValue = false;
            try
            {
                if (cnn != null && cnn.State == ConnectionState.Closed)
                    cnn.Open();

                // var cmd = oracleConnection.CreateCommand();

                // cmd.CommandText = "SELECT NVL (A.ASSETID, ' '), NVL (Z.KIEU_CTO, ' '),A.SERIALNUM FROM A_ASSET A JOIN ZAG_MDMS_DIEMDO Z ON A.ASSETID = Z.OBJID";
                // cmd.CommandType = CommandType.Text;
                // OracleDataReader dr = cmd.ExecuteReader();
                // dr.Read();
                //string test = dr.GetString(0);




                string query = @"insert into EVNHES.LP (MeterSerial,P_C,P_T,Q_C,Q_T,TimeData) values (:MeterSerial,:P_C,:P_T,:Q_C,:Q_T,:TimeData)";

                using (var command = cnn.CreateCommand())
                {
                    command.CommandText = query;
                    command.CommandType = CommandType.Text;
                    command.BindByName = true;
                    // In order to use ArrayBinding, the ArrayBindCount property
                    // of OracleCommand object must be set to the number of records to be inserted
                    command.ArrayBindCount = bulkData.Count;
                    command.Parameters.Add(":MeterSerial", OracleDbType.Varchar2, bulkData.Select(c => c.MeterSerial).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":P_C", OracleDbType.Varchar2, bulkData.Select(c => c.P_C).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":P_T", OracleDbType.Varchar2, bulkData.Select(c => c.P_T).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":Q_C", OracleDbType.Varchar2, bulkData.Select(c => c.Q_C).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":Q_T", OracleDbType.Varchar2, bulkData.Select(c => c.Q_T).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":TimeData", OracleDbType.Date, bulkData.Select(c => c.TimeData).ToArray(), ParameterDirection.Input);
                    int result = command.ExecuteNonQuery();
                    if (result == bulkData.Count)
                        returnValue = true;
                }
            }
            catch (OracleException ex)
            {
                //Log error thrown
            }
            finally
            {
                cnn.Close();
            }
            return returnValue;
        }

        public string GenerateNumber()
        {
            Random random = new Random();
            string r = "";
            int i;
            for (i = 1; i < 11; i++)
            {
                r += random.Next(0, 9).ToString();
            }
            return r;
        }

        public void do_analyze_3phaHuuhong(string path, RichTextBox rtbLog)
        {
            try
            {
                string sourcePath = @"D:\EVNHES_3G";
                System.IO.Directory.CreateDirectory(@"D:\EVNHES_3G\RES_ERR\" + DateTime.Now.ToString("ddMMyy"));
                string targetPath = @"D:\EVNHES_3G\RES_ERR\" + DateTime.Now.ToString("ddMMyy");
                string oradb = General.connectString;
                OracleConnection conn = new OracleConnection(oradb);  // C#
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                dynamic obj = new ExpandoObject();
                dynamic objHIS = new ExpandoObject();
                dynamic objCUR = new ExpandoObject();
                dynamic objLoad = new ExpandoObject();
                dynamic objCosfi = new ExpandoObject();


                string error = "";
                string[] files = System.IO.Directory.GetFiles(path, "*.res");
                int result = -1;


                /* test bug insert 
                 */
                List<CustomerDTO> bulkData = new List<CustomerDTO>();
                List<LP> bulkDataLP = new List<LP>();

                DateTime dateTime = new DateTime(2018, 7, 13);
                String Serial = GenerateNumber();


                for (int i = 0; i < 10; i++)
                {
                    Serial = GenerateNumber();
                    dateTime = dateTime.AddHours(30);
                    bulkDataLP.Add(new LP(Serial, "Ayobami", "Adewole", "xyzcom", "xyzcom", dateTime));
                    if (i == 100)
                    {
                        dateTime = new DateTime(2018, 7, 13);
                    }
                }

                UploadBulkDataLP(bulkDataLP, conn);



                foreach (string filename in files)
                {
                    string strNgaygio = "";
                    //    string fn = new FileInfo(filename).Name;
                    //   strNgaygio = fn.Substring(fn.IndexOf("_")+1, 10);
                    string sourceFile = System.IO.Path.Combine(sourcePath, filename);
                    string destFile = System.IO.Path.Combine(targetPath);
                    string ketqua = "";
                    int _startNum;


                    string values = System.IO.File.ReadAllText(filename);
                    char[] delimiters = new char[] { '\r', '\n' };

                    string strTempdata = "";
                    string[] value_item = values.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    string tenfile = "";

                    double doubleValue = 0;
                    double douValue1 = 0;
                    double douValue2 = 0;
                    double douValue3 = 0;
                    DateTime dateValue = new DateTime(1000, 1, 1);

                    objCUR.P_V_A = null;
                    objCUR.P_V_B = null;
                    objCUR.P_V_C = null;
                    objCUR.P_A_A = null;
                    objCUR.P_A_B = null;
                    objCUR.P_A_C = null;
                    objCUR.TU_TU = null;
                    objCUR.TU_MAU = null;
                    objCUR.TI_TU = null;
                    objCUR.TI_MAU = null;
                    objCUR.HSN = null;

                    tenfile = Path.GetFileName(filename);

                    General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                    File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, tenfile), true);
                    System.IO.File.Delete(filename);

                    if (result != -1)
                    {
                        General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                        System.IO.File.Delete(filename);
                    }


                    if (value_item[0].Contains("Vinasino"))
                    {
                        analyzeVinasino(value_item, filename);
                    }
                    else
                    {
                        if (value_item[0].ToString().Count() == 12)
                        {
                            objCUR.P_MA_DIEMDO = value_item[0].ToString();
                            _startNum = 1;
                            goto SUCCESS;
                        }
                        else
                        {
                            objCUR.P_MA_DIEMDO = "";
                            _startNum = 0;

                        }

                        for (int i = _startNum; i < value_item.Length; i++)
                        {

                            //obj.P_SERIALID = value_item[2].Split(new char[] { '\t' })[1];
                            //  strNgaygio = value_item[1].Split(new char[] { '\t' })[1];
                            byte[] array = StringToByteArray(value_item[i].ToString());

                            // Message Event co mã hóa  : 0c; 08;1B --> 
                            // Sự kiện Mất điện số lần mất điện thời gian mất điện (Thời gian bắt đầu ; Thời gian kết thúc)
                            // Công tơ bị thay đổi thời gian trả về 10 lần gần nhất 
                            // Pin yếu 
                            // Mất cân bằng pha
                            // Ngược thứ pha
                            // Ngược dòng pha A
                            // Ngược dòng pha B 
                            // Ngược dòng pha C
                            // Mất pha A
                            // Mất pha B
                            // Mất pha C
                            // Thay đổi pass công tơ 
                            // Số lần lập trình 


                            // Kiem tra auto report 
                            if (array[12] == 0x10 && array[16] == 0x08 && array[17] == 0x0C)
                            {
                                if (Processing.CheckFrame(value_item[i].ToString()) == true)
                                {
                                    value_item[i] = FrameProcessing.Processing.StringDataDecryption(value_item[i].ToString(), "613E3367F67B8E24F09117B445191C92");
                                }
                                // Trong trường hợp này thì chuyển luồng 
                                if (value_item[i].ToString().Length > 100)
                                {
                                    analyze3GHUUHONG_AutoReport(value_item[i].ToString(), ref objLoad, ref obj, ref objHIS, ref objCUR, ref ketqua);
                                    insert_Imiscumvals(conn, objCUR);
                                    insert_Loadprofile(conn, objCUR);
                                    insert_Imishistory(conn, objCUR);
                                    insert_ImisPmax(conn, objCUR);
                                    goto SUCCESS;
                                }
                                else
                                    goto SUCCESS;

                            }

                            if (array[12] == 0x0C && array[16] == 0x80 && array[17] == 0x18)
                            {
                                // Trong trường hợp này thì chuyển luồng 
                                if (value_item[i].ToString().Length > 100)
                                {
                                    analyze3GHUUHONG_AutoReport(value_item[i].ToString(), ref objLoad, ref obj, ref objHIS, ref objCUR, ref ketqua);
                                    insert_Imiscumvals(conn, objCUR);
                                    insert_Loadprofile(conn, objCUR);
                                    insert_Imishistory(conn, objCUR);
                                    insert_ImisPmax(conn, objCUR);
                                    goto SUCCESS;
                                }
                                else
                                    goto SUCCESS;

                            }

                            strTempdata = value_item[i].ToString().Substring(116, value_item[i].ToString().Length - 116);

                            string dataType = array[56].ToString("X2") + array[57].ToString("X2");
                            string commandType = array[59].ToString("X2") + array[60].ToString("X2") + array[61].ToString("X2") + array[62].ToString("X2");
                            switch (dataType)
                            {
                                case "6891":
                                case "68B1":
                                    switch (commandType)
                                    {
                                        case "33323434":
                                            //    PMAX: 33323434
                                            analyzeValue_pmax(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3, ref dateValue);
                                            objCUR.PMAX_IMPORTKWH = doubleValue;
                                            objCUR.PMAX_DATE = dateValue;
                                            break;
                                        case "33323433":
                                            // HCT 

                                            analyzeValue_new(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                            objCUR.P_IMPORTKWH = doubleValue;
                                            objCUR.P_IMPBT = douValue1;
                                            objCUR.P_IMPCD = douValue2;
                                            objCUR.P_IMPTD = douValue3;
                                            break;
                                        case "33323533":
                                            //  HCN: 33323533
                                            analyzeValue_new(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                            objCUR.P_EXPORTKWH = doubleValue;
                                            objCUR.P_EXPBT = douValue1;
                                            objCUR.P_EXPCD = douValue2;
                                            objCUR.P_EXPTD = douValue3;
                                            break;
                                        case "33323633":
                                            //  VCT: 33323533
                                            analyzeValue_new(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                            objCUR.P_C1 = doubleValue;
                                            break;
                                        case "33323733":
                                            //  VCN: 33323733
                                            break;
                                        case "33323435":
                                            // DIENAP: 33323435
                                            analyzeValue_Dienap_auto(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                            objCUR.P_V_A = douValue1;
                                            objCUR.P_V_B = douValue2;
                                            objCUR.P_V_C = douValue3;
                                            break;
                                        case "33323535":
                                            //   DONGDIEN: 33323535
                                            analyzeValue_Dongdien_new(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                            objCUR.P_A_A = douValue1;
                                            objCUR.P_A_B = douValue2;
                                            objCUR.P_A_C = douValue3;
                                            break;
                                        case "33323635":
                                            //  P_HC: 33323635
                                            analyzeValue_new(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                            objCUR.P_AP_T = doubleValue;
                                            objCUR.P_AP_A = douValue1;
                                            objCUR.P_AP_B = douValue2;
                                            objCUR.P_AP_C = douValue3;
                                            break;
                                        case "33323735":
                                            // P_VC: 33323735
                                            analyzeValue_new(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                            objCUR.P_RP_T = doubleValue;
                                            objCUR.P_RP_A = douValue1;
                                            objCUR.P_RP_B = douValue2;
                                            objCUR.P_RP_C = douValue3;
                                            break;
                                        case "33323935":
                                            // COS_PHI: 33323935
                                            analyzeValue_cosfi_new(strTempdata, ref doubleValue);
                                            objCUR.COSFI = doubleValue;
                                            break;
                                        case "3533B335":
                                            //  TAN_SO: 3533B335
                                            break;
                                        case "3F343337":
                                            //   GIO_CT:3F343337
                                            //analyzeValue_GIO(strTempdata, ref datetimetemp);
                                            //objCUR.P_CLOCK_METER = datetimetemp;
                                            break;
                                        case "3C363337": // Tử VT
                                            analyzeValue_tuti(strTempdata, ref doubleValue);
                                            objCUR.TU_TU = doubleValue;
                                            break;

                                        case "3E363337": // Mẫu số VT
                                            analyzeValue_tuti(strTempdata, ref doubleValue);
                                            objCUR.TU_MAU = doubleValue;
                                            break;
                                        case "3B363337": // Tử sô : CT
                                            analyzeValue_tuti(strTempdata, ref doubleValue);
                                            objCUR.TI_TU = doubleValue;
                                            break;
                                        case "3D363337": // Tử sô : CT
                                            analyzeValue_tuti(strTempdata, ref doubleValue);
                                            objCUR.TI_MAU = doubleValue;
                                            break;
                                            //case "3F343337":
                                            //    analyzeValue_tuti(strTempdata, ref doubleValue);
                                            //    objCUR.TI_MAU = doubleValue;
                                            //    break;
                                    }
                                    break;
                                case "6893":
                                case "68B3":
                                    objCUR.P_SERIALID = (array[64] - 0x33).ToString("X2") + (array[63] - 0x33).ToString("X2") + (array[62] - 0x33).ToString("X2") + (array[61] - 0x33).ToString("X2") + (array[60] - 0x33).ToString("X2") + (array[59] - 0x33).ToString("X2");
                                    break;
                                default:
                                    break;

                            }

                        }// endfor0 PD1600T000036001
                        try
                        {
                            OracleCommand cmd = new OracleCommand("insert_imis_vcumvals_tnv", conn);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@p_MA_DIEMDO", OracleDbType.NVarchar2).Value = objCUR.P_MA_DIEMDO;
                            cmd.Parameters.Add("@p_IMEI", OracleDbType.NVarchar2).Value = objCUR.P_SERIALID;
                            cmd.Parameters.Add("@p_NGAYGIO", OracleDbType.Date).Value = System.DateTime.Now; // tạm thời lấy sysdate để kiểm tra sẽ phải add vào sau khi có trường da
                            cmd.Parameters.Add("@p_CLOCK_METER", OracleDbType.Date).Value = System.DateTime.Now;// tạm thời lấy sysdate để kiểm tra sẽ phải add vào sau khi có trường da
                            cmd.Parameters.Add("@p_NGAY", OracleDbType.Date).Value = System.DateTime.Now;// tạm thời lấy sysdate để kiểm tra sẽ phải add vào sau khi có trường da

                            cmd.Parameters.Add("@p_IMPORTKWH", OracleDbType.Double).Value = objCUR.P_IMPORTKWH;
                            cmd.Parameters.Add("@p_EXPORTKWH", OracleDbType.Double).Value = objCUR.P_EXPORTKWH;
                            cmd.Parameters.Add("@p_C1", OracleDbType.Double).Value = objCUR.P_C1;
                            cmd.Parameters.Add("@p_IMPBT", OracleDbType.Double).Value = objCUR.P_IMPBT;
                            cmd.Parameters.Add("@p_IMPCD", OracleDbType.Double).Value = objCUR.P_IMPCD;
                            cmd.Parameters.Add("@p_IMPTD", OracleDbType.Double).Value = objCUR.P_IMPTD;
                            cmd.Parameters.Add("@p_V_A", OracleDbType.Double).Value = objCUR.P_V_A;
                            cmd.Parameters.Add("@p_V_B", OracleDbType.Double).Value = objCUR.P_V_B;
                            cmd.Parameters.Add("@p_V_C", OracleDbType.Double).Value = objCUR.P_V_C;
                            cmd.Parameters.Add("@p_A_A", OracleDbType.Double).Value = objCUR.P_A_A;
                            cmd.Parameters.Add("@p_A_B", OracleDbType.Double).Value = objCUR.P_A_B;
                            cmd.Parameters.Add("@p_A_C", OracleDbType.Double).Value = objCUR.P_A_C;
                            cmd.Parameters.Add("@p_TU_IN", OracleDbType.NVarchar2).Value = objCUR.TU_TU / objCUR.TU_MAU;
                            cmd.Parameters.Add("@p_TI_IN", OracleDbType.NVarchar2).Value = objCUR.TI_TU / objCUR.TI_MAU;
                            cmd.ExecuteNonQuery();
                        }

                        catch (Exception ex)
                        {
                            // MessageBox.Show(ex.Message);
                            conn.Close();
                        }

                        conn.Close();

                    SUCCESS:;

                        #region luu
                        /*
                       *       strTemp = strNgaygio.Substring(10, 2);
                          if (Convert.ToInt32(strTemp) < 30)
                          {
                              strTemp = "00";
                          }
                          else
                          {
                              strTemp = "30";
                          }

                          obj.P_NGAYGIO = DateTime.ParseExact(strNgaygio.Substring(0, 10) + strTemp + "00", "ddMMyyyyHHmmss", null);
                          //  obj.P_NGAYGIO = DateTime.Now;
                          obj.P_NGAY = DateTime.Now;

                          switch (value_item[0].Split(new char[] { '\t' })[1])
                          {
                              case "DCU_IFC":
                                  analyzeIFC(value_item, ref obj);
                                  break;
                              case "DCU_EMEC":
                                  analyzeEMEC(value_item, ref obj);
                                  result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_ROUTING_PATH, ref obj, ref error);
                                  break;
                              case "DTRFEXT":
                                  List<DataPQmax> DataPQmax = new List<DataPQmax>();
                                  List<DataChiso> DataChiso = new List<DataChiso>();
                                  List<DataLoadprofile> ListLoadprofile = new List<DataLoadprofile>();
                                  List<DataEvent> ListEvent = new List<DataEvent>();

                                  analyzeHUUHONG_NEW(value_item, ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListEvent, ref obj);
                                  result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.INSERT_CHISO_1PHA_HUUHONG, ref obj, ref error);
                                  tenfile = System.DateTime.Now.Millisecond.ToString() + "_" + System.DateTime.Now.Date.ToString("ddMMyy") + "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString();
                                  File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, tenfile + ".res"), true);
                                  break;
                              case "HUUHONG_3G":

                                  obj.P_SERIALID = value_item[2].Split(new char[] { '\t' })[1];



                                  objCUR.P_MA_DIEMDO = obj.P_MA_DIEMDO;
                                  objHIS.P_MA_DIEMDO = obj.P_MA_DIEMDO;
                                  objLoad.P_MA_DIEMDO = obj.P_MA_DIEMDO;
                                  objCUR.P_NGAYGIO = obj.P_NGAYGIO;
                                  objHIS.P_NGAYGIO = obj.P_NGAYGIO;
                                  objLoad.P_NGAYGIO = obj.P_NGAYGIO;
                                  objCUR.P_NGAY = obj.P_NGAY;
                                  objHIS.P_NGAY = obj.P_NGAY;
                                  objLoad.P_NGAY = obj.P_NGAY;
                                  objCUR.P_SERIALID = obj.P_SERIALID;
                                  objHIS.P_SERIALID = obj.P_SERIALID;
                                  objLoad.P_SERIALID = obj.P_SERIALID;



                                  //  p_NGAYGIO              DATE,
                                  //p_CLOCK_METER DATE,
                                  //p_NGAY                 DATE,



                                  // To copy a folder's contents to a new location:
                                  // Create a new target folder, if necessary.
                                  if (!System.IO.Directory.Exists(targetPath))
                                  {
                                      System.IO.Directory.CreateDirectory(targetPath);
                                  }



                                  analyze3GHUUHONG_LoadProfile(value_item, ref objLoad, ref obj, ref objHIS, ref objCUR, ref ketqua);

                                  if (ketqua == "")
                                      result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_LOADPRO, ref objLoad, ref error);
                                  else
                                  {
                                      //System.IO.File.Copy(filename, destFile);
                                      tenfile = obj.P_MA_DIEMDO + "_" + System.DateTime.Now.Date.ToString("ddMMyy") + "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString();
                                      // tenfile = filename.Substring(sourceDir.Length + 1);

                                      File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, tenfile + ".res"), true);
                                      System.IO.File.Delete(filename);
                                  }
                                  if (result < 0)
                                  {
                                      //System.IO.File.Copy(filename, destFile);
                                      tenfile = obj.P_MA_DIEMDO + "_" + System.DateTime.Now.Date.ToString("ddMMyy") + "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString();
                                      // tenfile = filename.Substring(sourceDir.Length + 1);

                                      File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, tenfile + ".res"), true);
                                      System.IO.File.Delete(filename);
                                  }

                                  ketqua = "";
                                  error = "";
                                  analyze3GHUUHONG_TSVH(value_item, ref objLoad, ref obj, ref objHIS, ref objCUR, ref ketqua);
                                  if (ketqua == "")
                                      result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_CUR, ref objCUR, ref error);


                                  analyze3GHUUHONG_CSC(value_item, ref objLoad, ref obj, ref objHIS, ref objCUR, ref ketqua);
                                  objHIS.P_STARTBILL = new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, 1);
                                  objHIS.P_ENDBILL = new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, 1).AddMonths(1).AddDays(-1);
                                  if (ketqua == "")
                                      result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_HIS, ref objHIS, ref error);



                                  if (result != -1)
                                  {
                                      General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                                      System.IO.File.Delete(filename);
                                  }


                                  return;

                              default:
                                  break;
                          }
                          string err = "";
                          result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_CHISO_1PHA, ref obj, ref err);
                          if (err != "")
                          {
                              MessageBox.Show(err);
                          }
                          if (result != -1)
                          {
                              General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                              if (ketqua != "")
                              {
                                  System.IO.File.Copy(filename, destFile, true);
                              }

                              System.IO.File.Delete(filename);
                          }
                      }
                  }

              }
                       */
                        #endregion

                    }
                }

            }
            catch (Exception ex)
            {

                General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex.Message), Color.Red);

            }
        }



        public void do_analyze_3phaHuuhong_block(string path, RichTextBox rtbLog)
        {
            try
            {


                string sourcePath = @"D:\EVNHES_3G";
                System.IO.Directory.CreateDirectory(@"D:\EVNHES_3G\RES_ERR\" + DateTime.Now.ToString("ddMMyy"));
                // System.IO.Directory.Delete(@"D:\EVNHES_3G\RES_ERR\" + DateTime.Now.AddDays(-1).ToString("ddMMyy") );
                string targetPath = @"D:\EVNHES_3G\RES_ERR\" + DateTime.Now.ToString("ddMMyy");
                string oradb = General.connectString;

                string ketqua = "";
                dynamic obj = new ExpandoObject();
                dynamic objHIS = new ExpandoObject();
                dynamic objCUR = new ExpandoObject();

                dynamic objLoad = new ExpandoObject();
                dynamic objCosfi = new ExpandoObject();
                int icount = 0;
                string[] value_item = new string[100];
                var DyObjectsList = new List<ExpandoObject>();

                List<DataBLock> DulieuObject = new List<DataBLock>();
                string tenfile = "";
                string[] files = System.IO.Directory.GetFiles(path, "*.res");
                if (files.Count() == 0)
                    return;
                int result = -1;
                foreach (string filename in files)
                {
                    string strNgaygio = "";
                    string sourceFile = System.IO.Path.Combine(sourcePath, filename);
                    string destFile = System.IO.Path.Combine(targetPath);

                    int _startNum;
                    string values = System.IO.File.ReadAllText(filename);
                    char[] delimiters = new char[] { '\r', '\n' };

                    string strTempdata = "";
                    value_item[icount] = values.ToString().Trim();
                    icount++;

                    tenfile = Path.GetFileName(filename);
                    General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                    File.Copy(Path.Combine(sourcePath, filename), Path.Combine(targetPath, tenfile), true);
                    System.IO.File.Delete(filename);
                    if (icount == 100)
                        break;
                }

                for (int i = 0; i < value_item.Length; i++)
                {
                    if (value_item[i] == null)
                        break;
                    byte[] array = StringToByteArray(value_item[i].ToString());
                    DataBLock databloctemp = new DataBLock();

                    if (array == null)
                        return;
                    // Message Event co mã hóa  : 0c; 08;1B --> 
                    // Sự kiện Mất điện số lần mất điện thời gian mất điện (Thời gian bắt đầu ; Thời gian kết thúc)
                    // Công tơ bị thay đổi thời gian trả về 10 lần gần nhất 
                    // Pin yếu 
                    // Mất cân bằng pha
                    // Ngược thứ pha
                    // Ngược dòng pha A
                    // Ngược dòng pha B 
                    // Ngược dòng pha C
                    // Mất pha A
                    // Mất pha B
                    // Mất pha C
                    // Thay đổi pass công tơ 
                    // Số lần lập trình 

                    // Kiem tra auto report 
                    if (array[12] == 0x10 && array[16] == 0x08 && array[17] == 0x0C)
                    {
                        if (Processing.CheckFrame(value_item[i].ToString()) == true)
                        {
                            value_item[i] = FrameProcessing.Processing.StringDataDecryption(value_item[i].ToString(), "613E3367F67B8E24F09117B445191C92");
                        }
                        // Trong trường hợp này thì chuyển luồng 
                        if (value_item[i].ToString().Length > 100)
                        {
                            databloctemp = new DataBLock();

                            analyze3GHUUHONG_AutoReport(value_item[i].ToString(), ref databloctemp, ref ketqua);
                            if (databloctemp.NGAYGIO != null)
                                DulieuObject.Add(databloctemp);

                            //insert_Imiscumvals(conn, objCUR);
                            //insert_Loadprofile(conn, objCUR);
                            //insert_Imishistory(conn, objCUR);
                            //insert_ImisPmax(conn, objCUR);

                        }
                        else
                        {

                        }

                    }

                    if (array[12] == 0x0C && array[16] == 0x80 && array[17] == 0x18)
                    {
                        // Trong trường hợp này thì chuyển luồng 
                        if (value_item[i].ToString().Length > 100)
                        {
                            databloctemp = new DataBLock();
                            // Kiểm tra dạng autoreport ;
                            byte[] CheckGT = StringToByteArray(value_item[i].ToString());
                            string CheckFW = CheckGT[29].ToString("X2");
                            if (CheckFW == "FF")
                            {
                                analyze3GHUUHONG_AutoReport_newprotocol(value_item[i].ToString(), ref databloctemp, ref ketqua);
                                if (databloctemp.NGAYGIO != null)
                                    DulieuObject.Add(databloctemp);
                            }
                            else
                            {
                                analyze3GHUUHONG_AutoReport(value_item[i].ToString(), ref databloctemp, ref ketqua);
                                if (databloctemp.NGAYGIO != null)
                                    DulieuObject.Add(databloctemp);
                            }


                            //insert_Imiscumvals(conn, objCUR);
                            //insert_Loadprofile(conn, objCUR);
                            //insert_Imishistory(conn, objCUR);
                            //insert_ImisPmax(conn, objCUR);

                        }
                        else
                        {

                        }

                    }
                }// End for phan tich 

                string LstSerial = "";
                for (int iicount = DulieuObject.Count - 1; iicount >= 0; iicount--)
                {

                    LstSerial = LstSerial + DulieuObject[iicount].SERIALID + ",";

                }
                OracleConnection conn = new OracleConnection(oradb);  // C#
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                DataTable _listmadiemdo = GetListMeter_BySerial(LstSerial, conn);
                int sequenid = 0;
                for (int iicount = DulieuObject.Count - 1; iicount >= 0; iicount--)
                {

                    var res = from row in _listmadiemdo.AsEnumerable()
                              where row.Field<string>("SERIALID") == DulieuObject[iicount].SERIALID
                              select row.Field<string>("MA_DIEMDO");
                    //  int.TryParse((System.DateTime.Now.Day.ToString() + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Year.ToString()), out sequenid);
                    int.TryParse((System.DateTime.Now.Year.ToString() + System.DateTime.Now.Millisecond.ToString()), out sequenid);
                    if (res.Count() > 0)
                    {
                        DulieuObject[iicount].MA_DIEMDO = res.ToList()[0].ToString();
                        DulieuObject[iicount].SeqID = sequenid;
                    }
                    else
                    {
                        DulieuObject[iicount].MA_DIEMDO = "PD00CHUAMA";
                        DulieuObject[iicount].SeqID = sequenid;

                    }

                }
                OracleConnection conn2 = new OracleConnection(oradb);

                insert_Imiscumvals_bulk(conn2, DulieuObject);
                insert_Imisload_bulk(conn2, DulieuObject);
                insert_ImisVcumbill_bulk(conn2, DulieuObject);
                insert_ImisPmax_bulk(conn2, DulieuObject);
                insert_Imiscumvals_current_bulk(conn2, DulieuObject);
                conn.Close();

                // OracleTransaction txn = conn.BeginTransaction(IsolationLevel.ReadCommitted);
                //if (conn != null && conn.State == ConnectionState.Closed)
                //    conn.Open();
                //for (int iicount = DulieuObject.Count-1; iicount >= 0;iicount--)
                //{

                //    insert_Imiscumvals(conn, DulieuObject[iicount]);

                //}
                //txn.Commit();
                //conn.Close();

            }
            catch (Exception ex)
            {

                General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex.Message), Color.Red);

            }
        }



        private bool clearDataLoadPro(ref dynamic objLoadDCU)
        {
            objLoadDCU.LOAD_P_IMPORTKW = null;
            objLoadDCU.LOAD_P_EXPORTKW = null;
            objLoadDCU.LOAD_P_C1 = null;
            objLoadDCU.LOAD_P_C2 = null;
            objLoadDCU.LOAD_P_CONG_CS = null;
            objLoadDCU.LOAD_P_TRU_CS = null;
            objLoadDCU.LOAD_Q_TRU_CS = null;
            objLoadDCU.LOAD_Q_CONG_CS = null;
            return true;
        }
        private bool insert_Loadprofile(OracleConnection conn, dynamic objLoadDCU)
        {
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                OracleCommand cmd = new OracleCommand("insert_imis_loadpro_serialid", conn);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("@p_SERIALID", OracleDbType.NVarchar2).Value = objLoadDCU.SERIALID;
                cmd.Parameters.Add("@p_NGAY", OracleDbType.Date).Value = objLoadDCU.P_NGAY;
                cmd.Parameters.Add("@p_IMPORTKW", OracleDbType.Double).Value = objLoadDCU.LOAD_P_IMPORTKW;
                cmd.Parameters.Add("@p_EXPORTKW", OracleDbType.Double).Value = objLoadDCU.LOAD_P_EXPORTKW;
                cmd.Parameters.Add("@p_C1", OracleDbType.Double).Value = objLoadDCU.LOAD_P_C1;
                cmd.Parameters.Add("@p_C2", OracleDbType.Double).Value = objLoadDCU.LOAD_P_C2;
                cmd.Parameters.Add("@p_INPUT1", OracleDbType.Double).Value = objLoadDCU.LOAD_P_CONG_CS;
                cmd.Parameters.Add("@p_INPUT2", OracleDbType.Double).Value = objLoadDCU.LOAD_P_TRU_CS;
                cmd.Parameters.Add("@p_VCG_CSO", OracleDbType.Double).Value = objLoadDCU.LOAD_Q_TRU_CS;
                cmd.Parameters.Add("@p_VCN_CSO", OracleDbType.Double).Value = objLoadDCU.LOAD_Q_CONG_CS;
                cmd.ExecuteNonQuery();
                clearDataLoadPro(ref objLoadDCU);


                return true;

            }
            catch (Exception ex)
            {
                return false;
                //    MessageBox.Show(ex.Message);
            }
        }


        private bool clearDataImisCumval(ref dynamic objLoadDCU)
        {
            objLoadDCU.P_IMPORTKWH = null;
            objLoadDCU.P_EXPORTKWH = null;

            objLoadDCU.P_C1 = null;
            objLoadDCU.P_IMPBT = null;
            objLoadDCU.P_IMPCD = null;
            objLoadDCU.P_IMPTD = null;
            objLoadDCU.P_V_A = null;
            objLoadDCU.P_V_B = null;
            objLoadDCU.P_V_C = null;

            objLoadDCU.P_A_A = null;
            objLoadDCU.P_A_B = null;
            objLoadDCU.P_A_C = null;
            return true;

        }
        //GET_LISTMETER_1PHASE

        public static DataTable GetListMeter_BySerial(string serialDcu, OracleConnection conn)
        {
            DataTable dt = new DataTable();
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                using (conn)
                {
                    OracleCommand cmd = conn.CreateCommand();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "PKG_DCU_MDMS.GET_LISTMETER_BY_SERIALMETER";
                    cmd.Parameters.Add("P_LSTSERIAL", OracleDbType.NVarchar2, serialDcu, ParameterDirection.Input);
                    cmd.Parameters.Add("cur", OracleDbType.RefCursor, ParameterDirection.Output);
                    OracleDataAdapter da = new OracleDataAdapter(cmd);
                    da.Fill(dt);
                }

            }
            catch (Exception ex)
            {

            }

            return dt;
        }


        public static DataTable GetListMeter_1phareSerial(string serialDcu, OracleConnection conn)
        {
            DataTable dt = new DataTable();
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                using (conn)
                {
                    OracleCommand cmd = conn.CreateCommand();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "PKG_DCU_MDMS.GET_LISTMETER_1PHASE";
                    cmd.Parameters.Add("P_LSTSERIAL", OracleDbType.NVarchar2, serialDcu, ParameterDirection.Input);
                    cmd.Parameters.Add("cur", OracleDbType.RefCursor, ParameterDirection.Output);
                    OracleDataAdapter da = new OracleDataAdapter(cmd);
                    da.Fill(dt);
                }
                conn.Close();
            }
            catch (Exception ex)
            {

            }

            return dt;
        }

        private bool insert_Imiscumvals_bulk(OracleConnection conn, List<DataBLock> objCUR)
        {
            bool returnValue = false;
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                string query = @"insert into EVNHES.IMIS_VCUMVALS(ID,MA_DIEMDO,SERIALID,NGAYGIO,NGAY,CLOCK_METER,IMPORTKWH,EXPORTKWH,C1,IMPBT,IMPCD,IMPTD,V_A,V_B,V_C,A_A,A_B,A_C,TU_IN,TI_IN,DATEID,MONTHID,YEARID,HOURID,MINUTEID) 
       values (:ID,:MA_DIEMDO,:SERIALID,:NGAYGIO,:NGAY,:CLOCK_METER,:IMPORTKWH,:EXPORTKWH,:C1,:IMPBT,:IMPCD,:IMPTD,:V_A,:V_B,:V_C,:A_A,:A_B,:A_C,:TU_IN,:TI_IN,:DATEID,:MONTHID,:YEARID,:HOURID,:MINUTEID)";


                using (var command = conn.CreateCommand())
                {
                    command.CommandText = query;
                    command.CommandType = CommandType.Text;
                    command.BindByName = true;
                    // In order to use ArrayBinding, the ArrayBindCount property
                    // of OracleCommand object must be set to the number of records to be inserted
                    command.ArrayBindCount = objCUR.Count();
                    command.Parameters.Add(":ID", OracleDbType.Varchar2, objCUR.Select(c => c.SeqID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MA_DIEMDO", OracleDbType.Varchar2, objCUR.Select(c => c.MA_DIEMDO).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":SERIALID", OracleDbType.Varchar2, objCUR.Select(c => c.SERIALID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":NGAYGIO", OracleDbType.Date, objCUR.Select(c => c.NGAYGIO).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":NGAY", OracleDbType.Date, objCUR.Select(c => c.P_NGAY).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":CLOCK_METER", OracleDbType.Date, objCUR.Select(c => c.P_CLOCK_METER).ToArray(), ParameterDirection.Input);

                    command.Parameters.Add(":IMPORTKWH", OracleDbType.Double, objCUR.Select(c => c.P_IMPORTKWH).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":EXPORTKWH", OracleDbType.Double, objCUR.Select(c => c.P_EXPORTKWH).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":C1", OracleDbType.Double, objCUR.Select(c => c.P_C1).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":IMPBT", OracleDbType.Double, objCUR.Select(c => c.P_IMPBT).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":IMPCD", OracleDbType.Double, objCUR.Select(c => c.P_IMPCD).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":IMPTD", OracleDbType.Double, objCUR.Select(c => c.P_IMPTD).ToArray(), ParameterDirection.Input);

                    command.Parameters.Add(":V_A", OracleDbType.Double, objCUR.Select(c => c.P_V_A).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":V_B", OracleDbType.Double, objCUR.Select(c => c.P_V_B).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":V_C", OracleDbType.Double, objCUR.Select(c => c.P_V_C).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":A_A", OracleDbType.Double, objCUR.Select(c => c.P_A_A).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":A_B", OracleDbType.Double, objCUR.Select(c => c.P_A_B).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":A_C", OracleDbType.Double, objCUR.Select(c => c.P_A_C).ToArray(), ParameterDirection.Input);
                    //TU_IN,:TI_IN,:DATEID,:MONTHID,:YEARID,:HOURID,:MINUTEID)";
                    command.Parameters.Add(":TU_IN", OracleDbType.Int32, objCUR.Select(c => c.TU_TU / c.TU_MAU).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":TI_IN", OracleDbType.Int32, objCUR.Select(c => c.TI_TU / c.TI_MAU).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MONTHID", OracleDbType.Int32, objCUR.Select(c => c.MONTHID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":YEARID", OracleDbType.Int32, objCUR.Select(c => c.YEARID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":DATEID", OracleDbType.Int32, objCUR.Select(c => c.DATEID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":HOURID", OracleDbType.Int32, objCUR.Select(c => c.HOURID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MINUTEID", OracleDbType.Int32, objCUR.Select(c => c.MINUTEID).ToArray(), ParameterDirection.Input);
                    int result = command.ExecuteNonQuery();
                    if (result == objCUR.Count)
                        returnValue = true;
                }

            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                // conn.Close();
                return false;
            }

            //  conn.Close();
            return true;

        }

        private bool insert_Imiscumvals_current_bulk(OracleConnection conn, List<DataBLock> objCUR)
        {
            bool returnValue = false;
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                string query = @"insert into EVNHES.IMIS_VCUMVALS_CURRENT(ID,MA_DIEMDO,SERIALID,NGAYGIO,NGAY,CLOCK_METER,IMPORTKWH,EXPORTKWH,C1,IMPBT,IMPCD,IMPTD,V_A,V_B,V_C,A_A,A_B,A_C,TU_IN,TI_IN,DATEID,MONTHID,YEARID,HOURID,MINUTEID) 
       values (:ID,:MA_DIEMDO,:SERIALID,:NGAYGIO,:NGAY,:CLOCK_METER,:IMPORTKWH,:EXPORTKWH,:C1,:IMPBT,:IMPCD,:IMPTD,:V_A,:V_B,:V_C,:A_A,:A_B,:A_C,:TU_IN,:TI_IN,:DATEID,:MONTHID,:YEARID,:HOURID,:MINUTEID)";

                string query3 = @"MERGE INTO IMIS_VCUMVALS_CURRENT d
     USING (SELECT :ID AS ID,
                    :MA_DIEMDO MA_DIEMDO,
                   :SERIALID SERIALID,
                   :IMPORTKWH AS IMPORTKWH,
                   :EXPORTKWH AS EXPORTKWH,
                   :C1 AS C1,
                   :IMPBT AS IMPBT,
                   :IMPCD AS IMPCD,
                   :IMPTD AS IMPTD,
                   :V_A AS V_A,
                   :V_B AS V_B,
                   :V_C AS V_C,
                   :A_A AS A_A,
                   :A_B AS A_B,
                   :A_C AS A_C,
                   :TU_IN AS TU_IN,
                   :TI_IN AS TI_IN,
                  :DATEID AS DATEID,
                  :MONTHID AS MONTHID,
                   :YEARID AS YEARID,
                   :HOURID AS HOURID,
                   :MINUTEID AS MINUTEID,
                   :NGAY AS NGAY
              FROM DUAL) s
        ON (    d.MA_DIEMDO = s.MA_DIEMDO
            AND d.SERIALID = s.SERIALID
            AND d.DATEID = s.DATEID
            AND d.YEARID = s.YEARID
            AND d.MONTHID = s.MONTHID
            AND d.HOURID = s.HOURID
            AND d.MINUTEID = s.MINUTEID)
WHEN MATCHED
THEN
   UPDATE SET d.IMPORTKWH = s.IMPORTKWH,
              d.EXPORTKWH = s.EXPORTKWH,
              d.C1 = s.C1,
              d.IMPBT = s.IMPBT,
              D.IMPCD = s.IMPCD,
              d.V_A = s.V_A,
              d.V_B = s.V_B,
              d.V_C = s.V_C,
              d.A_A = s.A_A,
              d.A_B = s.A_B,
              d.A_C = s.A_C
              d.NGAY = S.NGAY
            
WHEN NOT MATCHED
THEN
   INSERT     (ID,
               MA_DIEMDO,
               SERIALID,
               NGAYGIO,
               NGAY,
               CLOCK_METER,
               IMPORTKWH,
               EXPORTKWH,
               C1,
               IMPBT,
               IMPCD,
               IMPTD,
               V_A,
               V_B,
               V_C,
               A_A,
               A_B,
               A_C,
               TU_IN,
               TI_IN,
               DATEID,
               MONTHID,
               YEARID,
               HOURID,
               MINUTEID
                             )
       VALUES (s.ID,
               s.MA_DIEMDO,
               s.SERIALID,
               :NGAYGIO,
               :NGAY,
               :CLOCK_METER,
               :IMPORTKWH ,
               :EXPORTKWH ,
               :C1,
               :IMPBT,
               :IMPCD ,
               :IMPTD ,
               :V_A ,
               :V_B ,
               :V_C ,
               :A_A ,
               :A_B ,
               :A_C ,
               :TU_IN ,
               :TI_IN ,
               :DATEID ,
               :MONTHID ,
               :YEARID ,
               :HOURID ,
               :MINUTEID );";

                using (var command = conn.CreateCommand())
                {
                    command.CommandText = query;
                    command.CommandType = CommandType.Text;
                    command.BindByName = true;
                    // In order to use ArrayBinding, the ArrayBindCount property
                    // of OracleCommand object must be set to the number of records to be inserted
                    command.ArrayBindCount = objCUR.Count();
                    command.Parameters.Add(":ID", OracleDbType.Varchar2, objCUR.Select(c => c.SeqID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MA_DIEMDO", OracleDbType.Varchar2, objCUR.Select(c => c.MA_DIEMDO).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":SERIALID", OracleDbType.Varchar2, objCUR.Select(c => c.SERIALID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":NGAYGIO", OracleDbType.Date, objCUR.Select(c => c.NGAYGIO).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":NGAY", OracleDbType.Date, objCUR.Select(c => c.P_NGAY).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":CLOCK_METER", OracleDbType.Date, objCUR.Select(c => c.P_CLOCK_METER).ToArray(), ParameterDirection.Input);

                    command.Parameters.Add(":IMPORTKWH", OracleDbType.Double, objCUR.Select(c => c.P_IMPORTKWH).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":EXPORTKWH", OracleDbType.Double, objCUR.Select(c => c.P_EXPORTKWH).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":C1", OracleDbType.Double, objCUR.Select(c => c.P_C1).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":IMPBT", OracleDbType.Double, objCUR.Select(c => c.P_IMPBT).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":IMPCD", OracleDbType.Double, objCUR.Select(c => c.P_IMPCD).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":IMPTD", OracleDbType.Double, objCUR.Select(c => c.P_IMPTD).ToArray(), ParameterDirection.Input);

                    command.Parameters.Add(":V_A", OracleDbType.Double, objCUR.Select(c => c.P_V_A).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":V_B", OracleDbType.Double, objCUR.Select(c => c.P_V_B).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":V_C", OracleDbType.Double, objCUR.Select(c => c.P_V_C).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":A_A", OracleDbType.Double, objCUR.Select(c => c.P_A_A).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":A_B", OracleDbType.Double, objCUR.Select(c => c.P_A_B).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":A_C", OracleDbType.Double, objCUR.Select(c => c.P_A_C).ToArray(), ParameterDirection.Input);
                    //TU_IN,:TI_IN,:DATEID,:MONTHID,:YEARID,:HOURID,:MINUTEID)";
                    command.Parameters.Add(":TU_IN", OracleDbType.Int32, objCUR.Select(c => c.TU_TU / c.TU_MAU).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":TI_IN", OracleDbType.Int32, objCUR.Select(c => c.TI_TU / c.TI_MAU).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MONTHID", OracleDbType.Int32, objCUR.Select(c => c.MONTHID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":YEARID", OracleDbType.Int32, objCUR.Select(c => c.YEARID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":DATEID", OracleDbType.Int32, objCUR.Select(c => c.DATEID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":HOURID", OracleDbType.Int32, objCUR.Select(c => c.HOURID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MINUTEID", OracleDbType.Int32, objCUR.Select(c => c.MINUTEID).ToArray(), ParameterDirection.Input);
                    int result = command.ExecuteNonQuery();
                    if (result == objCUR.Count)
                        returnValue = true;
                }

            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                // conn.Close();
                return false;
            }

            //  conn.Close();
            return true;

        }




        private bool insert_ImisPmax_bulk(OracleConnection conn, List<DataBLock> objCUR)
        {
            bool returnValue = false;
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                string query = @"insert into EVNHES.IMIS_MAXDEMAND(MA_DIEMDO,SERIALID,IMPORTKW,IMPORTKW_TIME,TU,TI,PDATE,PMONTH,PYEAR) 
                                                         values (:MA_DIEMDO,:SERIALID,:IMPORTKWH,:IMPORTKWH_TIME,:TU_IN,:TI_IN,:DATEID,:MONTHID,:YEARID)";


                using (var command = conn.CreateCommand())
                {
                    command.CommandText = query;
                    command.CommandType = CommandType.Text;
                    command.BindByName = true;
                    // In order to use ArrayBinding, the ArrayBindCount property
                    // of OracleCommand object must be set to the number of records to be inserted
                    command.ArrayBindCount = objCUR.Count();
                    //  command.Parameters.Add(":ID", OracleDbType.Varchar2, objCUR.Select(c => c.SeqID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MA_DIEMDO", OracleDbType.Varchar2, objCUR.Select(c => c.MA_DIEMDO).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":SERIALID", OracleDbType.Varchar2, objCUR.Select(c => c.SERIALID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":IMPORTKWH", OracleDbType.Double, objCUR.Select(c => c.PMAX_IMPORTKWH).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":IMPORTKWH_TIME", OracleDbType.Date, objCUR.Select(c => c.PMAX_DATE).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":TU_IN", OracleDbType.Int32, objCUR.Select(c => c.TU_TU / c.TU_MAU).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":TI_IN", OracleDbType.Int32, objCUR.Select(c => c.TI_TU / c.TI_MAU).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MONTHID", OracleDbType.Int32, objCUR.Select(c => c.MONTHID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":YEARID", OracleDbType.Int32, objCUR.Select(c => c.YEARID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":DATEID", OracleDbType.Int32, objCUR.Select(c => c.DATEID).ToArray(), ParameterDirection.Input);
                    int result = command.ExecuteNonQuery();
                    if (result == objCUR.Count)
                        returnValue = true;
                }

            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                // conn.Close();
                return false;
            }

            //  conn.Close();
            return true;

        }

        private bool insert_ImisVcumbill_bulk(OracleConnection conn, List<DataBLock> objCUR)
        {
            bool returnValue = false;
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                string query = @"insert into EVNHES.IMIS_VCUMBILL(MA_DIEMDO,SERIALID,STARTBILL,ENDBILL,IMPORTKWH,EXPORTKWH,C1,C2,IMPBT,IMPCD,IMPTD,HSN,TU,TI) 
                                                              values (:MA_DIEMDO,:SERIALID,:STARTDATE,:ENDDATE,:IMPORTKWH,:EXPORTKWH,:C1,:C2,:IMPBT,:IMPCD,:IMPTD,:HSN,:TU,:TI)";


                using (var command = conn.CreateCommand())
                {
                    command.CommandText = query;
                    command.CommandType = CommandType.Text;
                    command.BindByName = true;
                    // In order to use ArrayBinding, the ArrayBindCount property
                    // of OracleCommand object must be set to the number of records to be inserted
                    command.ArrayBindCount = objCUR.Count();
                    //  command.Parameters.Add(":ID", OracleDbType.Varchar2, objCUR.Select(c => c.SeqID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":MA_DIEMDO", OracleDbType.Varchar2, objCUR.Select(c => c.MA_DIEMDO).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":SERIALID", OracleDbType.Varchar2, objCUR.Select(c => c.SERIALID).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":STARTDATE", OracleDbType.Date, objCUR.Select(c => c.START_BILL).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":ENDDATE", OracleDbType.Date, objCUR.Select(c => c.END_BILL).ToArray(), ParameterDirection.Input);
                    //command.Parameters.Add(":NGAY", OracleDbType.Date, objCUR.Select(c => c.P_NGAY).ToArray(), ParameterDirection.Input);
                    //  command.Parameters.Add(":CLOCK_METER", OracleDbType.Date, objCUR.Select(c => c.P_CLOCK_METER).ToArray(), ParameterDirection.Input);

                    command.Parameters.Add(":IMPORTKWH", OracleDbType.Double, objCUR.Select(c => c.CHOT_HCGCS).ToArray(), ParameterDirection.Input); // Công suất giao 
                    command.Parameters.Add(":EXPORTKWH", OracleDbType.Double, objCUR.Select(c => c.CHOT_HCNCS).ToArray(), ParameterDirection.Input); // Công suất nhận 
                    command.Parameters.Add(":C1", OracleDbType.Double, objCUR.Select(c => c.CHOT_VCGCS).ToArray(), ParameterDirection.Input); // Vô công giao Q
                    command.Parameters.Add(":C2", OracleDbType.Double, objCUR.Select(c => c.CHOT_VCNCS).ToArray(), ParameterDirection.Input); // Vô công nhận Q

                    command.Parameters.Add(":IMPBT", OracleDbType.Double, objCUR.Select(c => c.CHOT_HCGCSB1).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":IMPCD", OracleDbType.Double, objCUR.Select(c => c.CHOT_HCGCSB2).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":IMPTD", OracleDbType.Double, objCUR.Select(c => c.CHOT_HCGCSB3).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":HSN", OracleDbType.Double, objCUR.Select(c => c.HSN).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":TU", OracleDbType.Int32, objCUR.Select(c => c.TU_TU / c.TU_MAU).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":TI", OracleDbType.Int32, objCUR.Select(c => c.TI_TU / c.TI_MAU).ToArray(), ParameterDirection.Input);
                    int result = command.ExecuteNonQuery();
                    if (result == objCUR.Count)
                        returnValue = true;
                }

            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                // conn.Close();
                return false;
            }

            //  conn.Close();
            return true;

        }


        /*
         * -	Công suất giao : IMPORTKW   
-	Công suất nhận: EXPORTKW   
-	INPUT1 : chỉ số giao
-	INPUT2     : Chỉ số nhận
-	INPUT3     : Sản lượng giao
-	INPUT4     : Sản lượng nhận
-	VCG_CSO    chỉ số vô công giao,
-	VCN_CSO    chỉ số vô công nhận,
-	VCG_SL     Sản lượng vô công giao,
-	VCN_SL     sản lượng vô công nhận

Trong đó
-	TH1: nếu công tơ chỉ có kênh công suất
o	Lưu giá trị vào : IMPORTKW   và EXPORTKW   
o	INPUT3     = IMPORTKW/2
o	INPUT4     = EXPORTKW   /2
-	TH2: Nếu công tơ chỉ có kênh chỉ số
o	Lưu giá trị vào : INPUT1, INPUT2     
o	INPUT3     , INPUT4     : được tính toán = chu kỳ sau – chu kỳ trước
o	IMPORTKW   = INPUT1*2
o	EXPORTKW    = INPUT2     *2
-	TH3: Nếu công tơ có cả kênh công suất và kênh chỉ số
o	Lưu giá trị vào : IMPORTKW   và EXPORTKW   
o	Lưu giá trị vào : INPUT1, INPUT2     
o	INPUT3     , INPUT4     : được tính toán = chu kỳ sau – chu kỳ trước

         * */


        private bool insert_Imisload_bulk(OracleConnection conn, List<DataBLock> objCUR)
        {
            bool returnValue = false;
            try
            {
                if (objCUR.Count > 0)
                {
                    if (conn != null && conn.State == ConnectionState.Closed)
                        conn.Open();

                    string query = @"insert into EVNHES.IMIS_LOADPRO(ID,MA_DIEMDO,SERIALID,STARTDATE,ENDDATE,NGAY,IMPORTKW,EXPORTKW,C1,C2,INPUT1,INPUT2,HCG_CS,HCN_CS,TU,TI,DATEID,MONTHID,YEARID,HOURID,MINUTEID) 
                                                              values (:ID,:MA_DIEMDO,:SERIALID,:STARTDATE,:ENDDATE,:NGAY,:IMPORTKWH,:EXPORTKWH,:C1,:C2,:HCG_CS,:HCN_CS,:INPUT1,:INPUT2,:TU,:TI,:DATEID,:MONTHID,:YEARID,:HOURID,:MINUTEID)";


                    using (var command = conn.CreateCommand())
                    {
                        command.CommandText = query;
                        command.CommandType = CommandType.Text;
                        command.BindByName = true;
                        // In order to use ArrayBinding, the ArrayBindCount property
                        // of OracleCommand object must be set to the number of records to be inserted
                        command.ArrayBindCount = objCUR.Count();
                        command.Parameters.Add(":ID", OracleDbType.Varchar2, objCUR.Select(c => c.SeqID).ToArray(), ParameterDirection.Input);
                        command.Parameters.Add(":MA_DIEMDO", OracleDbType.Varchar2, objCUR.Select(c => c.MA_DIEMDO).ToArray(), ParameterDirection.Input);
                        command.Parameters.Add(":SERIALID", OracleDbType.Varchar2, objCUR.Select(c => c.SERIALID).ToArray(), ParameterDirection.Input);
                        command.Parameters.Add(":STARTDATE", OracleDbType.Date, objCUR.Select(c => c.START_DATE).ToArray(), ParameterDirection.Input);
                        command.Parameters.Add(":ENDDATE", OracleDbType.Date, objCUR.Select(c => c.END_DATE).ToArray(), ParameterDirection.Input);
                        command.Parameters.Add(":NGAY", OracleDbType.Date, objCUR.Select(c => c.P_NGAY).ToArray(), ParameterDirection.Input);
                        //  command.Parameters.Add(":CLOCK_METER", OracleDbType.Date, objCUR.Select(c => c.P_CLOCK_METER).ToArray(), ParameterDirection.Input);

                        command.Parameters.Add(":IMPORTKWH", OracleDbType.Double, objCUR.Select(c => c.LOAD_P_IMPORTKW).ToArray(), ParameterDirection.Input); // Công suất giao 
                        command.Parameters.Add(":EXPORTKWH", OracleDbType.Double, objCUR.Select(c => c.LOAD_P_EXPORTKW).ToArray(), ParameterDirection.Input); // Công suất nhận 
                        command.Parameters.Add(":C1", OracleDbType.Double, objCUR.Select(c => c.LOAD_P_C1).ToArray(), ParameterDirection.Input); // Vô công giao Q
                        command.Parameters.Add(":C2", OracleDbType.Double, objCUR.Select(c => c.LOAD_P_C2).ToArray(), ParameterDirection.Input); // Vô công nhận Q

                        command.Parameters.Add(":INPUT1", OracleDbType.Double, objCUR.Select(c => c.LOAD_P_CONG_CS).ToArray(), ParameterDirection.Input); // Chỉ số giao 
                        command.Parameters.Add(":INPUT2", OracleDbType.Double, objCUR.Select(c => c.LOAD_P_TRU_CS).ToArray(), ParameterDirection.Input); // Chỉ số nhân 
                        command.Parameters.Add(":HCG_CS", OracleDbType.Double, objCUR.Select(c => c.LOAD_P_CONG_CS).ToArray(), ParameterDirection.Input); // Chỉ số giao 
                        command.Parameters.Add(":HCN_CS", OracleDbType.Double, objCUR.Select(c => c.LOAD_P_TRU_CS).ToArray(), ParameterDirection.Input); // Chỉ số nhân 

                        //TU_IN,:TI_IN,:DATEID,:MONTHID,:YEARID,:HOURID,:MINUTEID)";
                        command.Parameters.Add(":TU", OracleDbType.Int32, objCUR.Select(c => c.TU_TU / c.TU_MAU).ToArray(), ParameterDirection.Input);
                        command.Parameters.Add(":TI", OracleDbType.Int32, objCUR.Select(c => c.TI_TU / c.TI_MAU).ToArray(), ParameterDirection.Input);
                        command.Parameters.Add(":DATEID", OracleDbType.Int32, objCUR.Select(c => c.DATEID).ToArray(), ParameterDirection.Input);
                        command.Parameters.Add(":MONTHID", OracleDbType.Int32, objCUR.Select(c => c.MONTHID).ToArray(), ParameterDirection.Input);
                        command.Parameters.Add(":YEARID", OracleDbType.Int32, objCUR.Select(c => c.YEARID).ToArray(), ParameterDirection.Input);
                        command.Parameters.Add(":HOURID", OracleDbType.Int32, objCUR.Select(c => c.HOURID).ToArray(), ParameterDirection.Input);
                        command.Parameters.Add(":MINUTEID", OracleDbType.Int32, objCUR.Select(c => c.MINUTEID).ToArray(), ParameterDirection.Input);
                        int result = command.ExecuteNonQuery();
                        if (result == objCUR.Count)
                            returnValue = true;
                    }
                }



            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                // conn.Close();
                return false;
            }

            //  conn.Close();
            return true;

        }

        private bool insert_Imiscumvals(OracleConnection conn, DataBLock objCUR)
        {
            try
            {

                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                OracleCommand cmd = new OracleCommand("insert_imis_vcumvals_full", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@p_IMEI", OracleDbType.NVarchar2).Value = objCUR.SERIALID;
                cmd.Parameters.Add("@p_NGAYGIO", OracleDbType.Date).Value = objCUR.NGAYGIO;
                cmd.Parameters.Add("@p_CLOCK_METER", OracleDbType.Date).Value = System.DateTime.Now;
                cmd.Parameters.Add("@p_NGAY", OracleDbType.Date).Value = objCUR.NGAYGIO;

                cmd.Parameters.Add("@p_IMPORTKWH", OracleDbType.Double).Value = objCUR.P_IMPORTKWH;
                cmd.Parameters.Add("@p_EXPORTKWH", OracleDbType.Double).Value = objCUR.P_EXPORTKWH;
                cmd.Parameters.Add("@p_C1", OracleDbType.Double).Value = objCUR.P_C1;
                cmd.Parameters.Add("@p_IMPBT", OracleDbType.Double).Value = objCUR.P_IMPBT;
                cmd.Parameters.Add("@p_IMPCD", OracleDbType.Double).Value = objCUR.P_IMPCD;
                cmd.Parameters.Add("@p_IMPTD", OracleDbType.Double).Value = objCUR.P_IMPTD;
                cmd.Parameters.Add("@p_V_A", OracleDbType.Double).Value = objCUR.P_V_A;
                cmd.Parameters.Add("@p_V_B", OracleDbType.Double).Value = objCUR.P_V_B;
                cmd.Parameters.Add("@p_V_C", OracleDbType.Double).Value = objCUR.P_V_C;
                cmd.Parameters.Add("@p_A_A", OracleDbType.Double).Value = objCUR.P_A_A;
                cmd.Parameters.Add("@p_A_B", OracleDbType.Double).Value = objCUR.P_A_B;
                cmd.Parameters.Add("@p_A_C", OracleDbType.Double).Value = objCUR.P_A_C;
                cmd.Parameters.Add("@p_TU_IN", OracleDbType.NVarchar2).Value = objCUR.TU_TU / objCUR.TU_MAU;
                cmd.Parameters.Add("@p_TI_IN", OracleDbType.NVarchar2).Value = objCUR.TI_TU / objCUR.TI_MAU;
                cmd.ExecuteNonQuery();
                //  clearDataImisCumval(ref objCUR);

            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                // conn.Close();
                return false;
            }

            //  conn.Close();
            return true;

        }

        private bool insert_Imiscumvals(OracleConnection conn, dynamic objCUR)
        {
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                OracleCommand cmd = new OracleCommand("insert_imis_vcumvals_full", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@p_IMEI", OracleDbType.NVarchar2).Value = objCUR.SERIALID;
                cmd.Parameters.Add("@p_NGAYGIO", OracleDbType.Date).Value = objCUR.NGAYGIO;
                cmd.Parameters.Add("@p_CLOCK_METER", OracleDbType.Date).Value = System.DateTime.Now;
                cmd.Parameters.Add("@p_NGAY", OracleDbType.Date).Value = objCUR.NGAYGIO;

                cmd.Parameters.Add("@p_IMPORTKWH", OracleDbType.Double).Value = objCUR.P_IMPORTKWH;
                cmd.Parameters.Add("@p_EXPORTKWH", OracleDbType.Double).Value = objCUR.P_EXPORTKWH;
                cmd.Parameters.Add("@p_C1", OracleDbType.Double).Value = objCUR.P_C1;
                cmd.Parameters.Add("@p_IMPBT", OracleDbType.Double).Value = objCUR.P_IMPBT;
                cmd.Parameters.Add("@p_IMPCD", OracleDbType.Double).Value = objCUR.P_IMPCD;
                cmd.Parameters.Add("@p_IMPTD", OracleDbType.Double).Value = objCUR.P_IMPTD;
                cmd.Parameters.Add("@p_V_A", OracleDbType.Double).Value = objCUR.P_V_A;
                cmd.Parameters.Add("@p_V_B", OracleDbType.Double).Value = objCUR.P_V_B;
                cmd.Parameters.Add("@p_V_C", OracleDbType.Double).Value = objCUR.P_V_C;
                cmd.Parameters.Add("@p_A_A", OracleDbType.Double).Value = objCUR.P_A_A;
                cmd.Parameters.Add("@p_A_B", OracleDbType.Double).Value = objCUR.P_A_B;
                cmd.Parameters.Add("@p_A_C", OracleDbType.Double).Value = objCUR.P_A_C;
                cmd.Parameters.Add("@p_TU_IN", OracleDbType.NVarchar2).Value = objCUR.TU_TU / objCUR.TU_MAU;
                cmd.Parameters.Add("@p_TI_IN", OracleDbType.NVarchar2).Value = objCUR.TI_TU / objCUR.TI_MAU;
                cmd.ExecuteNonQuery();
                //  clearDataImisCumval(ref objCUR);

            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                conn.Close();
                return false;
            }

            conn.Close();
            return true;

        }

        private bool insert_Imishistory(OracleConnection conn, dynamic objCUR)
        {
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                OracleCommand cmd = new OracleCommand("insert_imis_vcumbill_tnv", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@p_SERIALID", OracleDbType.NVarchar2).Value = objCUR.SERIALID;
                cmd.Parameters.Add("@p_STARTBILL", OracleDbType.Date).Value = new DateTime(System.DateTime.Today.Year, System.DateTime.Today.Month - 1, 1);
                cmd.Parameters.Add("@p_ENDBILL", OracleDbType.Date).Value = new DateTime(System.DateTime.Today.Year, System.DateTime.Today.Month, 1);
                cmd.Parameters.Add("@p_IMPORTKWH", OracleDbType.Double).Value = objCUR.CHOT_HCGCS;
                cmd.Parameters.Add("@p_EXPORTKWH", OracleDbType.Double).Value = objCUR.CHOT_HCNCS;
                cmd.Parameters.Add("@p_C1", OracleDbType.Double).Value = objCUR.CHOT_VCGCS;
                cmd.Parameters.Add("@p_C2", OracleDbType.Double).Value = objCUR.CHOT_VCNCS; // Q TỔng 
                cmd.Parameters.Add("@p_IMPBT", OracleDbType.Double).Value = objCUR.CHOT_HCGCSB1;
                cmd.Parameters.Add("@p_IMPCD", OracleDbType.Double).Value = objCUR.CHOT_HCGCSB2;
                cmd.Parameters.Add("@p_IMPTD", OracleDbType.Double).Value = objCUR.CHOT_HCGCSB3;
                cmd.Parameters.Add("@p_EXPBT", OracleDbType.Double).Value = objCUR.CHOT_HCNCSB1;
                cmd.Parameters.Add("@p_EXPCD", OracleDbType.Double).Value = objCUR.CHOT_HCNCSB2;
                cmd.Parameters.Add("@p_EXPTD", OracleDbType.Double).Value = objCUR.CHOT_HCNCSB3;
                cmd.Parameters.Add("@p_HSN", OracleDbType.Double).Value = objCUR.HSN;
                cmd.Parameters.Add("@p_TU", OracleDbType.Double).Value = objCUR.TU_TU / objCUR.TU_MAU;
                cmd.Parameters.Add("@p_TI", OracleDbType.Double).Value = objCUR.TI_TU / objCUR.TI_MAU;
                cmd.ExecuteNonQuery();
                clearDataImisHISl(ref objCUR);
            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                conn.Close();
                return false;
            }

            conn.Close();
            return true;

        }

        private bool clearDataImisHISl(ref dynamic objLoadDCU)
        {
            objLoadDCU.CHOT_HCGCS = null;
            objLoadDCU.CHOT_HCNCS = null;
            objLoadDCU.CHOT_VCGCS = null;
            objLoadDCU.CHOT_VCNCS = null;

            objLoadDCU.CHOT_HCGCSB1 = null;
            objLoadDCU.CHOT_HCGCSB2 = null;
            objLoadDCU.CHOT_HCGCSB3 = null;
            objLoadDCU.CHOT_HCNCSB1 = null;
            objLoadDCU.CHOT_HCNCSB2 = null;
            objLoadDCU.CHOT_HCNCSB3 = null;
            objLoadDCU.HSN = null;

            return true;

        }


        private bool insert_ImisPmax(OracleConnection conn, dynamic objCUR)
        {
            try
            {
                if (conn != null && conn.State == ConnectionState.Closed)
                    conn.Open();

                OracleCommand cmd = new OracleCommand("INSERT_IMIS_MAXDEMAND_SERIAL", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@P_IMPORTKW", OracleDbType.Double).Value = objCUR.PMAX_IMPORTKWH;
                cmd.Parameters.Add("@P_IMPORTKW_TIME", OracleDbType.Date).Value = objCUR.PMAX_DATE;
                cmd.Parameters.Add("@p_SERIALID", OracleDbType.NVarchar2).Value = objCUR.SERIALID;

                cmd.ExecuteNonQuery();
            }

            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                conn.Close();
                return false;
            }

            conn.Close();
            return true;

        }


        private void analyzeHUUHONG(string[] value_item, ref dynamic obj)
        {
            for (int i = 2; i < value_item.Length; i++)
            {
                string[] row_item = value_item[i].Split(new char[] { '\t' });
                try
                {
                    if (row_item[0] == "TSVH")
                    {
                        obj.P_SERIALID = row_item[1];
                        obj.P_DATETIME = DateTime.Now;
                        obj.P_IMPORTKWH = double.Parse(row_item[2]);
                    }
                }
                catch (Exception ex)
                {
                }
            }
        }
        public string Analyze_Diennang_msgCTO(ref List<DataRepeater> DataRepeater, ref List<DataPQmax> DataPQmax, ref List<DataChiso> DataChiso, ref List<DataChiso> DataChisoGio, ref List<DataLoadprofile> ListLoadprofile, ref List<DataEvent> ListdataEvent, string data, ref string Datestring, ref DateTime iDate, ref double idiennang, ref double idiennang1, ref double idiennang2, ref double idiennang3, ref string serialcongto, ref double Pmax, ref double QMax, ref double Dienap, ref double Dongdien, ref double Solancanhbao, ref string strStatus)
        {


            // 68D200D20068C48400A60C000D760000041B563562121600 110917 2000120917040032741000EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE8616

            try
            {


                //    data = data.Substring(5, data.Length-5);
                strStatus = "error";
                string result = "";
                // Kiem tra bo qua message dang ky cua DCU 
                string checkDCU = data[24].ToString() + data[25].ToString();
                serialcongto = "";
                if (checkDCU == "11")
                    return null;
                // Check type of message 

                string checktypedata = data[32].ToString() + data[33].ToString() + data[34].ToString() + data[35].ToString();
                switch (checktypedata)
                {
                    // 201C repeater 

                    case "201C":// repeater
                        serialcongto = data.Substring(46, 2) + data.Substring(44, 2) + data.Substring(42, 2) + data.Substring(40, 2) + data.Substring(38, 2) + data.Substring(36, 2);
                        iDate = new DateTime(int.Parse("20" + data.Substring(52, 2)),
                                                                                  int.Parse(data.Substring(60, 2)),
                                                                                  int.Parse(data.Substring(58, 2)),
                                                                                  int.Parse(data.Substring(56, 2)),
                                                                                  int.Parse(data.Substring(54, 2)),
                                                                                  int.Parse("0"));
                        DataRepeater dataRepet = new DataRepeater();
                        dataRepet.chiso = Convert.ToDouble(idiennang);

                        dataRepet.Loaidulieu = "201C";
                        dataRepet.chiso = 0;
                        dataRepet.Thoigian = iDate;
                        dataRepet.strSerial = serialcongto;
                        dataRepet.monthid = iDate.Month;
                        dataRepet.yearid = iDate.Year;
                        dataRepet.dateid = iDate.Day;

                        DataRepeater.Add(dataRepet);

                        strStatus = "ok";
                        return result;

                    case "041B":// Dien nang 
                        serialcongto = data.Substring(46, 2) + data.Substring(44, 2) + data.Substring(42, 2) + data.Substring(40, 2) + data.Substring(38, 2) + data.Substring(36, 2);
                        iDate = new DateTime(int.Parse("20" + data.Substring(52, 2)),
                                                          int.Parse(data.Substring(60, 2)),
                                                          int.Parse(data.Substring(58, 2)),
                                                          int.Parse(data.Substring(56, 2)),
                                                          int.Parse(data.Substring(54, 2)),
                                                          int.Parse("0"));

                        string tempdiennang = data.Substring(66, 10);

                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        idiennang = double.Parse(tempdiennang) / 10000;
                        Datestring = iDate.ToString("dd/MM/yyyy HH:mm:ss");
                        if (data.Substring(76, 2) == "EE")
                        {
                            idiennang1 = 0;
                            idiennang2 = 0;
                            idiennang3 = 0;
                        }
                        else
                        {

                            idiennang1 = double.Parse(clsConvert.ReveserString(data.Substring(76, 10))) / 10000;
                            idiennang2 = double.Parse(clsConvert.ReveserString(data.Substring(86, 10))) / 10000;
                            idiennang3 = double.Parse(clsConvert.ReveserString(data.Substring(96, 10))) / 10000;
                        }

                        //  result += string.Format("{0}\t{1}\n", clsAnalyze.CLOCK, clock.ToString("s"));

                        DataChiso datacs = new DataChiso();
                        datacs.chiso = Convert.ToDouble(idiennang);
                        datacs.chiso1 = Convert.ToDouble(idiennang1);
                        datacs.chiso2 = Convert.ToDouble(idiennang2);
                        datacs.chiso3 = Convert.ToDouble(idiennang3);
                        datacs.Loaidulieu = "041B";
                        datacs.Thoigian = iDate;
                        datacs.monthid = iDate.Month;
                        datacs.yearid = iDate.Year;
                        datacs.dateid = iDate.Day;
                        datacs.strSerial = serialcongto;
                        DataChiso.Add(datacs);

                        strStatus = "ok";
                        return result;

                    case "081B": //Điện năng vô công

                        serialcongto = data.Substring(46, 2) + data.Substring(44, 2) + data.Substring(42, 2) + data.Substring(40, 2) + data.Substring(38, 2) + data.Substring(36, 2);
                        iDate = new DateTime(int.Parse("20" + data.Substring(52, 2)),
                                                          int.Parse(data.Substring(60, 2)),
                                                          int.Parse(data.Substring(58, 2)),
                                                          int.Parse(data.Substring(56, 2)),
                                                          int.Parse(data.Substring(54, 2)),
                                                          int.Parse("0"));

                        tempdiennang = data.Substring(66, 8);
                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        idiennang = double.Parse(tempdiennang) / 100;
                        Datestring = iDate.ToString("dd/MM/yyyy HH:mm:ss");

                        DataChiso datacsvc = new DataChiso();
                        datacsvc.Qtong = Convert.ToDouble(idiennang);
                        datacsvc.Loaidulieu = "081B";
                        datacsvc.Thoigian = iDate;
                        datacsvc.monthid = iDate.Month;
                        datacsvc.yearid = iDate.Year;
                        datacsvc.dateid = iDate.Day;

                        datacsvc.strSerial = serialcongto;
                        DataChiso.Add(datacsvc);

                        return result;

                    case "401C": // PMAX
                        serialcongto = data.Substring(46, 2) + data.Substring(44, 2) + data.Substring(42, 2) + data.Substring(40, 2) + data.Substring(38, 2) + data.Substring(36, 2);
                        iDate = new DateTime(int.Parse("20" + data.Substring(62, 2)),
                                                        int.Parse(data.Substring(60, 2)),
                                                        int.Parse(data.Substring(58, 2)),
                                                        int.Parse(data.Substring(56, 2)),
                                                        int.Parse(data.Substring(54, 2)),
                                                        int.Parse("0"));

                        tempdiennang = data.Substring(66, 6);
                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        Pmax = double.Parse(tempdiennang) / 10000;
                        Datestring = iDate.ToString("dd/MM/yyyy HH:mm:ss");
                        DataPQmax pqmax = new DataPQmax();
                        pqmax.pmax = Pmax;
                        pqmax.Loaidulieu = "401C";
                        pqmax.timepmax = iDate;
                        pqmax.strSerial = serialcongto;
                        pqmax.dateid = iDate.Day;
                        pqmax.monthid = iDate.Month;
                        pqmax.yearid = iDate.Year;
                        DataPQmax.Add(pqmax);
                        return result;

                    case "801C":
                        tempdiennang = data.Substring(66, 6);
                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        QMax = double.Parse(tempdiennang) / 10000;
                        return result;

                    case "4016":
                        tempdiennang = data.Substring(64, 4);
                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        Dienap = double.Parse(tempdiennang) / 100;
                        return result;

                    case "8016":
                        tempdiennang = data.Substring(64, 4);
                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        Dongdien = double.Parse(tempdiennang) / 10;
                        return result;
                    // Pmax -  "0417": 
                    // Xử lý với công tơ PLC 
                    case "0117":
                    case "0217":
                    case "0417":
                    case "0817":
                        Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                        return result;
                    case "0114":
                        Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                        return result;
                    case "0127":
                        Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                        return result;
                    // Sự kiện công tơ 
                    case "0100":
                        // Analyze_EVENT_PLC_DCU(ref ListdataEvent, data);
                        return result;
                    // Đọc dữ liệu công tơ theo giờ 
                    case "041A":
                        Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                        return result;

                    default:
                        return result;
                }
                //   Datestring = data[36].ToString() + data[37].ToString() + "/" + data[38].ToString() + data[39].ToString() + "/" + data[40].ToString() + data[41].ToString();
                MessageBox.Show("Ok");
            }
            catch (Exception ex)
            {
                return null;
            }

        }
        public string Analyze_Diennang_msgCTO(ref List<DataRepeater> DataRepeater, ref List<DataPQmax> DataPQmax, ref List<DataChiso> DataChiso, ref List<DataLoadprofile> ListLoadprofile, ref List<DataEvent> ListdataEvent, string data, ref string Datestring, ref DateTime iDate, ref double idiennang, ref double idiennang1, ref double idiennang2, ref double idiennang3, ref string serialcongto, ref double Pmax, ref double QMax, ref double Dienap, ref double Dongdien, ref double Solancanhbao, ref string strStatus)
        {


            // 68D200D20068C48400A60C000D760000041B563562121600 110917 2000120917040032741000EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE8616

            try
            {


                //    data = data.Substring(5, data.Length-5);
                strStatus = "error";
                string result = "";
                // Kiem tra bo qua message dang ky cua DCU 
                string checkDCU = data[24].ToString() + data[25].ToString();
                serialcongto = "";
                if (checkDCU == "11")
                    return null;
                // Check type of message 

                string checktypedata = data[32].ToString() + data[33].ToString() + data[34].ToString() + data[35].ToString();
                switch (checktypedata)
                {
                    // 201C repeater 

                    case "201C":// repeater
                        serialcongto = data.Substring(46, 2) + data.Substring(44, 2) + data.Substring(42, 2) + data.Substring(40, 2) + data.Substring(38, 2) + data.Substring(36, 2);
                        iDate = new DateTime(int.Parse("20" + data.Substring(52, 2)),
                                                                                  int.Parse(data.Substring(60, 2)),
                                                                                  int.Parse(data.Substring(58, 2)),
                                                                                  int.Parse(data.Substring(56, 2)),
                                                                                  int.Parse(data.Substring(54, 2)),
                                                                                  int.Parse("0"));
                        DataRepeater dataRepet = new DataRepeater();
                        dataRepet.chiso = Convert.ToDouble(idiennang);

                        dataRepet.Loaidulieu = "201C";
                        dataRepet.chiso = 0;
                        dataRepet.Thoigian = iDate;
                        dataRepet.strSerial = serialcongto;
                        dataRepet.monthid = iDate.Month;
                        dataRepet.yearid = iDate.Year;
                        dataRepet.dateid = iDate.Day;

                        DataRepeater.Add(dataRepet);

                        strStatus = "ok";
                        return result;

                    case "041B":// Dien nang 
                        serialcongto = data.Substring(46, 2) + data.Substring(44, 2) + data.Substring(42, 2) + data.Substring(40, 2) + data.Substring(38, 2) + data.Substring(36, 2);
                        iDate = new DateTime(int.Parse("20" + data.Substring(52, 2)),
                                                          int.Parse(data.Substring(60, 2)),
                                                          int.Parse(data.Substring(58, 2)),
                                                          int.Parse(data.Substring(56, 2)),
                                                          int.Parse(data.Substring(54, 2)),
                                                          int.Parse("0"));

                        string tempdiennang = data.Substring(66, 10);

                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        idiennang = double.Parse(tempdiennang) / 10000;
                        Datestring = iDate.ToString("dd/MM/yyyy HH:mm:ss");
                        if (data.Substring(76, 2) == "EE")
                        {
                            idiennang1 = 0;
                            idiennang2 = 0;
                            idiennang3 = 0;
                        }
                        else
                        {

                            idiennang1 = double.Parse(clsConvert.ReveserString(data.Substring(76, 10))) / 10000;
                            idiennang2 = double.Parse(clsConvert.ReveserString(data.Substring(86, 10))) / 10000;
                            idiennang3 = double.Parse(clsConvert.ReveserString(data.Substring(96, 10))) / 10000;
                        }

                        //  result += string.Format("{0}\t{1}\n", clsAnalyze.CLOCK, clock.ToString("s"));

                        DataChiso datacs = new DataChiso();
                        datacs.chiso = Convert.ToDouble(idiennang);
                        datacs.chiso1 = Convert.ToDouble(idiennang1);
                        datacs.chiso2 = Convert.ToDouble(idiennang2);
                        datacs.chiso3 = Convert.ToDouble(idiennang3);
                        datacs.Loaidulieu = "041B";
                        datacs.Thoigian = iDate;
                        datacs.monthid = iDate.Month;
                        datacs.yearid = iDate.Year;
                        datacs.dateid = iDate.Day;
                        datacs.strSerial = serialcongto;
                        DataChiso.Add(datacs);

                        strStatus = "ok";
                        return result;

                    case "081B": //Điện năng vô công

                        serialcongto = data.Substring(46, 2) + data.Substring(44, 2) + data.Substring(42, 2) + data.Substring(40, 2) + data.Substring(38, 2) + data.Substring(36, 2);
                        iDate = new DateTime(int.Parse("20" + data.Substring(52, 2)),
                                                          int.Parse(data.Substring(60, 2)),
                                                          int.Parse(data.Substring(58, 2)),
                                                          int.Parse(data.Substring(56, 2)),
                                                          int.Parse(data.Substring(54, 2)),
                                                          int.Parse("0"));

                        tempdiennang = data.Substring(66, 8);
                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        idiennang = double.Parse(tempdiennang) / 100;
                        Datestring = iDate.ToString("dd/MM/yyyy HH:mm:ss");

                        DataChiso datacsvc = new DataChiso();
                        datacsvc.Qtong = Convert.ToDouble(idiennang);
                        datacsvc.Loaidulieu = "081B";
                        datacsvc.Thoigian = iDate;
                        datacsvc.monthid = iDate.Month;
                        datacsvc.yearid = iDate.Year;
                        datacsvc.dateid = iDate.Day;

                        datacsvc.strSerial = serialcongto;
                        DataChiso.Add(datacsvc);

                        return result;

                    case "401C": // PMAX
                        serialcongto = data.Substring(46, 2) + data.Substring(44, 2) + data.Substring(42, 2) + data.Substring(40, 2) + data.Substring(38, 2) + data.Substring(36, 2);
                        iDate = new DateTime(int.Parse("20" + data.Substring(62, 2)),
                                                        int.Parse(data.Substring(60, 2)),
                                                        int.Parse(data.Substring(58, 2)),
                                                        int.Parse(data.Substring(56, 2)),
                                                        int.Parse(data.Substring(54, 2)),
                                                        int.Parse("0"));

                        tempdiennang = data.Substring(66, 6);
                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        Pmax = double.Parse(tempdiennang) / 10000;
                        Datestring = iDate.ToString("dd/MM/yyyy HH:mm:ss");
                        DataPQmax pqmax = new DataPQmax();
                        pqmax.pmax = Pmax;
                        pqmax.Loaidulieu = "401C";
                        pqmax.timepmax = iDate;
                        pqmax.strSerial = serialcongto;
                        pqmax.dateid = iDate.Day;
                        pqmax.monthid = iDate.Month;
                        pqmax.yearid = iDate.Year;
                        DataPQmax.Add(pqmax);
                        return result;

                    case "801C":
                        tempdiennang = data.Substring(66, 6);
                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        QMax = double.Parse(tempdiennang) / 10000;
                        return result;

                    case "4016":
                        tempdiennang = data.Substring(64, 4);
                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        Dienap = double.Parse(tempdiennang) / 100;
                        return result;

                    case "8016":
                        tempdiennang = data.Substring(64, 4);
                        tempdiennang = clsConvert.ReveserString(tempdiennang);
                        Dongdien = double.Parse(tempdiennang) / 10;
                        return result;
                    // Pmax -  "0417": 
                    // Xử lý với công tơ PLC 
                    case "0117":
                    case "0217":
                    case "0417":
                    case "0817":
                        Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                        return result;
                    case "0114":
                        Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                        return result;
                    case "0127":
                        Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                        return result;
                    // Sự kiện công tơ 
                    case "0100":
                        // Analyze_EVENT_PLC_DCU(ref ListdataEvent, data);
                        return result;

                    default:
                        return result;
                }
                //   Datestring = data[36].ToString() + data[37].ToString() + "/" + data[38].ToString() + data[39].ToString() + "/" + data[40].ToString() + data[41].ToString();
                MessageBox.Show("Ok");
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        public string Analyze_EVENT_PLC_DCU(ref List<DataEvent> ListdataEvent, string data)
        {
            DataEvent tempdtEvent = new DataEvent();
            string nocongto = "";
            nocongto = ReverseByte(data.Substring(62, 12));

            tempdtEvent.maevent = data.Substring(74, 2);
            tempdtEvent.Thoigian = new DateTime(int.Parse("20" + data.Substring(84, 2)),
                                            int.Parse(data.Substring(82, 2)),
                                            int.Parse(data.Substring(80, 2)),
                                            int.Parse(data.Substring(78, 2)),
                                            int.Parse(data.Substring(76, 2)),
                                            int.Parse(data.Substring(78, 2)));
            tempdtEvent.strSerial = nocongto;
            ListdataEvent.Add(tempdtEvent);

            data = data.Substring(0, 44) + data.Substring(88, data.Length - 88);
            if (data.Length > 88)
            {
                Analyze_EVENT_PLC_DCU(ref ListdataEvent, data);
                return "ok";
            }
            else
                return "ok";
        }

        public string Analyze_PLC_DCU(ref List<DataPQmax> DataPQmax, ref List<DataChiso> DataChiso, ref List<DataChiso> DataChisoGio, ref List<DataLoadprofile> ListLoadprofile, ref List<DataEvent> ListdataEvent, string data, ref string strStatus)
        {
            try
            {
                if (data.Count() <= 32)
                    return "ok";

                string tempdiennang = "";
                //    data = data.Substring(5, data.Length-5);
                strStatus = "error";
                DateTime iDate;
                string result = "";
                string soblock = "";
                // Kiem tra bo qua message dang ky cua DCU 
                string checkDCU = data[24].ToString() + data[25].ToString();
                int lengthframe = 0;
                string serialDCU = "";
                if (checkDCU == "11")
                    return null;
                // Check type of message 

                serialDCU = data.Substring(16, 2) + data.Substring(14, 2) + data.Substring(20, 2) + data.Substring(18, 2);

                string datablock = "";
                datablock = data.Substring(28, 78);
                string checktypedata = datablock[4].ToString() + datablock[5].ToString() + datablock[6].ToString() + datablock[7].ToString();
                //if (checktypedata == "0127" || checktypedata == "0227" || checktypedata == "0327" || checktypedata == "0427" || checktypedata == "0527" || checktypedata == "0627" || checktypedata == "0727" || checktypedata == "0827")
                //{
                //    datablock = data.Substring(28, 46);
                //}

                //iDate = new DateTime(int.Parse("20" + datablock.Substring(22, 2)),
                //                                  int.Parse(datablock.Substring(20, 2)),
                //                                  int.Parse(datablock.Substring(18, 2)),
                //                                  int.Parse(datablock.Substring(16, 2)),
                //                                  int.Parse(datablock.Substring(14, 2)),
                //                                  int.Parse("0"));

                string chiso = "";
                string chisobieu1 = "";
                string chisobieu2 = "";
                string chisobieu3 = "";
                DateTime Datechiso;
                string framestruct = "";
                string nocongto = "";
                switch (checktypedata)
                {
                    case "0100":
                        DataEvent tempdataevent = new DataEvent();
                        datablock = data.Substring(28, 88);
                        data = data.Substring(0, 28) + data.Substring(116, data.Length - 116);
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        tempdataevent.strSerial = nocongto;
                        return result;
                    case "0201":
                    case "0429":
                        DataTUTI tempdataTUTI = new DataTUTI();
                        datablock = data.Substring(28, 44);
                        data = data.Substring(0, 28) + data.Substring(72, data.Length - 72);
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        ReverseTUTI(datablock.Substring(36, 4), ref tempdataTUTI.CT_TU);
                        ReverseTUTI(datablock.Substring(40, 4), ref tempdataTUTI.CT_MAU);
                        ReverseTUTI(datablock.Substring(44, 4), ref tempdataTUTI.VT_TU);
                        ReverseTUTI(datablock.Substring(48, 4), ref tempdataTUTI.VT_MAU);
                        tempdataevent.strSerial = nocongto;
                        return result;
                    case "0114": // Điện năng hữu công chiều thuận 
                    case "0414":
                        datablock = data.Substring(28, 88);
                        data = data.Substring(0, 28) + data.Substring(116, data.Length - 116);


                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        if (nocongto == "000018030661")
                        {
                            nocongto = "000018030661";
                        }

                        chiso = datablock.Substring(38, 10);
                        if (chiso == "EEEEEEEEEE")
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);


                        chisobieu1 = datablock.Substring(48, 10);
                        chisobieu2 = datablock.Substring(58, 10);
                        chisobieu3 = datablock.Substring(68, 10);
                        Datechiso = new DateTime(int.Parse("20" + datablock.Substring(34, 2)),
                                                           int.Parse(datablock.Substring(32, 2)),
                                                           int.Parse(datablock.Substring(30, 2)),
                                                           int.Parse(datablock.Substring(28, 2)),
                                                           int.Parse(datablock.Substring(26, 2)),
                                                           int.Parse("0"));


                        DataChiso datacs = new DataChiso();
                        Reversechiso(chiso, ref datacs.chiso);
                        Reversechiso(chisobieu1, ref datacs.chiso1);
                        Reversechiso(chisobieu2, ref datacs.chiso2);
                        Reversechiso(chisobieu3, ref datacs.chiso3);
                        datacs.Loaidulieu = checktypedata;
                        datacs.Thoigian = Datechiso;
                        datacs.strSerial = nocongto;
                        DataChiso.Add(datacs);



                        if (data.Count() > 32)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return null;


                        }

                    case "4026":// Điện năng hữu công chiều thuận  đọc theo gio 
                    case "012A":// Điện năng hữu công chiều ngược  đọc theo gio 
                        datablock = data.Substring(28, 88);
                        data = data.Substring(0, 28) + data.Substring(116, data.Length - 116);


                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));

                        chiso = datablock.Substring(38, 10);
                        if (chiso == "EEEEEEEEEE")
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);


                        chisobieu1 = datablock.Substring(48, 10);
                        chisobieu2 = datablock.Substring(58, 10);
                        chisobieu3 = datablock.Substring(68, 10);
                        Datechiso = new DateTime(int.Parse("20" + datablock.Substring(34, 2)),
                                                           int.Parse(datablock.Substring(32, 2)),
                                                           int.Parse(datablock.Substring(30, 2)),
                                                           int.Parse(datablock.Substring(28, 2)),
                                                           int.Parse(datablock.Substring(26, 2)),
                                                           int.Parse("0"));


                        DataChiso datacsgio = new DataChiso();
                        Reversechiso(chiso, ref datacsgio.chiso);
                        Reversechiso(chisobieu1, ref datacsgio.chiso1);
                        Reversechiso(chisobieu2, ref datacsgio.chiso2);
                        Reversechiso(chisobieu3, ref datacsgio.chiso3);
                        datacsgio.Loaidulieu = checktypedata;
                        datacsgio.Thoigian = Datechiso;
                        datacsgio.strSerial = nocongto;
                        DataChisoGio.Add(datacsgio);

                        if (data.Count() > 32)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return null;


                        }
                    case "0214": // Điện năng vô công chiều thuận 
                    case "0814": // Điện năng vô công chiều ngược 
                        datablock = data.Substring(28, 78);
                        data = data.Substring(0, 28) + data.Substring(106, data.Length - 106);
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));

                        chiso = datablock.Substring(38, 8);
                        if (chiso == "EEEEEEEE")
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                        chisobieu1 = datablock.Substring(46, 8);
                        chisobieu2 = datablock.Substring(54, 8);
                        chisobieu3 = datablock.Substring(62, 8);
                        Datechiso = new DateTime(int.Parse("20" + datablock.Substring(34, 2)),
                                                     int.Parse(datablock.Substring(32, 2)),
                                                     int.Parse(datablock.Substring(30, 2)),
                                                     int.Parse(datablock.Substring(28, 2)),
                                                     int.Parse(datablock.Substring(26, 2)),
                                                     int.Parse("0"));

                        DataChiso datacsvc = new DataChiso();
                        ReverseQ(chiso, ref datacsvc.chiso);
                        ReverseQ(chisobieu1, ref datacsvc.chiso1);
                        ReverseQ(chisobieu2, ref datacsvc.chiso2);
                        ReverseQ(chisobieu3, ref datacsvc.chiso3);
                        datacsvc.Loaidulieu = checktypedata;
                        datacsvc.Thoigian = Datechiso;
                        datacsvc.strSerial = nocongto;
                        DataChiso.Add(datacsvc);

                        if (data.Count() > 0)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return "ok";
                        }
                        break;
                    case "200C": // Điện năng vô công chiều thuận đọc theo giờ  
                    case "800C": // Điện năng vô công chiều ngược đọc theo giờ  
                        datablock = data.Substring(28, 78);
                        data = data.Substring(0, 28) + data.Substring(106, data.Length - 106);
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));

                        chiso = datablock.Substring(38, 8);
                        if (chiso == "EEEEEEEE")
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                        chisobieu1 = datablock.Substring(46, 8);
                        chisobieu2 = datablock.Substring(54, 8);
                        chisobieu3 = datablock.Substring(62, 8);
                        Datechiso = new DateTime(int.Parse("20" + datablock.Substring(34, 2)),
                                                     int.Parse(datablock.Substring(32, 2)),
                                                     int.Parse(datablock.Substring(30, 2)),
                                                     int.Parse(datablock.Substring(28, 2)),
                                                     int.Parse(datablock.Substring(26, 2)),
                                                     int.Parse("0"));

                        DataChiso datacsvcgio = new DataChiso();
                        ReverseQ(chiso, ref datacsvcgio.chiso);
                        ReverseQ(chisobieu1, ref datacsvcgio.chiso1);
                        ReverseQ(chisobieu2, ref datacsvcgio.chiso2);
                        ReverseQ(chisobieu3, ref datacsvcgio.chiso3);
                        datacsvcgio.Loaidulieu = checktypedata;
                        datacsvcgio.Thoigian = Datechiso;
                        datacsvcgio.strSerial = nocongto;
                        DataChisoGio.Add(datacsvcgio);

                        if (data.Count() > 0)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return "ok";
                        }
                        break;
                    case "0117":
                    case "0217":
                    case "0417":
                    case "0817":
                        datablock = data.Substring(28, 108);
                        data = data.Substring(0, 28) + data.Substring(136, data.Length - 136);
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        string pmax = datablock.Substring(38, 6);
                        string pmaxbieu1 = datablock.Substring(52, 6);
                        string pmaxbieu2 = datablock.Substring(66, 6);
                        string pmaxbieu3 = datablock.Substring(80, 6);


                        DateTime DatePmax, DatePmax1, DatePmax2, DatePmax3;
                        if ((datablock.Substring(50, 2) != "00") && (pmax != "EEEEEE"))
                        {
                            DatePmax = new DateTime(DateTime.Now.Year,
                                                          int.Parse(datablock.Substring(50, 2)),
                                                          int.Parse(datablock.Substring(48, 2)),
                                                          int.Parse(datablock.Substring(46, 2)),
                                                          int.Parse(datablock.Substring(44, 2)),
                                                          int.Parse("0"));

                            DatePmax1 = new DateTime(DateTime.Now.Year,
                                                              int.Parse(datablock.Substring(64, 2)),
                                                              int.Parse(datablock.Substring(62, 2)),
                                                              int.Parse(datablock.Substring(60, 2)),
                                                              int.Parse(datablock.Substring(58, 2)),
                                                              int.Parse("0"));

                        }
                        else
                        {
                            DatePmax = new DateTime(1900, 1, 1);
                            DatePmax1 = new DateTime(1900, 1, 1);

                        }

                        if ((datablock.Substring(78, 2) != "00") && (datablock.Substring(78, 2) != "EE"))
                        {
                            DatePmax2 = new DateTime(DateTime.Now.Year,
                                                         int.Parse(datablock.Substring(78, 2)),
                                                         int.Parse(datablock.Substring(76, 2)),
                                                         int.Parse(datablock.Substring(74, 2)),
                                                         int.Parse(datablock.Substring(72, 2)),
                                                         int.Parse("0"));
                            DatePmax3 = new DateTime(DateTime.Now.Year,
                                                            int.Parse(datablock.Substring(92, 2)),
                                                            int.Parse(datablock.Substring(90, 2)),
                                                            int.Parse(datablock.Substring(88, 2)),
                                                            int.Parse(datablock.Substring(86, 2)),
                                                            int.Parse("0"));

                        }
                        else
                        {
                            DatePmax2 = new DateTime(1900, 1, 1);
                            DatePmax3 = new DateTime(1900, 1, 1);

                        }

                        DataPQmax datapqm = new DataPQmax();
                        ReversePmax(pmax, ref datapqm.pmax);
                        ReversePmax(pmaxbieu1, ref datapqm.pmax1);
                        ReversePmax(pmaxbieu2, ref datapqm.pmax2);
                        ReversePmax(pmaxbieu3, ref datapqm.pmax3);
                        datapqm.timepmax = DatePmax;
                        datapqm.timepmax1 = DatePmax1;
                        datapqm.timepmax2 = DatePmax2;
                        datapqm.timepmax3 = DatePmax3;
                        datapqm.Loaidulieu = checktypedata;
                        DataPQmax.Add(datapqm);

                        //  tempdiennang = data.Substring(66, 10);

                        if (data.Count() > 0)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return "ok";
                        }
                    case "0127":
                    case "0327":
                    case "0427":
                    case "0527":
                    case "0627":
                    case "0727":
                        soblock = datablock.Substring(32, 2);
                        lengthframe = int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6;
                        lengthframe = lengthframe + 34;
                        datablock = data.Substring(28, lengthframe);
                        data = data.Substring(0, 28) + data.Substring(lengthframe + 28, data.Length - (lengthframe + 28));
                        framestruct = "";
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        DateTime timechuky = new DateTime(int.Parse("20" + datablock.Substring(28, 2)),
                                                         int.Parse(datablock.Substring(26, 2)),
                                                         int.Parse(datablock.Substring(24, 2)),
                                                         int.Parse(datablock.Substring(22, 2)),
                                                         int.Parse(datablock.Substring(20, 2)),
                                                         int.Parse("0"));
                        if (datablock.Substring(datablock.Length - 6, 6) != "EEEEEE")
                        {

                            framestruct = datablock.Substring(datablock.Length - int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6, int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6);
                            for (int i = 0; i < Convert.ToInt32(soblock); i++)
                            {
                                DataLoadprofile temp2 = new DataLoadprofile();
                                temp2.P = Reverse(framestruct.Substring(0, 6));
                                temp2.Thoigian = timechuky.AddMinutes(i * 30);

                                temp2.Loaidulieu = checktypedata;
                                temp2.strSerial = nocongto;
                                ListLoadprofile.Add(temp2);
                            }
                        }

                        if (data.Count() > 0)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return "ok";
                        }

                    case "0227":
                    case "0827":
                        soblock = datablock.Substring(32, 2);
                        lengthframe = int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6;
                        lengthframe = lengthframe + 34;
                        datablock = data.Substring(28, lengthframe);
                        data = data.Substring(0, 28) + data.Substring(lengthframe + 28, data.Length - (lengthframe + 28));
                        framestruct = "";
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        timechuky = new DateTime(int.Parse("20" + datablock.Substring(28, 2)),
                                                        int.Parse(datablock.Substring(26, 2)),
                                                        int.Parse(datablock.Substring(24, 2)),
                                                        int.Parse(datablock.Substring(22, 2)),
                                                        int.Parse(datablock.Substring(20, 2)),
                                                        int.Parse("0"));
                        if (datablock.Substring(datablock.Length - 6, 6) != "EEEEEE")
                        {
                            framestruct = datablock.Substring(datablock.Length - int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6, int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6);
                            for (int i = 0; i < Convert.ToInt32(soblock); i++)
                            {
                                DataLoadprofile temp = new DataLoadprofile();
                                byte[] ba = StrToByteArray(datablock.Substring(0, 6));
                                ba[2] = (byte)(ba[2] & 0x7F);
                                Array.Reverse(ba);
                                string strtemp = ConvertByteArrayToString(ba);

                                double dvalue = Convert.ToDouble(strtemp);
                                //   string hexResult = result2.ToString("X");

                                temp.P = (double)dvalue / 10000;
                                temp.Thoigian = timechuky.AddMinutes(i * 30);
                                // temp.Thoigian = timechuky;
                                temp.Loaidulieu = checktypedata;
                                temp.strSerial = nocongto;
                                ListLoadprofile.Add(temp);
                            }


                        }





                        if (data.Count() > 0)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref DataChisoGio, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return "ok";
                        }
                        return "ok";
                    default:
                        return "ok";
                }
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        public string Analyze_PLC_DCU(ref List<DataPQmax> DataPQmax, ref List<DataChiso> DataChiso, ref List<DataLoadprofile> ListLoadprofile, ref List<DataEvent> ListdataEvent, string data, ref string strStatus)
        {
            try
            {
                if (data.Count() <= 32)
                    return "ok";

                string tempdiennang = "";
                //    data = data.Substring(5, data.Length-5);
                strStatus = "error";
                DateTime iDate;
                string result = "";
                string soblock = "";
                // Kiem tra bo qua message dang ky cua DCU 
                string checkDCU = data[24].ToString() + data[25].ToString();
                int lengthframe = 0;
                string serialDCU = "";
                if (checkDCU == "11")
                    return null;
                // Check type of message 

                serialDCU = data.Substring(16, 2) + data.Substring(14, 2) + data.Substring(20, 2) + data.Substring(18, 2);

                string datablock = "";
                datablock = data.Substring(28, 78);
                string checktypedata = datablock[4].ToString() + datablock[5].ToString() + datablock[6].ToString() + datablock[7].ToString();
                //if (checktypedata == "0127" || checktypedata == "0227" || checktypedata == "0327" || checktypedata == "0427" || checktypedata == "0527" || checktypedata == "0627" || checktypedata == "0727" || checktypedata == "0827")
                //{
                //    datablock = data.Substring(28, 46);
                //}

                //iDate = new DateTime(int.Parse("20" + datablock.Substring(22, 2)),
                //                                  int.Parse(datablock.Substring(20, 2)),
                //                                  int.Parse(datablock.Substring(18, 2)),
                //                                  int.Parse(datablock.Substring(16, 2)),
                //                                  int.Parse(datablock.Substring(14, 2)),
                //                                  int.Parse("0"));

                string chiso = "";
                string chisobieu1 = "";
                string chisobieu2 = "";
                string chisobieu3 = "";
                DateTime Datechiso;
                string framestruct = "";
                string nocongto = "";
                switch (checktypedata)
                {
                    case "0100":
                        DataEvent tempdataevent = new DataEvent();
                        datablock = data.Substring(28, 88);
                        data = data.Substring(0, 28) + data.Substring(116, data.Length - 116);
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        tempdataevent.strSerial = nocongto;
                        return result;
                    case "0201":
                    case "0429":
                        DataTUTI tempdataTUTI = new DataTUTI();
                        datablock = data.Substring(28, 44);
                        data = data.Substring(0, 28) + data.Substring(72, data.Length - 72);
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        ReverseTUTI(datablock.Substring(36, 4), ref tempdataTUTI.CT_TU);
                        ReverseTUTI(datablock.Substring(40, 4), ref tempdataTUTI.CT_MAU);
                        ReverseTUTI(datablock.Substring(44, 4), ref tempdataTUTI.VT_TU);
                        ReverseTUTI(datablock.Substring(48, 4), ref tempdataTUTI.VT_MAU);
                        tempdataevent.strSerial = nocongto;
                        return result;
                    case "0114": // Điện năng hữu công chiều thuận 
                    case "0414":
                        datablock = data.Substring(28, 88);
                        data = data.Substring(0, 28) + data.Substring(116, data.Length - 116);


                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        if (nocongto == "000018030661")
                        {
                            nocongto = "000018030661";
                        }

                        chiso = datablock.Substring(38, 10);
                        if (chiso == "EEEEEEEEEE")
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);


                        chisobieu1 = datablock.Substring(48, 10);
                        chisobieu2 = datablock.Substring(58, 10);
                        chisobieu3 = datablock.Substring(68, 10);
                        Datechiso = new DateTime(int.Parse("20" + datablock.Substring(34, 2)),
                                                           int.Parse(datablock.Substring(32, 2)),
                                                           int.Parse(datablock.Substring(30, 2)),
                                                           int.Parse(datablock.Substring(28, 2)),
                                                           int.Parse(datablock.Substring(26, 2)),
                                                           int.Parse("0"));


                        DataChiso datacs = new DataChiso();
                        Reversechiso(chiso, ref datacs.chiso);
                        Reversechiso(chisobieu1, ref datacs.chiso1);
                        Reversechiso(chisobieu2, ref datacs.chiso2);
                        Reversechiso(chisobieu3, ref datacs.chiso3);
                        datacs.Loaidulieu = checktypedata;
                        datacs.Thoigian = Datechiso;
                        datacs.strSerial = nocongto;
                        DataChiso.Add(datacs);



                        if (data.Count() > 32)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return null;


                        }


                    case "0214": // Điện năng vô công chiều thuận 
                    case "0814": // Điện năng vô công chiều ngược 
                        datablock = data.Substring(28, 78);
                        data = data.Substring(0, 28) + data.Substring(106, data.Length - 106);
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));

                        chiso = datablock.Substring(38, 8);
                        if (chiso == "EEEEEEEE")
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                        chisobieu1 = datablock.Substring(46, 8);
                        chisobieu2 = datablock.Substring(54, 8);
                        chisobieu3 = datablock.Substring(62, 8);
                        Datechiso = new DateTime(int.Parse("20" + datablock.Substring(34, 2)),
                                                     int.Parse(datablock.Substring(32, 2)),
                                                     int.Parse(datablock.Substring(30, 2)),
                                                     int.Parse(datablock.Substring(28, 2)),
                                                     int.Parse(datablock.Substring(26, 2)),
                                                     int.Parse("0"));

                        DataChiso datacsvc = new DataChiso();
                        ReverseQ(chiso, ref datacsvc.chiso);
                        ReverseQ(chisobieu1, ref datacsvc.chiso1);
                        ReverseQ(chisobieu2, ref datacsvc.chiso2);
                        ReverseQ(chisobieu3, ref datacsvc.chiso3);
                        datacsvc.Loaidulieu = checktypedata;
                        datacsvc.Thoigian = Datechiso;
                        datacsvc.strSerial = nocongto;
                        DataChiso.Add(datacsvc);

                        if (data.Count() > 0)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return "ok";
                        }
                        break;
                    case "0117":
                    case "0217":
                    case "0417":
                    case "0817":
                        datablock = data.Substring(28, 108);
                        data = data.Substring(0, 28) + data.Substring(136, data.Length - 136);
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        string pmax = datablock.Substring(38, 6);
                        string pmaxbieu1 = datablock.Substring(52, 6);
                        string pmaxbieu2 = datablock.Substring(66, 6);
                        string pmaxbieu3 = datablock.Substring(80, 6);


                        DateTime DatePmax, DatePmax1, DatePmax2, DatePmax3;
                        if ((datablock.Substring(50, 2) != "00") && (pmax != "EEEEEE"))
                        {
                            DatePmax = new DateTime(DateTime.Now.Year,
                                                          int.Parse(datablock.Substring(50, 2)),
                                                          int.Parse(datablock.Substring(48, 2)),
                                                          int.Parse(datablock.Substring(46, 2)),
                                                          int.Parse(datablock.Substring(44, 2)),
                                                          int.Parse("0"));

                            DatePmax1 = new DateTime(DateTime.Now.Year,
                                                              int.Parse(datablock.Substring(64, 2)),
                                                              int.Parse(datablock.Substring(62, 2)),
                                                              int.Parse(datablock.Substring(60, 2)),
                                                              int.Parse(datablock.Substring(58, 2)),
                                                              int.Parse("0"));

                        }
                        else
                        {
                            DatePmax = new DateTime(1900, 1, 1);
                            DatePmax1 = new DateTime(1900, 1, 1);

                        }

                        if ((datablock.Substring(78, 2) != "00") && (datablock.Substring(78, 2) != "EE"))
                        {
                            DatePmax2 = new DateTime(DateTime.Now.Year,
                                                         int.Parse(datablock.Substring(78, 2)),
                                                         int.Parse(datablock.Substring(76, 2)),
                                                         int.Parse(datablock.Substring(74, 2)),
                                                         int.Parse(datablock.Substring(72, 2)),
                                                         int.Parse("0"));
                            DatePmax3 = new DateTime(DateTime.Now.Year,
                                                            int.Parse(datablock.Substring(92, 2)),
                                                            int.Parse(datablock.Substring(90, 2)),
                                                            int.Parse(datablock.Substring(88, 2)),
                                                            int.Parse(datablock.Substring(86, 2)),
                                                            int.Parse("0"));

                        }
                        else
                        {
                            DatePmax2 = new DateTime(1900, 1, 1);
                            DatePmax3 = new DateTime(1900, 1, 1);

                        }

                        DataPQmax datapqm = new DataPQmax();
                        ReversePmax(pmax, ref datapqm.pmax);
                        ReversePmax(pmaxbieu1, ref datapqm.pmax1);
                        ReversePmax(pmaxbieu2, ref datapqm.pmax2);
                        ReversePmax(pmaxbieu3, ref datapqm.pmax3);
                        datapqm.timepmax = DatePmax;
                        datapqm.timepmax1 = DatePmax1;
                        datapqm.timepmax2 = DatePmax2;
                        datapqm.timepmax3 = DatePmax3;
                        datapqm.Loaidulieu = checktypedata;
                        DataPQmax.Add(datapqm);

                        //  tempdiennang = data.Substring(66, 10);

                        if (data.Count() > 0)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return "ok";
                        }
                    case "0127":
                    case "0327":
                    case "0427":
                    case "0527":
                    case "0627":
                    case "0727":
                        soblock = datablock.Substring(32, 2);
                        lengthframe = int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6;
                        lengthframe = lengthframe + 34;
                        datablock = data.Substring(28, lengthframe);
                        data = data.Substring(0, 28) + data.Substring(lengthframe + 28, data.Length - (lengthframe + 28));
                        framestruct = "";
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        DateTime timechuky = new DateTime(int.Parse("20" + datablock.Substring(28, 2)),
                                                         int.Parse(datablock.Substring(26, 2)),
                                                         int.Parse(datablock.Substring(24, 2)),
                                                         int.Parse(datablock.Substring(22, 2)),
                                                         int.Parse(datablock.Substring(20, 2)),
                                                         int.Parse("0"));
                        if (datablock.Substring(datablock.Length - 6, 6) != "EEEEEE")
                        {

                            framestruct = datablock.Substring(datablock.Length - int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6, int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6);
                            for (int i = 0; i < Convert.ToInt32(soblock); i++)
                            {
                                DataLoadprofile temp2 = new DataLoadprofile();
                                temp2.P = Reverse(framestruct.Substring(0, 6));
                                temp2.Thoigian = timechuky.AddMinutes(i * 30);

                                temp2.Loaidulieu = checktypedata;
                                temp2.strSerial = nocongto;
                                ListLoadprofile.Add(temp2);
                            }
                        }

                        if (data.Count() > 0)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return "ok";
                        }
                        return "ok";
                    case "0227":
                    case "0827":
                        soblock = datablock.Substring(32, 2);
                        lengthframe = int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6;
                        lengthframe = lengthframe + 34;
                        datablock = data.Substring(28, lengthframe);
                        data = data.Substring(0, 28) + data.Substring(lengthframe + 28, data.Length - (lengthframe + 28));
                        framestruct = "";
                        nocongto = "";
                        nocongto = ReverseByte(datablock.Substring(8, 12));
                        timechuky = new DateTime(int.Parse("20" + datablock.Substring(28, 2)),
                                                        int.Parse(datablock.Substring(26, 2)),
                                                        int.Parse(datablock.Substring(24, 2)),
                                                        int.Parse(datablock.Substring(22, 2)),
                                                        int.Parse(datablock.Substring(20, 2)),
                                                        int.Parse("0"));
                        if (datablock.Substring(datablock.Length - 6, 6) != "EEEEEE")
                        {
                            framestruct = datablock.Substring(datablock.Length - int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6, int.Parse(soblock, System.Globalization.NumberStyles.HexNumber) * 6);
                            for (int i = 0; i < Convert.ToInt32(soblock); i++)
                            {
                                DataLoadprofile temp = new DataLoadprofile();
                                byte[] ba = StrToByteArray(datablock.Substring(0, 6));
                                ba[2] = (byte)(ba[2] & 0x7F);
                                Array.Reverse(ba);
                                string strtemp = ConvertByteArrayToString(ba);

                                double dvalue = Convert.ToDouble(strtemp);
                                //   string hexResult = result2.ToString("X");

                                temp.P = (double)dvalue / 10000;
                                temp.Thoigian = timechuky.AddMinutes(i * 30);
                                // temp.Thoigian = timechuky;
                                temp.Loaidulieu = checktypedata;
                                temp.strSerial = nocongto;
                                ListLoadprofile.Add(temp);
                            }


                        }




                        if (data.Count() > 0)
                        {
                            Analyze_PLC_DCU(ref DataPQmax, ref DataChiso, ref ListLoadprofile, ref ListdataEvent, data, ref strStatus);
                            return "ok";
                        }
                        else
                        {
                            return "ok";
                        }
                        return "ok";
                    default:
                        return "ok";
                }
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        public double ConvertByteArrayToDouble(byte[] b)

        {
            double temp = 0;

            for (int i = 0; i < b.Length; i++)
            {
                temp = temp + (double)b[i];
            }
            return temp;
        }

        public string ConvertByteArrayToString(byte[] b)

        {
            string temp = "";

            for (int i = 0; i < b.Length; i++)
            {
                temp = temp + b[i].ToString("X2");
            }
            return temp;
        }

        public static byte[] StrToByteArray(string str)
        {
            Dictionary<string, byte> hexindex = new Dictionary<string, byte>();
            for (int i = 0; i <= 255; i++)
                hexindex.Add(i.ToString("X2"), (byte)i);

            List<byte> hexres = new List<byte>();
            for (int i = 0; i < str.Length; i += 2)
                hexres.Add(hexindex[str.Substring(i, 2)]);

            return hexres.ToArray();
        }

        private void ReversePmax(string inputdata, ref double outputdata)
        {
            try
            {
                if (inputdata == "EEEEEEEEEE")
                {
                    outputdata = 0;
                }
                else
                {
                    string serialDCU = inputdata.Substring(4, 2) + inputdata.Substring(2, 2) + inputdata.Substring(0, 2);
                    outputdata = double.Parse(serialDCU) / 10000;

                }
            }
            catch (Exception ex)
            {

            }
        }

        private void ReverseQ(string inputdata, ref double outputdata)
        {
            try
            {
                if (inputdata == "EEEEEEEEEE")
                {
                    outputdata = 0;
                }
                else
                {
                    string serialDCU = inputdata.Substring(6, 2) + inputdata.Substring(4, 2) + inputdata.Substring(2, 2) + inputdata.Substring(0, 2);
                    outputdata = double.Parse(serialDCU) / 100;

                }
            }
            catch (Exception ex)
            {

            }
        }

        private void Reversechiso(string inputdata, ref double outputdata)
        {
            try
            {
                if (inputdata == "EEEEEEEEEE")
                {
                    outputdata = 0;
                }
                else
                {
                    string serialDCU = inputdata.Substring(8, 2) + inputdata.Substring(6, 2) + inputdata.Substring(4, 2) + inputdata.Substring(2, 2) + inputdata.Substring(0, 2);
                    outputdata = double.Parse(serialDCU) / 10000;

                }
            }
            catch (Exception ex)
            {

            }
        }
        private void ReverseTUTI(string inputdata, ref double outputdata)
        {
            try
            {
                if (inputdata == "EEEE")
                {
                    outputdata = 0;
                }
                else
                {
                    string serialDCU = inputdata.Substring(2, 2) + inputdata.Substring(0, 2);
                    outputdata = double.Parse(serialDCU) / 10000;

                }
            }
            catch (Exception ex)
            {

            }
        }
        private double Reverse(string inputdata)
        {
            try
            {
                string serialDCU = inputdata.Substring(4, 2) + inputdata.Substring(2, 2) + inputdata.Substring(0, 2);
                return double.Parse(serialDCU) / 10000;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }




        private void analyzeHUUHONG_NEW(string[] value_item, ref List<DataRepeater> ListRepeter, ref List<DataPQmax> PQMax, ref List<DataChiso> ListChiso, ref List<DataLoadprofile> ListLoadprofile, ref List<DataEvent> ListdataEvent, ref dynamic obj)
        {
            string strCheckTypedata = "";
            string[] row_item = value_item[0].Split(new char[] { '\t' });
            string abc = "";
            try
            {
                // Kiểm tra loại giá trị dữ liệu
                //if (row_item[0].IndexOf(":") > 0)
                //    strCheckTypedata = row_item[0].Substring(0, row_item[0].IndexOf(":"));
                //else
                //    strCheckTypedata = "";

                string strdate = "";
                DateTime tempdatetime = new DateTime();
                double ddiennang = 0, diennangc1 = 0, diennangc2 = 0, diennangc3 = 0;
                string strSerial = "";
                string strStatust = "";
                double dQmax = 0;
                double dDienap = 0;
                double dDongdien = 0;
                double dSolancanhbao = 0;
                double dPmax = 0;

                for (int i = 0; i < value_item.Length; i++)
                {
                    ddiennang = 0;
                    diennangc1 = 0;
                    diennangc2 = 0;
                    diennangc3 = 0;
                    Analyze_Diennang_msgCTO(ref ListRepeter, ref PQMax, ref ListChiso, ref ListLoadprofile, ref ListdataEvent, value_item[i], ref strdate, ref tempdatetime, ref ddiennang, ref diennangc1, ref diennangc2, ref diennangc3, ref strSerial, ref dPmax, ref dQmax, ref dDienap, ref dDongdien, ref dSolancanhbao, ref strStatust);
                    abc = value_item[i];
                    if (abc is null)
                        break;

                }
                obj.P_SERIALID = strSerial;
                obj.P_IMPORTKWH = ddiennang;
                obj.P_EXPORTKWH = 0;
                obj.P_DATETIME = tempdatetime;


            }
            catch (Exception ex)
            {
            }

        }

        private void analyzeHUUHONG_NEW(string[] value_item, ref List<DataRepeater> ListRepeter, ref List<DataPQmax> PQMax, ref List<DataChiso> ListChiso, ref List<DataChiso> ListChisogio, ref List<DataLoadprofile> ListLoadprofile, ref List<DataEvent> ListdataEvent, ref dynamic obj)
        {
            string strCheckTypedata = "";
            string[] row_item = value_item[0].Split(new char[] { '\t' });
            string abc = "";
            try
            {
                // Kiểm tra loại giá trị dữ liệu
                //if (row_item[0].IndexOf(":") > 0)
                //    strCheckTypedata = row_item[0].Substring(0, row_item[0].IndexOf(":"));
                //else
                //    strCheckTypedata = "";

                string strdate = "";
                DateTime tempdatetime = new DateTime();
                double ddiennang = 0, diennangc1 = 0, diennangc2 = 0, diennangc3 = 0;
                string strSerial = "";
                string strStatust = "";
                double dQmax = 0;
                double dDienap = 0;
                double dDongdien = 0;
                double dSolancanhbao = 0;
                double dPmax = 0;

                for (int i = 0; i < value_item.Length; i++)
                {
                    ddiennang = 0;
                    diennangc1 = 0;
                    diennangc2 = 0;
                    diennangc3 = 0;
                    Analyze_Diennang_msgCTO(ref ListRepeter, ref PQMax, ref ListChiso, ref ListChisogio, ref ListLoadprofile, ref ListdataEvent, value_item[i], ref strdate, ref tempdatetime, ref ddiennang, ref diennangc1, ref diennangc2, ref diennangc3, ref strSerial, ref dPmax, ref dQmax, ref dDienap, ref dDongdien, ref dSolancanhbao, ref strStatust);
                    abc = value_item[i];
                    if (abc is null)
                        break;

                }
                obj.P_SERIALID = strSerial;
                obj.P_IMPORTKWH = ddiennang;
                obj.P_EXPORTKWH = 0;
                obj.P_DATETIME = tempdatetime;


            }
            catch (Exception ex)
            {
            }

        }

        public void analyzeIFC(string[] value_item, ref dynamic obj)
        {
            for (int i = 2; i < value_item.Length; i++)
            {
                string[] row_item = value_item[i].Split(new char[] { '\t' });
                try
                {
                    if (row_item[0] == "FTSVH")
                    {
                        obj.P_SERIALID = row_item[1];
                        obj.P_DATETIME = DateTime.ParseExact(row_item[2] + row_item[3], "yyMMddssmmHH",
                                       System.Globalization.CultureInfo.InvariantCulture);
                        obj.P_VOLPHASE = double.Parse(row_item[4]) / 10;
                        obj.P_CURPHASE = double.Parse(row_item[5]) / 1000;
                        obj.P_ACTIVE_POWER_PHASE = double.Parse(row_item[6]);
                        obj.P_COSPHI = double.Parse(row_item[7]) / 1000;
                        obj.P_IMPORTKWH = double.Parse(row_item[8]) / 100;
                    }
                }
                catch (Exception ex)
                {
                }
            }
        }

        public void analyzeVinasino(string[] value_item, string filename)
        {
            dynamic obj = new ExpandoObject();
            string error = "";
            int result = -1;
            for (int i = 2; i < value_item.Length; i++)
            {

                string[] row_item = value_item[i].Split(new char[] { '\t' });
                if (row_item[0] == "00000" || row_item[2] == "FFFF")
                {
                    return;
                }
                try
                {
                    obj.P_MA_DIEMDO = row_item[0];
                    obj.P_SERIALID = row_item[1];
                    obj.P_DATETIME = DateTime.ParseExact(value_item[1] + row_item[2], "yyMMddHHmm",
                                   System.Globalization.CultureInfo.InvariantCulture);

                    obj.P_VOLPHASE = double.Parse(ReverseByte(row_item[4].ToString())) / 10;
                    obj.P_CURPHASE = double.Parse(ReverseByte(row_item[5].ToString())) / 10;
                    obj.P_IMPORTKWH = double.Parse(ReverseByte(row_item[3].ToString())) / 100;
                    result = OracleDataAccess1.execute_NonQuery(ref error, General.connectString, PKG_Library.MDMSDCU_INSERT_CHISO_1PHA, ref obj, ref error);
                    if (error != "")
                    {
                        MessageBox.Show(error);
                    }
                }
                catch (Exception ex)
                {
                }
            }
            if (result != -1)
            {
                General.AppendRichTextBox(ref rtbLog, string.Format("{0}\t{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), filename), Color.Blue);
                System.IO.File.Delete(filename);
            }
        }

        public static string ReverseByte(string str)
        {
            string str2 = "";
            for (int i = 0; i < (str.Length / 2); i++)
            {
                str2 = str.Substring(i * 2, 2) + str2;
            }
            return str2;
        }
        public void analyzeEMEC(string[] value_item, ref dynamic obj)
        {
            for (int i = 2; i < value_item.Length; i++)
            {
                string[] row_item = value_item[i].Split(new char[] { '\t' });
                {
                    obj.P_SERIALID = row_item[1];
                    obj.P_DATETIME = Convert.ToDateTime(row_item[8]);

                    obj.P_ROUTING_PATH = row_item[3];
                    obj.P_RF_RSSI = int.Parse(row_item[4]);
                    obj.P_RF_LEVER = int.Parse(row_item[5]);
                    obj.P_RF_POWER_MAX = int.Parse(row_item[6]);
                    obj.P_READ_PATH_SUCCESS = row_item[7];
                    string[] item_para = row_item[2].Split(new char[] { '_' });
                    if (item_para.Count() > 3)
                    {
                        obj.P_IMPORTKWH = double.Parse(item_para[0].Replace(",", "."));
                        obj.P_EXPORTKWH = double.Parse(item_para[1].Replace(",", "."));
                        obj.P_VOLPHASE = double.Parse(item_para[3].Replace(",", "."));
                        obj.P_CURPHASE = double.Parse(item_para[2].Replace(",", "."));
                        obj.P_ACTIVE_POWER_PHASE = double.Parse(item_para[4].Replace(",", "."));
                    }
                    else
                    {
                        obj.P_IMPORTKWH = double.Parse(item_para[0].Replace(",", "."));
                    }

                }
            }
        }

        public string seprarateString(string[] value_item)
        {

            return value_item[0].Substring(value_item[0].IndexOf(":") + 1, value_item[0].Length - value_item[0].IndexOf(":") - 1);
        }



        #region Tungnv code 3G Huu hồng

        public string ToStringDate(string strinput)
        {
            byte[] arrsent = clsConvert.ToByteArray(strinput);
            for (int i = 0; i < arrsent.Length; i++)
            {
                arrsent[i] = (byte)(arrsent[i] - 0x33);
            }

            string strSs = clsConvert.ByteArrayToString(arrsent);
            strSs = strSs.Substring(6, 2) + "/" + strSs.Substring(8, 2) + "/20" + strSs.Substring(10, 2) + " " + strSs.Substring(4, 2) + ":" + strSs.Substring(2, 2) + ":" + strSs.Substring(0, 2);

            return strSs;
        }

        public string ToStringDate(string strinput, ref string date1, ref string date2)
        {
            byte[] arrsent = clsConvert.ToByteArray(strinput);
            for (int i = 0; i < arrsent.Length; i++)
            {
                arrsent[i] = (byte)(arrsent[i] - 0x33);
            }

            string strSs = clsConvert.ByteArrayToString(arrsent);
            date1 = strSs.Substring(6, 2) + "/" + strSs.Substring(8, 2) + "/20" + strSs.Substring(10, 2) + " " + strSs.Substring(4, 2) + ":" + strSs.Substring(2, 2);
            date2 = strSs.Substring(6, 2) + "/" + strSs.Substring(8, 2) + "/20" + strSs.Substring(10, 2) + " " + strSs.Substring(4, 2) + ":" + (int.Parse(strSs.Substring(2, 2)) + 30).ToString();
            strSs = strSs.Substring(6, 2) + "/" + strSs.Substring(8, 2) + "/20" + strSs.Substring(10, 2);

            return strSs;
        }

        public string ToStringDate(string strinput, decimal decValue)
        {
            byte[] arrsent = clsConvert.ToByteArray(strinput);
            for (int i = 0; i < arrsent.Length; i++)
            {
                arrsent[i] = (byte)(arrsent[i] - 0x33);
            }

            string strSs = clsConvert.ByteArrayToString(arrsent);
            strSs = strSs.Substring(6, 2) + "/" + strSs.Substring(8, 2) + "/20" + strSs.Substring(10, 2) + " " + strSs.Substring(4, 2) + ":" + strSs.Substring(2, 2) + ":" + strSs.Substring(0, 2);
            decValue = (decimal)int.Parse(strSs.Substring(2, 2).ToString());

            return strSs;
        }
        public string ToStringValue(string strinput, int count)
        {
            byte[] arrsent = clsConvert.ToByteArray(strinput);
            byte[] arrsentkq = new byte[count];
            int j = 0;
            for (int i = arrsent.Length - 1; i >= 0; i--)
            {
                arrsentkq[j] = (byte)(arrsent[i] - 0x33);
                j++;
            }

            return clsConvert.ByteArrayToString(arrsentkq);

        }

        public string ToStringValue_nomal(string strinput, int count)
        {
            byte[] arrsent = clsConvert.ToByteArray(strinput);
            byte[] arrsentkq = new byte[count];
            int j = 0;
            for (int i = arrsent.Length - 1; i >= 0; i--)
            {
                arrsentkq[j] = (byte)(arrsent[i]);
                j++;
            }

            return clsConvert.ByteArrayToString(arrsentkq);

        }


        public string ToStringValue(string strinput)
        {
            byte[] arrsent = clsConvert.ToByteArray(strinput);
            byte[] arrsentkq = new byte[4];
            int j = 0;
            for (int i = arrsent.Length - 1; i >= 0; i--)
            {
                arrsentkq[j] = (byte)(arrsent[i] - 0x33);
                j++;
            }

            return clsConvert.ByteArrayToString(arrsentkq);

        }
        public string ToStringValue_nomal(string strinput)
        {
            byte[] arrsent = clsConvert.ToByteArray(strinput);
            byte[] arrsentkq = new byte[4];
            int j = 0;
            for (int i = arrsent.Length - 1; i >= 0; i--)
            {
                arrsentkq[j] = (byte)(arrsent[i]);
                j++;
            }

            return clsConvert.ByteArrayToString(arrsentkq);

        }

        public string ToStringValue8Byte(string strinput)
        {
            byte[] arrsent = clsConvert.ToByteArray(strinput);
            byte[] arrsentkq = new byte[8];
            int j = 0;
            for (int i = arrsent.Length - 1; i >= 0; i--)
            {
                arrsentkq[j] = (byte)(arrsent[i] - 0x33);
                j++;
            }

            return clsConvert.ByteArrayToString(arrsentkq);

        }

        public string analyzeLoadHistory(string strMgs, ref dynamic obj)
        {
            // Kiểm tra định dạng message 
            string startmesssage = "68";
            string endMessage = "16";
            string tempdata = "";
            string strvalue = "";


            if (strMgs.Substring(0, 2) == startmesssage)
                if (strMgs.Substring(strMgs.Length - 2, 2) == endMessage)
                {
                    // Tìm đến vị trí có data
                    int vi_tri = strMgs.IndexOf("6891");
                    if (vi_tri < 0)
                    {
                        vi_tri = strMgs.IndexOf("68B1");
                        if (vi_tri < 0)
                        {
                            return "001";
                        }

                    }
                    strvalue = strMgs.Substring(vi_tri + 10, 8);
                    strvalue = ToStringValue(strvalue);
                    obj.decChotthang = double.Parse(strvalue) / 100;

                }
            return tempdata;

        }

        public string analyzeValue_EVENT(string strMgs, ref DateTime datetime)
        {
            string tempdata = "";
            string strvalue = "";
            string length = "";
            datetime = new DateTime(1000, 1, 1);
            try
            {
                if (strMgs != "000000000000")
                    datetime = DateTime.ParseExact(strvalue, "ssmmHHddMMyy", System.Globalization.CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {

            }


            return tempdata;

        }

        public string analyzeValue_GIO(string strMgs, ref DateTime datetime)
        {
            string tempdata = "";
            string strvalue = "";
            string length = "";
            datetime = new DateTime(1000, 1, 1);
            try
            {
                //strvalue = strMgs.Substring(vi_tri + 16, 12);
                strvalue = strMgs.Substring(0, 6); // ss;minute :h : 
                strvalue = strMgs.Substring(8, 6) + strvalue; // ss;minute :h : 
                datetime = DateTime.ParseExact(strvalue, "ddMMyyssmmHH", System.Globalization.CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {

            }


            return tempdata;

        }


        public string analyzeValue_cosfi_new(string strMgs, ref double objTong)
        {
            // Kiểm tra định dạng message 

            string strvalue = "";
            string length = "";
            length = strMgs.Substring(0, 2);
            if (length == "0C")
            {
                strvalue = strMgs.Substring(10, 4);
                strvalue = ToStringValue(strvalue, 2);
                objTong = double.Parse(strvalue) / 1000;
            }
            return "";

        }
        public string analyzeValue_cosfi_auto(string strMgs, ref double objTong)
        {
            // Kiểm tra định dạng message 
            string strvalue = "";
            objTong = 0;

            try
            {
                strvalue = strMgs.Substring(0, 4);
                if (strvalue != "EEEE")
                {
                    strvalue = ToStringValue_nomal(strvalue, 2);
                    objTong = double.Parse(strvalue) / 1000;
                }
                else
                    objTong = 0;
                return "";
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }




        }



        public string analyzeValue_tuti(string strMgs, ref double objTong)
        {
            // Kiểm tra định dạng message 

            string strvalue = "";
            string length = "";
            length = strMgs.Substring(0, 2);
            if (length == "08")
            {
                strvalue = strMgs.Substring(10, 8);
                strvalue = ToStringValue(strvalue, 4);
                objTong = double.Parse(strvalue);
            }
            return "";

        }

        public string analyzeValue_tuti_auto(string strMgs, ref double objTong)
        {
            objTong = 0;


            // Kiểm tra định dạng message 
            strMgs = ToStringValue_nomal(strMgs, 4);
            if (strMgs != "EEEE")
            {
                objTong = double.Parse(strMgs);
            }
            else
            {
                objTong = 0;
            }
            return "";

        }

        public string analyzeValue_new(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {
            // Kiểm tra định dạng message 

            string strvalue = "";
            string length = "";



            // Kiểm tra length để phân biệt loại công tơ 
            length = strMgs.Substring(0, 2);
            if (length == "08")
            {
                // Công tơ 1 pha 1 biểu gia 
                strvalue = strMgs.Substring(10, 8);
                strvalue = ToStringValue(strvalue);
                objTong = double.Parse(strvalue) / 100;
            }
            else
            {
                strvalue = strMgs.Substring(10, 8);
                strvalue = ToStringValue(strvalue);
                objTong = double.Parse(strvalue) / 100;
                strvalue = strMgs.Substring(18, 8);
                strvalue = ToStringValue(strvalue);
                objB1 = double.Parse(strvalue) / 100;
                strvalue = strMgs.Substring(26, 8);
                strvalue = ToStringValue(strvalue);
                objB2 = double.Parse(strvalue) / 100;
                strvalue = strMgs.Substring(34, 8);
                strvalue = ToStringValue(strvalue);
                objB3 = double.Parse(strvalue) / 100;
            }
            return "";

        }

        public string analyzeValue_auto(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {

            objTong = 0;
            objB1 = 0;
            objB2 = 0;
            objB3 = 0;

            try
            {
                // Kiểm tra định dạng message 
                if (strMgs == "")
                    return "";
                string strvalue = "";

                if (strMgs.Length == 8)
                {

                    strvalue = strMgs.Substring(0, 6);
                    strvalue = ToStringValue_nomal(strvalue, 3);
                    objTong = double.Parse(strvalue) / 10000;
                    return "";
                }
                else
                {
                    strvalue = strMgs.Substring(0, 6);
                    strvalue = ToStringValue_nomal(strvalue, 3);
                    objTong = double.Parse(strvalue) / 10000;
                    strvalue = strMgs.Substring(6, 6);
                    strvalue = ToStringValue_nomal(strvalue, 3);
                    objB1 = double.Parse(strvalue) / 10000;
                    strvalue = strMgs.Substring(12, 6);
                    strvalue = ToStringValue_nomal(strvalue, 3);
                    objB2 = double.Parse(strvalue) / 10000;
                    strvalue = strMgs.Substring(18, 6);
                    strvalue = ToStringValue_nomal(strvalue, 3);
                    objB3 = double.Parse(strvalue) / 10000;
                    return "";
                }

            }
            catch (Exception ex)
            {
                return "err";
            }

        }

        public string analyzeValue_auto4Byte(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {


            objTong = 0;
            objB1 = 0;
            objB2 = 0;
            objB3 = 0;
            // Kiểm tra định dạng message 
            string strvalue = "";
            if (strMgs == "")
                return "";

            if (strMgs.Length == 8)
            {

                strvalue = strMgs.Substring(0, 8);
                strvalue = ToStringValue_nomal(strvalue);
                objTong = double.Parse(strvalue) / 100;
                objB1 = 0;
                objB2 = 0;
                objB3 = 0;
            }
            else
            {

                strvalue = strMgs.Substring(0, 8);
                strvalue = ToStringValue_nomal(strvalue);
                objTong = double.Parse(strvalue) / 100;
                strvalue = strMgs.Substring(8, 8);
                strvalue = ToStringValue_nomal(strvalue);
                objB1 = double.Parse(strvalue) / 100;
                strvalue = strMgs.Substring(16, 8);
                strvalue = ToStringValue_nomal(strvalue);
                objB2 = double.Parse(strvalue) / 100;
                strvalue = strMgs.Substring(24, 8);
                strvalue = ToStringValue_nomal(strvalue);
                objB3 = double.Parse(strvalue) / 100;

            }



            return "";
        }

        public string analyzeValue_Loadprofile(string strMgs, ref double P_CONG, ref double P_TRU, ref double Q_CONG, ref double Q_TRU)
        {


            P_CONG = 0;
            P_TRU = 0;
            Q_CONG = 0;
            Q_TRU = 0;
            // Kiểm tra định dạng message 
            string strvalue = "";
            string temp = "";
            if (strMgs == "")
                return "";



            temp = strMgs.Substring(0, 8);
            int _length = Convert.ToInt32(strMgs.Substring(8, 2), 16);

            _length = _length * 2;
            _length = _length & 127;

            if (temp == "010400FF")
            {
                strvalue = strMgs.Substring(10, _length);
                strMgs = strMgs.Substring(10 + _length, strMgs.Length - (10 + _length));
                strvalue = analyzeValue_newpro(strvalue, ref P_CONG);
                P_CONG = P_CONG / 10000;
            }
            temp = strMgs.Substring(0, 8);
            _length = Convert.ToInt32(strMgs.Substring(8, 2), 16);

            _length = _length * 2;
            _length = _length & 127;
            if (temp == "020400FF")
            {
                strvalue = strMgs.Substring(10, _length);
                strMgs = strMgs.Substring(10 + _length, strMgs.Length - (10 + _length));

                strvalue = analyzeValue_newpro(strvalue, ref P_TRU);
                P_TRU = P_TRU / 10000;
            }
            temp = strMgs.Substring(0, 8);
            _length = Convert.ToInt32(strMgs.Substring(8, 2), 16);

            _length = _length * 2;
            _length = _length & 127;
            if (temp == "030400FF")
            {
                strvalue = strMgs.Substring(10, _length);
                strMgs = strMgs.Substring(10 + _length, strMgs.Length - (10 + _length));
                strvalue = analyzeValue_newpro(strvalue, ref Q_CONG);
                Q_CONG = Q_CONG / 10000;
            }
            temp = strMgs.Substring(0, 8);
            _length = Convert.ToInt32(strMgs.Substring(8, 2), 16);

            _length = _length * 2;
            _length = _length & 127;
            if (temp == "040400FF")
            {
                strvalue = strMgs.Substring(10, _length);
                strMgs = strMgs.Substring(10 + _length, strMgs.Length - (10 + _length));
                strvalue = analyzeValue_newpro(strvalue, ref Q_TRU);
                Q_TRU = Q_TRU / 10000;
            }

            return "";
        }

        public string analyzeValue_autoDonvitinh(string strMgs, ref double objTong)
        {
            objTong = 0;
            // Kiểm tra định dạng message 
            if (strMgs == "")
                return "";
            if (strMgs == "02")
            {
                objTong = 1000;
            }
            else
                objTong = 1;

            return "";
        }

        public string analyzeValue_dulieu(string strMgs, ref double objTong)
        {
            objTong = 0;
            // Kiểm tra định dạng message 
            if (strMgs == "")
                return "";
            byte o = 0;

            if (strMgs.Length > 2 && strMgs.Length <= 4)
                o = 1;
            else if (strMgs.Length > 4)

                o = 2;
            objTong = ConvertArrayToInt2(StringToByteArray(strMgs), o);

            return "";
        }

        public string analyzeValue_newpro(string strMgs, ref double objTong)
        {
            objTong = 0;
            // Kiểm tra định dạng message 

            if (strMgs == "")
                return "";
            int lenghtstr = strMgs.Length / 2;
            string tempstrMgs = "";
            if (lenghtstr == 2)
            {
                strMgs = strMgs.Substring(2, 2) + strMgs.Substring(0, 2);
            }
            if (lenghtstr == 3)
            {
                strMgs = strMgs.Substring(4, 2) + strMgs.Substring(2, 2) + strMgs.Substring(0, 2);
            }

            if (lenghtstr == 4)
            {
                strMgs = strMgs.Substring(6, 2) + strMgs.Substring(4, 2) + strMgs.Substring(2, 2) + strMgs.Substring(0, 2);
            }
            if (lenghtstr == 5)
            {
                strMgs = strMgs.Substring(8, 2) + strMgs.Substring(6, 2) + strMgs.Substring(4, 2) + strMgs.Substring(2, 2) + strMgs.Substring(0, 2);
            }
            if (lenghtstr == 6)
            {
                strMgs = strMgs.Substring(10, 2) + strMgs.Substring(8, 2) + strMgs.Substring(6, 2) + strMgs.Substring(4, 2) + strMgs.Substring(2, 2) + strMgs.Substring(0, 2);
            }
            if (lenghtstr == 7)
            {
                strMgs = strMgs.Substring(12, 2) + strMgs.Substring(10, 2) + strMgs.Substring(8, 2) + strMgs.Substring(6, 2) + strMgs.Substring(4, 2) + strMgs.Substring(2, 2) + strMgs.Substring(0, 2);
            }
            if (lenghtstr == 8)
            {
                strMgs = strMgs.Substring(14, 2) + strMgs.Substring(12, 2) + strMgs.Substring(10, 2) + strMgs.Substring(8, 2) + strMgs.Substring(6, 2) + strMgs.Substring(4, 2) + strMgs.Substring(2, 2) + strMgs.Substring(0, 2);
            }


            objTong = int.Parse(strMgs, System.Globalization.NumberStyles.HexNumber);

            return "";
        }


        public string analyzeValue_QTong(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {

            objTong = 0;
            objB1 = 0;
            objB2 = 0;
            objB3 = 0;
            // Kiểm tra định dạng message 
            if (strMgs == "")
                return "";

            string strvalue = "";
            strvalue = strMgs.Substring(0, 8);
            strvalue = ToStringValue_nomal(strvalue);
            objTong = double.Parse(strvalue) / 100;
            return "";
        }
        public string analyzeValue_pmax(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3, ref DateTime? timePmax)
        {
            // Kiểm tra định dạng message 

            string strvalue = "";
            string length = "";



            // Kiểm tra length để phân biệt loại công tơ 
            length = strMgs.Substring(0, 2);
            if (length == "08")
            {
                // Công tơ 1 pha 1 biểu gia 
                strvalue = strMgs.Substring(10, 6);
                strvalue = ToStringValue(strvalue);
                objTong = double.Parse(strvalue) / 10000;
            }
            else
            {
                // Giá trị Pmax1
                strvalue = strMgs.Substring(10, 6);
                strvalue = ToStringValue(strvalue);
                objTong = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(16, 10);
                strvalue = ToStringValue(strvalue, 5);
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
                // phare time Pmax

                // Giá trị Pmax2
                strvalue = strMgs.Substring(26, 6);
                strvalue = ToStringValue(strvalue);
                objB1 = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(32, 10);
                // phare time Pmax
                strvalue = ToStringValue(strvalue, 5);
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);

                // Giá trị Pmax 3 
                strvalue = strMgs.Substring(42, 6);
                strvalue = ToStringValue(strvalue);
                objB2 = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(48, 10);
                strvalue = ToStringValue(strvalue, 5);
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
                // Giá trị Pmax Tổng  
                strvalue = strMgs.Substring(58, 6);
                strvalue = ToStringValue(strvalue);
                objB3 = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(64, 10);
                strvalue = ToStringValue(strvalue, 5);
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
            }



            return "";

        }
        public int ConvertArrayToInt(string input, byte opt = 0)
        {
            byte[] b = Encoding.ASCII.GetBytes(input);

            if (opt == 0)
            {
                return BitConverter.ToInt32(b, 0);
                //return BitConverter.ToUInt16(b, 0);
            }
            return BitConverter.ToInt32(b, 0);
        }


        public string analyzeValue_pmax_newPro(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3, ref DateTime timePmax)
        {
            // Kiểm tra định dạng message 
            objTong = 0;
            objB1 = 0;
            objB2 = 0;
            objB3 = 0;
            timePmax = new DateTime(1000, 1, 1);
            string strvalue = "";
            string length = "";
            timePmax = new DateTime(1000, 1, 1);



            // Kiểm tra length để phân biệt loại công tơ 
            if (strMgs.Length == 16)
            {
                // Giá trị Pmax1
                strvalue = strMgs.Substring(0, 6);
                strvalue = ToStringValue_nomal(strvalue, 3);
                objTong = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(6, 10);
                strvalue = ToStringValue_nomal(strvalue, 5);
                if (strvalue == "0000000000")
                    return "";
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
                // phare time Pmax

            }
            else
            {
                // Tính độ dài của biến 

                strvalue = strMgs.Substring(0, strMgs.Length - 10);
                strvalue = ToStringValue_nomal(strvalue, strvalue.Length / 2);
                // Giá trị Pmax1
                objTong = int.Parse(strvalue, System.Globalization.NumberStyles.HexNumber);
                objTong = objTong / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(strMgs.Length - 10, 10);
                strvalue = ToStringValue_nomal(strvalue, 5);
                if (strvalue == "0000000000")
                    return "";
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
                // phare time Pmax
            }

            return "";

        }




        public string analyzeValue_pmax_auto(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3, ref DateTime timePmax)
        {
            // Kiểm tra định dạng message 
            objTong = 0;
            objB1 = 0;
            objB2 = 0;
            objB3 = 0;
            timePmax = new DateTime(1000, 1, 1);
            string strvalue = "";
            string length = "";
            timePmax = new DateTime(1000, 1, 1);



            // Kiểm tra length để phân biệt loại công tơ 
            if (strMgs.Length == 16)
            {
                // Giá trị Pmax1
                strvalue = strMgs.Substring(0, 6);
                strvalue = ToStringValue_nomal(strvalue, 3);
                objTong = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(6, 10);
                strvalue = ToStringValue_nomal(strvalue, 5);
                if (strvalue == "0000000000")
                    return "";
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
                // phare time Pmax

            }
            else
            {
                // Giá trị Pmax1
                strvalue = strMgs.Substring(0, 6);
                strvalue = ToStringValue_nomal(strvalue, 3);
                objTong = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(6, 10);
                strvalue = ToStringValue_nomal(strvalue, 5);
                if (strvalue == "0000000000")
                    return "";
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
                // phare time Pmax

                // Giá trị Pmax2
                strvalue = strMgs.Substring(16, 6);
                strvalue = ToStringValue_nomal(strvalue, 3);
                objB1 = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(22, 10);
                // phare time Pmax
                strvalue = ToStringValue_nomal(strvalue, 5);
                //   timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);

                // Giá trị Pmax 3 
                strvalue = strMgs.Substring(32, 6);
                strvalue = ToStringValue_nomal(strvalue, 3);
                objB2 = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(38, 10);
                strvalue = ToStringValue_nomal(strvalue, 5);
                //   timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
                // Giá trị Pmax Tổng  
                strvalue = strMgs.Substring(48, 6);
                strvalue = ToStringValue_nomal(strvalue, 3);
                objB3 = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(54, 10);
                strvalue = ToStringValue_nomal(strvalue, 5);
                //    timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
            }

            return "";

        }

        public string analyzeValue_pmax(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3, ref DateTime timePmax)
        {
            // Kiểm tra định dạng message 

            string strvalue = "";
            string length = "";



            // Kiểm tra length để phân biệt loại công tơ 
            length = strMgs.Substring(0, 2);
            if (length == "08")
            {
                // Công tơ 1 pha 1 biểu gia 
                strvalue = strMgs.Substring(10, 6);
                strvalue = ToStringValue(strvalue);
                objTong = double.Parse(strvalue) / 100;
            }
            else
            {
                // Giá trị Pmax1
                strvalue = strMgs.Substring(10, 6);
                strvalue = ToStringValue(strvalue);
                objTong = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(16, 10);
                strvalue = ToStringValue(strvalue, 5);
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
                // phare time Pmax

                // Giá trị Pmax2
                strvalue = strMgs.Substring(26, 6);
                strvalue = ToStringValue(strvalue);
                objB1 = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(32, 10);
                // phare time Pmax
                strvalue = ToStringValue(strvalue, 5);
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);

                // Giá trị Pmax 3 
                strvalue = strMgs.Substring(42, 6);
                strvalue = ToStringValue(strvalue);
                objB2 = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(48, 10);
                strvalue = ToStringValue(strvalue, 5);
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
                // Giá trị Pmax Tổng  
                strvalue = strMgs.Substring(58, 6);
                strvalue = ToStringValue(strvalue);
                objB3 = double.Parse(strvalue) / 10000;
                // Thời gian xảy ra Pmax
                strvalue = strMgs.Substring(64, 10);
                strvalue = ToStringValue(strvalue, 5);
                timePmax = DateTime.ParseExact(strvalue, "yyMMddHHmm", null);
            }



            return "";

        }


        public string analyzeValue_Dienap_auto(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {
            // Kiểm tra định dạng message 
            string tempdata = "";
            string strvalue = "";
            objTong = 0;
            objB1 = 0;
            objB2 = 0;
            objB3 = 0;


            strvalue = strMgs.Substring(0, 4);
            strvalue = ToStringValue_nomal(strvalue, 2);
            objB1 = double.Parse(strvalue) / 10;
            strvalue = strMgs.Substring(4, 4);
            strvalue = ToStringValue_nomal(strvalue, 2);
            objB2 = double.Parse(strvalue) / 10;
            strvalue = strMgs.Substring(8, 4);
            strvalue = ToStringValue_nomal(strvalue, 2);
            objB3 = double.Parse(strvalue) / 10;

            return tempdata;

        }




        public string analyzeValue_Dienap_new(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {
            // Kiểm tra định dạng message 
            string startmesssage = "68";
            string endMessage = "16";
            string tempdata = "";
            string strvalue = "";
            string length = "";
            // Tìm đến vị trí có data
            length = strMgs.Substring(0, 2);
            if (length == "08")
            {
                // Công tơ 1 pha 1 biểu gia 
                strvalue = strMgs.Substring(14, 8);
                strvalue = ToStringValue(strvalue);
                objTong = double.Parse(strvalue) / 100;
            }
            else
            {
                strvalue = strMgs.Substring(10, 4);
                strvalue = ToStringValue(strvalue, 2);
                objB1 = double.Parse(strvalue) / 10;
                strvalue = strMgs.Substring(14, 4);
                strvalue = ToStringValue(strvalue, 2);
                objB2 = double.Parse(strvalue) / 10;
                strvalue = strMgs.Substring(18, 4);
                strvalue = ToStringValue(strvalue, 2);
                objB3 = double.Parse(strvalue) / 10;

            }



            return tempdata;

        }

        public string analyzeValue(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {
            // Kiểm tra định dạng message 
            string startmesssage = "68";
            string endMessage = "16";
            string tempdata = "";
            string strvalue = "";
            string length = "";


            if (strMgs.Substring(0, 2) == startmesssage)
                if (strMgs.Substring(strMgs.Length - 2, 2) == endMessage)
                {
                    // Tìm đến vị trí có data
                    int vi_tri = strMgs.IndexOf("6891");
                    if (vi_tri < 0)
                    {
                        vi_tri = strMgs.IndexOf("68B1");
                        if (vi_tri < 0)
                        {
                            return "001";
                        }
                    }
                    // Kiểm tra length để phân biệt loại công tơ 
                    length = strMgs.Substring(vi_tri + 4, 2);
                    if (length == "08")
                    {
                        // Công tơ 1 pha 1 biểu gia 
                        strvalue = strMgs.Substring(vi_tri + 14, 8);
                        strvalue = ToStringValue(strvalue);
                        objTong = double.Parse(strvalue) / 100;
                    }
                    else
                    {
                        strvalue = strMgs.Substring(vi_tri + 14, 8);
                        strvalue = ToStringValue(strvalue);
                        objTong = double.Parse(strvalue) / 100;
                        strvalue = strMgs.Substring(vi_tri + 22, 8);
                        strvalue = ToStringValue(strvalue);
                        objB1 = double.Parse(strvalue) / 100;
                        strvalue = strMgs.Substring(vi_tri + 30, 8);
                        strvalue = ToStringValue(strvalue);
                        objB2 = double.Parse(strvalue) / 100;
                        strvalue = strMgs.Substring(vi_tri + 38, 8);
                        strvalue = ToStringValue(strvalue);
                        objB3 = double.Parse(strvalue) / 100;


                    }


                }
            return "";

        }

        public string analyzeValue_3byte(string strMgs, ref double objValue)
        {
            // Kiểm tra định dạng message 
            string startmesssage = "68";
            string endMessage = "16";
            string tempdata = "";
            string strvalue = "";
            string length = "";


            if (strMgs.Substring(0, 2) == startmesssage)
                if (strMgs.Substring(strMgs.Length - 2, 2) == endMessage)
                {
                    // Tìm đến vị trí có data
                    int vi_tri = strMgs.IndexOf("6891");
                    if (vi_tri < 0)
                    {
                        vi_tri = strMgs.IndexOf("68B1");
                        if (vi_tri < 0)
                        {
                            return "001";
                        }
                    }
                    // Kiểm tra length để phân biệt loại công tơ 
                    length = strMgs.Substring(vi_tri + 4, 2);
                    if (length == "07")
                    {
                        // Công tơ 1 pha 1 biểu gia 
                        strvalue = strMgs.Substring(vi_tri + 14, 6);
                        strvalue = ToStringValue(strvalue, 3);
                        objValue = double.Parse(strvalue) / 100;


                    }


                }
            return "";

        }

        public string analyzeValue_P(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {
            // Kiểm tra định dạng message 
            string startmesssage = "68";
            string endMessage = "16";
            string tempdata = "";
            string strvalue = "";
            string length = "";


            if (strMgs.Substring(0, 2) == startmesssage)
                if (strMgs.Substring(strMgs.Length - 2, 2) == endMessage)
                {
                    // Tìm đến vị trí có data
                    int vi_tri = strMgs.IndexOf("6891");
                    if (vi_tri < 0)
                    {
                        vi_tri = strMgs.IndexOf("68B1");
                        if (vi_tri < 0)
                        {
                            // MessageBox.Show("Chuỗi không đúng");
                            return "001";

                        }

                    }
                    // Kiểm tra length để phân biệt loại công tơ 
                    length = strMgs.Substring(vi_tri + 4, 2);
                    if (length == "08")
                    {

                    }
                    else
                    {
                        strvalue = strMgs.Substring(vi_tri + 14, 6);
                        strvalue = ToStringValue(strvalue);
                        objTong = double.Parse(strvalue) / 10;
                        strvalue = strMgs.Substring(vi_tri + 20, 6);
                        strvalue = ToStringValue(strvalue);
                        objB1 = double.Parse(strvalue) / 10;
                        strvalue = strMgs.Substring(vi_tri + 26, 6);
                        strvalue = ToStringValue(strvalue);
                        objB2 = double.Parse(strvalue) / 10;
                        strvalue = strMgs.Substring(vi_tri + 32, 6);
                        strvalue = ToStringValue(strvalue);
                        objB3 = double.Parse(strvalue) / 10;


                    }


                }
            return tempdata;

        }



        public string analyzeValue_Dienap(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {
            // Kiểm tra định dạng message 
            string startmesssage = "68";
            string endMessage = "16";
            string tempdata = "";
            string strvalue = "";
            string length = "";


            if (strMgs.Substring(0, 2) == startmesssage)
                if (strMgs.Substring(strMgs.Length - 2, 2) == endMessage)
                {
                    // Tìm đến vị trí có data
                    int vi_tri = strMgs.LastIndexOf("6891");
                    if (vi_tri < 0)
                    {
                        vi_tri = strMgs.LastIndexOf("68B1");
                        if (vi_tri < 0)
                            return "001";
                    }
                    // Kiểm tra length để phân biệt loại công tơ 
                    length = strMgs.Substring(vi_tri + 4, 2);
                    if (length == "08")
                    {
                        // Công tơ 1 pha 1 biểu gia 
                        strvalue = strMgs.Substring(vi_tri + 14, 8);
                        strvalue = ToStringValue(strvalue);
                        objTong = double.Parse(strvalue) / 100;
                    }
                    else
                    {
                        strvalue = strMgs.Substring(vi_tri + 14, 4);
                        strvalue = ToStringValue(strvalue, 2);
                        objTong = double.Parse(strvalue) / 10;
                        strvalue = strMgs.Substring(vi_tri + 18, 4);
                        strvalue = ToStringValue(strvalue, 2);
                        objB1 = double.Parse(strvalue) / 10;
                        strvalue = strMgs.Substring(vi_tri + 22, 4);
                        strvalue = ToStringValue(strvalue, 2);
                        objB2 = double.Parse(strvalue) / 10;

                    }


                }
            return tempdata;

        }


        public string analyzeValue_Dongdien_auto(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {
            // Kiểm tra định dạng message 
            string tempdata = "";
            string strvalue = "";
            string length = "";
            objTong = 0;
            objB1 = 0;
            objB2 = 0;
            objB3 = 0;


            // Kiểm tra length để phân biệt loại công tơ 
            strvalue = strMgs.Substring(0, 6);
            strvalue = ToStringValue_nomal(strvalue, 3);
            objB1 = double.Parse(strvalue) / 1000;
            strvalue = strMgs.Substring(6, 6);
            strvalue = ToStringValue_nomal(strvalue, 3);
            objB2 = double.Parse(strvalue) / 1000;
            strvalue = strMgs.Substring(12, 6);
            strvalue = ToStringValue_nomal(strvalue, 3);
            objB3 = double.Parse(strvalue) / 1000;
            objTong = 0;
            return tempdata;

        }


        public string analyzeValue_Dongdien_new(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {
            // Kiểm tra định dạng message 
            string tempdata = "";
            string strvalue = "";
            string length = "";




            // Kiểm tra length để phân biệt loại công tơ 
            length = strMgs.Substring(0, 2);
            if (length == "08")
            {
                // Công tơ 1 pha 1 biểu gia 
                strvalue = strMgs.Substring(8, 6);
                strvalue = ToStringValue(strvalue);
                objTong = double.Parse(strvalue) / 1000;
            }
            else
            {

                strvalue = strMgs.Substring(08, 6);
                strvalue = ToStringValue(strvalue, 3);
                objB1 = double.Parse(strvalue) / 1000;
                strvalue = strMgs.Substring(14, 6);
                strvalue = ToStringValue(strvalue, 3);
                objB2 = double.Parse(strvalue) / 1000;
                strvalue = strMgs.Substring(20, 6);
                strvalue = ToStringValue(strvalue, 3);
                objB3 = double.Parse(strvalue) / 1000;
                objTong = 0;

            }



            return tempdata;

        }

        public string analyzeValue_Dongdien(string strMgs, ref double objTong, ref double objB1, ref double objB2, ref double objB3)
        {
            // Kiểm tra định dạng message 
            string startmesssage = "68";
            string endMessage = "16";
            string tempdata = "";
            string strvalue = "";
            string length = "";


            if (strMgs.Substring(0, 2) == startmesssage)
                if (strMgs.Substring(strMgs.Length - 2, 2) == endMessage)
                {
                    // Tìm đến vị trí có data
                    int vi_tri = strMgs.IndexOf("6891");
                    if (vi_tri < 0)
                    {
                        vi_tri = strMgs.IndexOf("68B1");
                        if (vi_tri < 0)
                            return "001";
                    }
                    // Kiểm tra length để phân biệt loại công tơ 
                    length = strMgs.Substring(vi_tri + 4, 2);
                    if (length == "08")
                    {
                        // Công tơ 1 pha 1 biểu gia 
                        strvalue = strMgs.Substring(vi_tri + 14, 6);
                        strvalue = ToStringValue(strvalue);
                        objTong = double.Parse(strvalue) / 1000;
                    }
                    else
                    {
                        strvalue = strMgs.Substring(vi_tri + 14, 6);
                        strvalue = ToStringValue(strvalue, 3);
                        objTong = double.Parse(strvalue) / 1000;
                        strvalue = strMgs.Substring(vi_tri + 20, 6);
                        strvalue = ToStringValue(strvalue, 3);
                        objB1 = double.Parse(strvalue) / 1000;
                        strvalue = strMgs.Substring(vi_tri + 26, 6);
                        strvalue = ToStringValue(strvalue, 3);
                        objB2 = double.Parse(strvalue) / 1000;

                    }


                }
            return tempdata;

        }
        public string analyzeLoadProfile75(string strMgs, ref dynamic obj)
        {
            // Kiểm tra định dạng message 
            string startmesssage = "68";
            string endMessage = "16";
            string tempdata = "";
            int ilength = 0;
            string malenh = "";
            string strCheckLoaiMsg = "";
            string strmessagesAnalyze = "";
            string strvalue = "";

            string templength = "";
            string data = "";
            double declength = 0;
            string strngay = "";
            string strStartdate = "";
            string strEnddate = "";

            //MA_DIEMDO Mã điểm đo
            //NGAY Ngày
            //STARTDATE Thời gian bắt đầu
            //ENDDATE Thời gian kết thúc
            //IMPORTKW Chỉ số công suất Tổng giao
            //C1  Chỉ số vô công giao
            //EXPORTKW    Chỉ số công suất Tổng nhận
            //C2 Chỉ số vô công nhận
            //FLAG Trạng thái dữ liệu

            int vi_tri = 0;
            if (strMgs.Substring(0, 2) == startmesssage)
                if (strMgs.Substring(strMgs.Length - 2, 2) == endMessage)
                {
                    // Tìm đến vị trí có data
                    vi_tri = strMgs.IndexOf("6891");
                    if (vi_tri < 0)
                    {
                        vi_tri = strMgs.IndexOf("68B1");
                        if (vi_tri < 0)
                        {
                            // 001 giao thức không đúng
                            return "001";
                        }

                    }
                    data = strMgs.Substring(vi_tri, strMgs.Length - vi_tri);
                    templength = data.Substring(4, 2).ToString();
                    if (templength == "04")
                        return "002";
                    declength = (int)int.Parse(templength, System.Globalization.NumberStyles.HexNumber);
                    //34333339-- > DI load profile
                    //D3D3-- > Đánh dấu vị trí bắt đầu ngày
                    malenh = data.Substring(6, 8);

                    if (malenh != "34333339")
                    {
                        data = data.Substring(4, data.Length - 4);
                        // Tìm đến vị trí có data
                        vi_tri = data.IndexOf("6891");
                        if (vi_tri < 0)
                        {
                            vi_tri = data.IndexOf("68B1");
                            if (vi_tri < 0)
                            {
                                // 001 giao thức không đúng
                                return "001";
                            }

                        }
                        data = data.Substring(vi_tri, data.Length - vi_tri);

                    }

                    //       strma = data.Substring(14, 4);
                    vi_tri = data.IndexOf("D3"); // Tìm vị tri D3 đầu tiên 
                    if (vi_tri < 0)
                    {
                        // MessageBox.Show("Chuỗi không đúng analyzeLoadProfile");
                        return "001-D3-1";
                    }
                    strmessagesAnalyze = data.Substring(vi_tri + 2, data.Length - vi_tri - 2);
                    vi_tri = strmessagesAnalyze.IndexOf("D3"); // Tìm vị tri D3 thứ 2 
                    if (vi_tri < 0)
                    {
                        //   MessageBox.Show("Chuỗi không đúng analyzeLoadProfile");
                        return "001-D3-2";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
                    strngay = strmessagesAnalyze.Substring(0, 12);
                    strngay = ToStringDate(strngay, ref strStartdate, ref strEnddate);
                    //obj.P_STARTDATE = DateTime.ParseExact(strStartdate, "dd/MM/yyyy HH:mm", System.Globalization.CultureInfo.InvariantCulture);
                    //obj.P_ENDDATE = DateTime.ParseExact(strStartdate, "dd/MM/yyyy HH:mm", System.Globalization.CultureInfo.InvariantCulture).AddMinutes(30);
                    //obj.P_NGAY = DateTime.ParseExact(strngay, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    // Xử lý đoạn lấy UA, UB, UC, IA, IB,IC

                    vi_tri = strmessagesAnalyze.IndexOf("DD"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
                    if (vi_tri < 0)
                    {
                        //  MessageBox.Show("Chuỗi không đúng analyzeLoadProfile");
                        return "001-01";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);

                    vi_tri = strmessagesAnalyze.IndexOf("DD"); // DD Cosfi 
                    if (vi_tri < 0)
                    {
                        return "001-02";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
                    vi_tri = strmessagesAnalyze.IndexOf("DD"); // DD Điện năng (Chỉ số P giao P nhận Q giao Q nhận )
                    if (vi_tri < 0)
                    {
                        return "001-03";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);

                    strvalue = strmessagesAnalyze.Substring(0, 2);
                    if (strvalue != "DD")
                    {
                        // P giao 
                        tempdata = strmessagesAnalyze.Substring(0, 8);
                        tempdata = ToStringValue(tempdata);
                        obj.P_IMPORTKW_CS_TRUOC = double.Parse(tempdata) / 1000;   //IMPORTKW P_IMPORTKW Chỉ số công suất Tổng giao
                        strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                        // P nhận
                        tempdata = strmessagesAnalyze.Substring(0, 8);
                        tempdata = ToStringValue(tempdata);
                        strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                        obj.P_EXPORTKW_CS_TRUOC = double.Parse(tempdata) / 1000;
                        // Q Nhận
                        tempdata = strmessagesAnalyze.Substring(0, 8);
                        tempdata = ToStringValue(tempdata);
                        strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                        obj.P_EXPORTKW_CS_TRUOC = double.Parse(tempdata) / 1000;
                        // Q Giao 
                        tempdata = strmessagesAnalyze.Substring(0, 8);
                        tempdata = ToStringValue(tempdata);
                        strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                        obj.P_EXPORTKW_CS_TRUOC = double.Parse(tempdata) / 1000;
                    }


                    // Tìm đến vị trí tiếp theo để lấy PMAX
                    // Xử lý trong trường hợp công tơ không cấu hình 
                    vi_tri = strmessagesAnalyze.IndexOf("DD"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
                    if (vi_tri < 0)
                    {
                        //MessageBox.Show("Phân tích Q1 Q2 Q3 Q4  lỗi analyzeLoadProfile");
                        return "001-04";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
                    if (strmessagesAnalyze.Substring(0, 2) == "DD")
                    {
                        // Khong xu ly gì cả vì dữ liệu không cài đặt 
                        // Set thông tin =0 
                    }
                    else
                    {
                        // Nếu lấy Q thì sẽ chèn thông số vào đây 
                    }
                    vi_tri = strmessagesAnalyze.IndexOf("DD"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
                    if (vi_tri < 0)
                    {
                        //MessageBox.Show("Phân tích PMAX,QMAX  lỗi analyzeLoadProfile");
                        return "001-05";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
                    if (strmessagesAnalyze.Substring(0, 2) == "DD")
                    {
                        // Khong xu ly gì cả vì dữ liệu không cài đặt 
                        // Set thông tin =0 
                    }
                    else
                    {
                        // sẽ chèn thông số vào đây 
                        // 3 byte Pmax, 3 byte Q max
                        tempdata = strmessagesAnalyze.Substring(0, 6);
                        strmessagesAnalyze = strmessagesAnalyze.Substring(6, strmessagesAnalyze.Length - 6);
                        tempdata = ToStringValue(tempdata);
                        // Kiểm tra trạng thái P âm hay P dương để thực hiện truyền vào Giao hay nhận
                        string checkdau = "";
                        checkdau = tempdata.Substring(0, 2);
                        if (checkdau == "80")
                        {
                            tempdata = tempdata.Substring(2, tempdata.Length - 2);
                            // obj.P_MAX = double.Parse(tempdata) / 1000;
                            //    obj.P_EXPORTKW = double.Parse(tempdata) / 1000; // Vô công giao 
                        }
                        else
                        {
                            //    obj.P_IMPORTKW = double.Parse(tempdata) / 1000;

                        }
                        // Đoạn này lấy Q 
                        tempdata = strmessagesAnalyze.Substring(0, 6);
                        strmessagesAnalyze = strmessagesAnalyze.Substring(6, strmessagesAnalyze.Length - 6);
                        tempdata = ToStringValue(tempdata, 3);
                        if (tempdata == "ffffff")
                        {


                        }
                        else
                        {
                            checkdau = tempdata.Substring(0, 2);
                            if (checkdau == "80")
                            {
                                tempdata = tempdata.Substring(2, tempdata.Length - 2);
                                // obj.P_MAX = double.Parse(tempdata) / 1000;
                                //     obj.P_C1 = double.Parse(tempdata) / 1000; // Vô công giao 
                            }
                            else
                            {
                                //     obj.P_C2 = double.Parse(tempdata) / 1000;

                            }
                        }


                    }

                    //strvalue = strMgs.Substring(vi_tri + 6, 8);
                    //strvalue = ToStringValue(strvalue);
                    //obj.P_IMPORTKW = double.Parse(strvalue) / 1000;   //IMPORTKW P_IMPORTKW Chỉ số công suất Tổng giao

                    //strvalue = strMgs.Substring(vi_tri + 14, 8);
                    //strvalue = ToStringValue(strvalue);
                    //obj.P_EXPORTKW = double.Parse(strvalue) / 1000; //  //EXPORTKW    Chỉ số công suất Tổng nhận
                    //strvalue = strMgs.Substring(vi_tri + 22, 8);
                    //strvalue = ToStringValue(strvalue);
                    //obj.P_C1 = double.Parse(strvalue) / 1000; // Vô công giao 
                    //strvalue = strMgs.Substring(vi_tri + 30, 8);
                    //strvalue = ToStringValue(strvalue);
                    //obj.P_C2 = double.Parse(strvalue) / 1000;  // Vô công nhận
                    //vi_tri = tempdata.IndexOf("DDDD");
                    //if (vi_tri < 0)
                    //{
                    //    MessageBox.Show("Chuỗi không đúng");
                    //}
                    //strvalue = tempdata.Substring(vi_tri + 4, 6);
                    //strvalue = ToStringValue(strvalue);
                    //obj.decCongsuat = double.Parse(strvalue) / 10000; // P 
                    //strvalue = tempdata.Substring(vi_tri + 10, 6);
                    //strvalue = ToStringValue(strvalue);
                    //obj.P_Q1 = double.Parse(strvalue) / 10000; // Q

                }
                else
                {
                    return "007"; // frame messsage loi 
                }
            return "";

        }
        public string analyzeLoadProfile(string strMgs, ref dynamic obj)
        {
            // Kiểm tra định dạng message 
            string startmesssage = "68";
            string endMessage = "16";
            string tempdata = "";
            int ilength = 0;
            string malenh = "";
            string strCheckLoaiMsg = "";
            string strmessagesAnalyze = "";
            string strvalue = "";

            string templength = "";
            string data = "";
            double declength = 0;
            string strngay = "";
            string strStartdate = "";
            string strEnddate = "";

            //MA_DIEMDO Mã điểm đo
            //NGAY Ngày
            //STARTDATE Thời gian bắt đầu
            //ENDDATE Thời gian kết thúc
            //IMPORTKW Chỉ số công suất Tổng giao
            //C1  Chỉ số vô công giao
            //EXPORTKW    Chỉ số công suất Tổng nhận
            //C2 Chỉ số vô công nhận
            //FLAG Trạng thái dữ liệu

            int vi_tri = 0;
            if (strMgs.Substring(0, 2) == startmesssage)
                if (strMgs.Substring(strMgs.Length - 2, 2) == endMessage)
                {
                    // Tìm đến vị trí có data
                    vi_tri = strMgs.IndexOf("6891");
                    if (vi_tri < 0)
                    {
                        vi_tri = strMgs.IndexOf("68B1");
                        if (vi_tri < 0)
                        {
                            // 001 giao thức không đúng
                            return "001";
                        }

                    }
                    data = strMgs.Substring(vi_tri, strMgs.Length - vi_tri);
                    templength = data.Substring(4, 2).ToString();
                    if (templength == "04")
                        return "002";
                    declength = (int)int.Parse(templength, System.Globalization.NumberStyles.HexNumber);
                    //34333339-- > DI load profile
                    //D3D3-- > Đánh dấu vị trí bắt đầu ngày
                    malenh = data.Substring(6, 8);

                    if (malenh != "34333339")
                    {
                        data = data.Substring(4, data.Length - 4);
                        // Tìm đến vị trí có data
                        vi_tri = data.IndexOf("6891");
                        if (vi_tri < 0)
                        {
                            vi_tri = data.IndexOf("68B1");
                            if (vi_tri < 0)
                            {
                                // 001 giao thức không đúng
                                return "001";
                            }

                        }
                        data = data.Substring(vi_tri, data.Length - vi_tri);

                    }

                    //       strma = data.Substring(14, 4);
                    vi_tri = data.IndexOf("D3"); // Tìm vị tri D3 đầu tiên 
                    if (vi_tri < 0)
                    {
                        // MessageBox.Show("Chuỗi không đúng analyzeLoadProfile");
                        return "001-D3-1";
                    }
                    strmessagesAnalyze = data.Substring(vi_tri + 2, data.Length - vi_tri - 2);
                    vi_tri = strmessagesAnalyze.IndexOf("D3"); // Tìm vị tri D3 thứ 2 
                    if (vi_tri < 0)
                    {
                        //   MessageBox.Show("Chuỗi không đúng analyzeLoadProfile");
                        return "001-D3-2";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
                    strngay = strmessagesAnalyze.Substring(0, 12);
                    strngay = ToStringDate(strngay, ref strStartdate, ref strEnddate);
                    obj.P_STARTDATE = DateTime.ParseExact(strStartdate, "dd/MM/yyyy HH:mm", System.Globalization.CultureInfo.InvariantCulture);
                    obj.P_ENDDATE = DateTime.ParseExact(strStartdate, "dd/MM/yyyy HH:mm", System.Globalization.CultureInfo.InvariantCulture).AddMinutes(30);
                    obj.P_NGAY = DateTime.ParseExact(strngay, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    // Xử lý đoạn lấy UA, UB, UC, IA, IB,IC

                    vi_tri = strmessagesAnalyze.IndexOf("DD"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
                    if (vi_tri < 0)
                    {
                        //  MessageBox.Show("Chuỗi không đúng analyzeLoadProfile");
                        return "001-01";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);

                    vi_tri = strmessagesAnalyze.IndexOf("DD"); // DD Cosfi 
                    if (vi_tri < 0)
                    {
                        return "001-02";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
                    vi_tri = strmessagesAnalyze.IndexOf("DD"); // DD Điện năng (Chỉ số P giao P nhận Q giao Q nhận )
                    if (vi_tri < 0)
                    {
                        return "001-03";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);

                    strvalue = strmessagesAnalyze.Substring(0, 2);
                    if (strvalue != "DD")
                    {
                        // P giao 
                        tempdata = strmessagesAnalyze.Substring(0, 8);
                        tempdata = ToStringValue(tempdata);
                        obj.P_IMPORTKW_CS = double.Parse(tempdata) / 1000;   //IMPORTKW P_IMPORTKW Chỉ số công suất Tổng giao
                        strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                        // P nhận
                        tempdata = strmessagesAnalyze.Substring(0, 8);
                        tempdata = ToStringValue(tempdata);
                        strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                        obj.P_EXPORTKW_CS = double.Parse(tempdata) / 1000;
                        // Q Nhận
                        tempdata = strmessagesAnalyze.Substring(0, 8);
                        tempdata = ToStringValue(tempdata);
                        strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                        obj.P_EXPORTKW_CS = double.Parse(tempdata) / 1000;
                        // Q Giao 
                        tempdata = strmessagesAnalyze.Substring(0, 8);
                        tempdata = ToStringValue(tempdata);
                        strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                        obj.P_EXPORTKW_CS = double.Parse(tempdata) / 1000;
                    }


                    // Tìm đến vị trí tiếp theo để lấy PMAX
                    // Xử lý trong trường hợp công tơ không cấu hình 
                    vi_tri = strmessagesAnalyze.IndexOf("DD"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
                    if (vi_tri < 0)
                    {
                        //MessageBox.Show("Phân tích Q1 Q2 Q3 Q4  lỗi analyzeLoadProfile");
                        return "001-04";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
                    if (strmessagesAnalyze.Substring(0, 2) == "DD")
                    {
                        // Khong xu ly gì cả vì dữ liệu không cài đặt 
                        // Set thông tin =0 
                    }
                    else
                    {
                        // Nếu lấy Q thì sẽ chèn thông số vào đây 
                    }
                    vi_tri = strmessagesAnalyze.IndexOf("DD"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
                    if (vi_tri < 0)
                    {
                        //MessageBox.Show("Phân tích PMAX,QMAX  lỗi analyzeLoadProfile");
                        return "001-05";
                    }
                    strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
                    if (strmessagesAnalyze.Substring(0, 2) == "DD")
                    {
                        // Khong xu ly gì cả vì dữ liệu không cài đặt 
                        // Set thông tin =0 
                    }
                    else
                    {
                        // sẽ chèn thông số vào đây 
                        // 3 byte Pmax, 3 byte Q max
                        tempdata = strmessagesAnalyze.Substring(0, 6);
                        strmessagesAnalyze = strmessagesAnalyze.Substring(6, strmessagesAnalyze.Length - 6);
                        tempdata = ToStringValue(tempdata);
                        // Kiểm tra trạng thái P âm hay P dương để thực hiện truyền vào Giao hay nhận
                        string checkdau = "";
                        checkdau = tempdata.Substring(0, 2);
                        if (checkdau == "80")
                        {
                            tempdata = tempdata.Substring(2, tempdata.Length - 2);
                            // obj.P_MAX = double.Parse(tempdata) / 1000;
                            obj.P_EXPORTKW = double.Parse(tempdata) / 1000; // Vô công giao 
                        }
                        else
                        {
                            obj.P_IMPORTKW = double.Parse(tempdata) / 1000;

                        }
                        // Đoạn này lấy Q 
                        tempdata = strmessagesAnalyze.Substring(0, 6);
                        strmessagesAnalyze = strmessagesAnalyze.Substring(6, strmessagesAnalyze.Length - 6);
                        tempdata = ToStringValue(tempdata, 3);
                        if (tempdata == "ffffff")
                        {


                        }
                        else
                        {
                            checkdau = tempdata.Substring(0, 2);
                            if (checkdau == "80")
                            {
                                tempdata = tempdata.Substring(2, tempdata.Length - 2);
                                // obj.P_MAX = double.Parse(tempdata) / 1000;
                                obj.P_C1 = double.Parse(tempdata) / 1000; // Vô công giao 
                            }
                            else
                            {
                                obj.P_C2 = double.Parse(tempdata) / 1000;

                            }
                        }


                    }

                    //strvalue = strMgs.Substring(vi_tri + 6, 8);
                    //strvalue = ToStringValue(strvalue);
                    //obj.P_IMPORTKW = double.Parse(strvalue) / 1000;   //IMPORTKW P_IMPORTKW Chỉ số công suất Tổng giao

                    //strvalue = strMgs.Substring(vi_tri + 14, 8);
                    //strvalue = ToStringValue(strvalue);
                    //obj.P_EXPORTKW = double.Parse(strvalue) / 1000; //  //EXPORTKW    Chỉ số công suất Tổng nhận
                    //strvalue = strMgs.Substring(vi_tri + 22, 8);
                    //strvalue = ToStringValue(strvalue);
                    //obj.P_C1 = double.Parse(strvalue) / 1000; // Vô công giao 
                    //strvalue = strMgs.Substring(vi_tri + 30, 8);
                    //strvalue = ToStringValue(strvalue);
                    //obj.P_C2 = double.Parse(strvalue) / 1000;  // Vô công nhận
                    //vi_tri = tempdata.IndexOf("DDDD");
                    //if (vi_tri < 0)
                    //{
                    //    MessageBox.Show("Chuỗi không đúng");
                    //}
                    //strvalue = tempdata.Substring(vi_tri + 4, 6);
                    //strvalue = ToStringValue(strvalue);
                    //obj.decCongsuat = double.Parse(strvalue) / 10000; // P 
                    //strvalue = tempdata.Substring(vi_tri + 10, 6);
                    //strvalue = ToStringValue(strvalue);
                    //obj.P_Q1 = double.Parse(strvalue) / 10000; // Q

                }
                else
                {
                    return "007"; // frame messsage loi 
                }
            return "";

        }



        public string analyzeLoadProfile_auto(string strMgs, ref dynamic obj)
        {
            // Kiểm tra định dạng message 
            string startmesssage = "68";
            string endMessage = "16";
            string tempdata = "";
            int ilength = 0;
            string malenh = "";
            string strCheckLoaiMsg = "";
            string strmessagesAnalyze = "";
            string strvalue = "";

            string templength = "";
            string data = "";
            double declength = 0;
            string strngay = "";
            string strStartdate = "";
            string strEnddate = "";
            DateTime? _datetimeLoadpro = null;

            //MA_DIEMDO Mã điểm đo
            //NGAY Ngày
            //STARTDATE Thời gian bắt đầu
            //ENDDATE Thời gian kết thúc
            //IMPORTKW Chỉ số công suất Tổng giao
            //C1  Chỉ số vô công giao
            //EXPORTKW    Chỉ số công suất Tổng nhận
            //C2 Chỉ số vô công nhận
            //FLAG Trạng thái dữ liệu
            strmessagesAnalyze = strMgs;
            obj.LOAD_P_IMPORTKW = 0;
            obj.LOAD_P_EXPORTKW = 0;
            obj.LOAD_P_C1 = 0;
            obj.LOAD_P_C2 = 0;
            obj.LOAD_P_CONG_CS = 0;
            obj.LOAD_P_TRU_CS = 0;
            obj.LOAD_Q_TRU_CS = 0;
            obj.LOAD_Q_CONG_CS = 0;
            int vi_tri = 0;
            //34333339-- > DI load profile
            //Đánh dấu vị trí bắt đầu ngày -- Trương họp auto report A0A0
            malenh = strMgs.Substring(0, 4);
            obj.P_NGAY = new DateTime(1000, 1, 1);
            if (malenh == "A0A0")
            {
                strngay = strmessagesAnalyze.Substring(6, 10);
                _datetimeLoadpro = DateTime.ParseExact(strngay, "mmHHddMMyy", null);
                obj.P_NGAY = _datetimeLoadpro;
            }
            // strngay = ToStringDate(strngay, ref strStartdate, ref strEnddate);

            // Xử lý đoạn lấy UA, UB, UC, IA, IB,IC

            vi_tri = strmessagesAnalyze.IndexOf("AA"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
            if (vi_tri < 0)
            {
                //  MessageBox.Show("Chuỗi không đúng analyzeLoadProfile");
                return "001-01";
            }
            strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);

            vi_tri = strmessagesAnalyze.IndexOf("AA"); // DD Cosfi 
            if (vi_tri < 0)
            {
                return "001-02";
            }
            strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
            vi_tri = strmessagesAnalyze.IndexOf("AA"); // DD Điện năng (Chỉ số P giao P nhận Q giao Q nhận )
            if (vi_tri < 0)
            {
                return "001-03";
            }

            strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);


            // strmessagesAnalyze.Substring(0, 8);

            // P giao 
            String checktempdata = strmessagesAnalyze.Substring(0, 4);
            if (checktempdata == "AAAA")
            {

            }
            else
            {
                tempdata = ToStringValue_nomal(tempdata);
                obj.LOAD_P_CONG_CS = double.Parse(tempdata) / 1000;   //IMPORTKW P_IMPORTKW Chỉ số công suất Tổng giao
                strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                // P nhận
                tempdata = strmessagesAnalyze.Substring(0, 8);
                tempdata = ToStringValue_nomal(tempdata);
                strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                obj.LOAD_P_TRU_CS = double.Parse(tempdata) / 1000;
                // Q Nhận
                tempdata = strmessagesAnalyze.Substring(0, 8);
                tempdata = ToStringValue_nomal(tempdata);
                strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                obj.LOAD_Q_CONG_CS = double.Parse(tempdata) / 1000;
                // Q Giao 
                tempdata = strmessagesAnalyze.Substring(0, 8);
                tempdata = ToStringValue_nomal(tempdata);
                strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                obj.LOAD_Q_TRU_CS = double.Parse(tempdata) / 1000;
            }



            // Tìm đến vị trí tiếp theo để lấy PMAX
            // Xử lý trong trường hợp công tơ không cấu hình 
            vi_tri = strmessagesAnalyze.IndexOf("AA"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
            if (vi_tri < 0)
            {
                //MessageBox.Show("Phân tích Q1 Q2 Q3 Q4  lỗi analyzeLoadProfile");
                return "001-04";
            }
            strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
            if (strmessagesAnalyze.Substring(0, 2) == "AA")
            {
                // Khong xu ly gì cả vì dữ liệu không cài đặt 
                // Set thông tin =0 
            }
            else
            {
                // Nếu lấy Q thì sẽ chèn thông số vào đây 
            }
            vi_tri = strmessagesAnalyze.IndexOf("AA"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
            if (vi_tri < 0)
            {
                //MessageBox.Show("Phân tích PMAX,QMAX  lỗi analyzeLoadProfile");
                return "001-05";
            }
            strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
            if (strmessagesAnalyze.Substring(0, 2) == "AA")
            {
                // Khong xu ly gì cả vì dữ liệu không cài đặt 
                // Set thông tin =0 
            }
            else
            {
                // sẽ chèn thông số vào đây 
                // 3 byte Pmax, 3 byte Q max
                tempdata = strmessagesAnalyze.Substring(0, 6);
                string checkdau = "";
                if (tempdata == "EEEEEE")
                {


                }
                else
                {
                    strmessagesAnalyze = strmessagesAnalyze.Substring(6, strmessagesAnalyze.Length - 6);

                    tempdata = ToStringValue_nomal(tempdata, 3);
                    // Kiểm tra trạng thái P âm hay P dương để thực hiện truyền vào Giao hay nhận


                    checkdau = tempdata.Substring(0, 2);
                    if (checkdau == "80")
                    {
                        tempdata = tempdata.Substring(2, tempdata.Length - 2);
                        // obj.P_MAX = double.Parse(tempdata) / 1000;
                        obj.LOAD_P_EXPORTKW = double.Parse(tempdata) / 10000; // Vô công giao 
                    }
                    else
                    {
                        obj.LOAD_P_IMPORTKW = double.Parse(tempdata) / 10000;

                    }

                }

                // Đoạn này lấy Q 
                tempdata = strmessagesAnalyze.Substring(0, 6);


                if (tempdata == "EEEEEE")
                {


                }
                else
                {
                    tempdata = ToStringValue_nomal(tempdata, 3);
                    strmessagesAnalyze = strmessagesAnalyze.Substring(6, strmessagesAnalyze.Length - 6);
                    checkdau = tempdata.Substring(0, 2);
                    if (checkdau == "80")
                    {
                        tempdata = tempdata.Substring(2, tempdata.Length - 2);
                        // obj.P_MAX = double.Parse(tempdata) / 1000;
                        obj.LOAD_P_C1 = double.Parse(tempdata) / 10000; // Vô công giao 
                    }
                    else
                    {
                        obj.LOAD_P_C2 = double.Parse(tempdata) / 10000;

                    }
                }


            }


            return "";

        }

        public string analyzeLoadProfile_auto(string strMgs, ref DataBLock obj)
        {
            // Kiểm tra định dạng message 
            string startmesssage = "68";
            string endMessage = "16";
            string tempdata = "";
            int ilength = 0;
            string malenh = "";
            string strCheckLoaiMsg = "";
            string strmessagesAnalyze = "";
            string strvalue = "";

            string templength = "";
            string data = "";
            double declength = 0;
            string strngay = "";
            string strStartdate = "";
            string strEnddate = "";
            DateTime? _datetimeLoadpro = null;

            //MA_DIEMDO Mã điểm đo
            //NGAY Ngày
            //STARTDATE Thời gian bắt đầu
            //ENDDATE Thời gian kết thúc
            //IMPORTKW Chỉ số công suất Tổng giao
            //C1  Chỉ số vô công giao
            //EXPORTKW    Chỉ số công suất Tổng nhận
            //C2 Chỉ số vô công nhận
            //FLAG Trạng thái dữ liệu
            strmessagesAnalyze = strMgs;
            obj.LOAD_P_IMPORTKW = null;
            obj.LOAD_P_EXPORTKW = null;
            obj.LOAD_P_C1 = 0;
            obj.LOAD_P_C2 = 0;
            obj.LOAD_P_CONG_CS = 0;
            obj.LOAD_P_TRU_CS = 0;
            obj.LOAD_Q_TRU_CS = 0;
            obj.LOAD_Q_CONG_CS = 0;
            int vi_tri = 0;
            //34333339-- > DI load profile
            //Đánh dấu vị trí bắt đầu ngày -- Trương họp auto report A0A0
            malenh = strMgs.Substring(0, 4);
            obj.P_NGAY = new DateTime(1000, 1, 1);
            if (malenh == "A0A0")
            {
                strngay = strmessagesAnalyze.Substring(6, 10);
                _datetimeLoadpro = DateTime.ParseExact(strngay, "mmHHddMMyy", null);
                obj.P_NGAY = _datetimeLoadpro;
            }
            // strngay = ToStringDate(strngay, ref strStartdate, ref strEnddate);

            // Xử lý đoạn lấy UA, UB, UC, IA, IB,IC

            vi_tri = strmessagesAnalyze.IndexOf("AA"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
            if (vi_tri < 0)
            {
                //  MessageBox.Show("Chuỗi không đúng analyzeLoadProfile");
                return "001-01";
            }
            strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);

            vi_tri = strmessagesAnalyze.IndexOf("AA"); // DD Cosfi 
            if (vi_tri < 0)
            {
                return "001-02";
            }
            strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
            vi_tri = strmessagesAnalyze.IndexOf("AA"); // DD Điện năng (Chỉ số P giao P nhận Q giao Q nhận )
            if (vi_tri < 0)
            {
                return "001-03";
            }

            strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);


            // strmessagesAnalyze.Substring(0, 8);

            // P giao 
            tempdata = strmessagesAnalyze.Substring(0, 8);
            if (tempdata.Substring(0, 4) == "AAAA")
            {

            }
            else
            {
                tempdata = ToStringValue_nomal(tempdata);
                obj.LOAD_P_CONG_CS = double.Parse(tempdata) / 100;   //IMPORTKW P_IMPORTKW Chỉ số công suất Tổng giao
                strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                // P nhận
                tempdata = strmessagesAnalyze.Substring(0, 8);
                tempdata = ToStringValue_nomal(tempdata);
                strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                obj.LOAD_P_TRU_CS = double.Parse(tempdata) / 100;
                // Q Nhận
                tempdata = strmessagesAnalyze.Substring(0, 8);
                tempdata = ToStringValue_nomal(tempdata);
                strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                obj.LOAD_Q_CONG_CS = double.Parse(tempdata) / 100;
                // Q Giao 
                tempdata = strmessagesAnalyze.Substring(0, 8);
                tempdata = ToStringValue_nomal(tempdata);
                strmessagesAnalyze = strmessagesAnalyze.Substring(8, strmessagesAnalyze.Length - 8);
                obj.LOAD_Q_TRU_CS = double.Parse(tempdata) / 100;
            }



            // Tìm đến vị trí tiếp theo để lấy PMAX
            // Xử lý trong trường hợp công tơ không cấu hình 
            vi_tri = strmessagesAnalyze.IndexOf("AA"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
            if (vi_tri < 0)
            {
                //MessageBox.Show("Phân tích Q1 Q2 Q3 Q4  lỗi analyzeLoadProfile");
                return "001-04";
            }
            strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
            if (strmessagesAnalyze.Substring(0, 2) == "AA")
            {
                // Khong xu ly gì cả vì dữ liệu không cài đặt 
                // Set thông tin =0 
            }
            else
            {
                // Nếu lấy Q thì sẽ chèn thông số vào đây 
            }
            vi_tri = strmessagesAnalyze.IndexOf("AA"); // DD Công suất tổng , Công suất pha A, Công suất pha B , công suất pha C, Công suât tổng Q, Công suất Q pha A, Công suất Q pha B, Công suất Q pha C  
            if (vi_tri < 0)
            {
                //MessageBox.Show("Phân tích PMAX,QMAX  lỗi analyzeLoadProfile");
                return "001-05";
            }
            strmessagesAnalyze = strmessagesAnalyze.Substring(vi_tri + 2, strmessagesAnalyze.Length - vi_tri - 2);
            if (strmessagesAnalyze.Substring(0, 2) == "AA")
            {
                // Khong xu ly gì cả vì dữ liệu không cài đặt 
                // Set thông tin =0 
            }
            else
            {
                // sẽ chèn thông số vào đây 
                // 3 byte Pmax, 3 byte Q max
                tempdata = strmessagesAnalyze.Substring(0, 6);
                string checkdau = "";
                if (tempdata == "EEEEEE")
                {


                }
                else
                {
                    strmessagesAnalyze = strmessagesAnalyze.Substring(6, strmessagesAnalyze.Length - 6);

                    tempdata = ToStringValue_nomal(tempdata, 3);
                    // Kiểm tra trạng thái P âm hay P dương để thực hiện truyền vào Giao hay nhận


                    checkdau = tempdata.Substring(0, 2);
                    if (checkdau == "80")
                    {
                        tempdata = tempdata.Substring(2, tempdata.Length - 2);
                        // obj.P_MAX = double.Parse(tempdata) / 1000;
                        obj.LOAD_P_EXPORTKW = double.Parse(tempdata) / 10000; // Vô công giao 
                    }
                    else
                    {
                        obj.LOAD_P_IMPORTKW = double.Parse(tempdata) / 10000;

                    }

                }

                // Đoạn này lấy Q 
                tempdata = strmessagesAnalyze.Substring(0, 6);


                if (tempdata == "EEEEEE")
                {


                }
                else
                {
                    tempdata = ToStringValue_nomal(tempdata, 3);
                    strmessagesAnalyze = strmessagesAnalyze.Substring(6, strmessagesAnalyze.Length - 6);
                    checkdau = tempdata.Substring(0, 2);
                    if (checkdau == "80")
                    {
                        tempdata = tempdata.Substring(2, tempdata.Length - 2);
                        // obj.P_MAX = double.Parse(tempdata) / 1000;
                        obj.LOAD_P_C1 = double.Parse(tempdata) / 10000; // Vô công giao 
                    }
                    else
                    {
                        obj.LOAD_P_C2 = double.Parse(tempdata) / 10000;

                    }
                }


            }


            return "";

        }
        #endregion



        public void analyze3GHUUHONG_LoadProfile(string[] value_item, ref dynamic objLoad, ref dynamic obj, ref dynamic objHIS, ref dynamic objCUR, ref string ketqua)
        {
            string strTempdata;
            string strCheckTypedata;
            double doubleValue = 0;
            double douValue1 = 0;
            double douValue2 = 0;
            double douValue3 = 0;
            DateTime datetimetemp = DateTime.Now;

            try
            {

                for (int i = 4; i < value_item.Length; i++)
                {
                    string[] row_item = value_item[i].Split(new char[] { '\t' });
                    {
                        // Kiểm tra loại giá trị dữ liệu
                        if (row_item[0].IndexOf(":") > 0)
                            strCheckTypedata = row_item[0].Substring(0, row_item[0].IndexOf(":"));
                        else
                            strCheckTypedata = "";



                        switch (strCheckTypedata)
                        {
                            case "LOADPROFILE1": //-75
                                strTempdata = seprarateString(row_item);
                                // ketqua = analyzeLoadProfile75(strTempdata, ref objLoad);
                                ketqua = analyzeLoadProfile(strTempdata, ref objLoad);
                                break;
                            case "LOADPROFILE2": //-45
                                strTempdata = seprarateString(row_item);
                                ketqua = analyzeLoadProfile(strTempdata, ref objLoad);
                                break;

                            /*
                        case "CSC_HCT":
                                     strTempdata = seprarateString(row_item);
                                     analyzeValue(strTempdata, ref doubleValue,ref douValue1, ref douValue2 , ref douValue3);
                                     objHIS.P_IMPORTKWH = doubleValue;
                                     objHIS.P_IMPBT = douValue1;
                                     objHIS.P_IMPCD = douValue2;
                                     objHIS.P_IMPTD = douValue3;


                                     break;
                                 case "CSC_HCN":
                                     strTempdata = seprarateString(row_item);
                                     analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                     objHIS.P_EXPORTKWH = doubleValue;
                                     objHIS.P_EXPBT = douValue1;
                                     objHIS.P_EXPCD = douValue2;
                                     objHIS.P_EXPTD = douValue3;


                                     break;
                                 case "CSC_VCT":
                                     strTempdata = seprarateString(row_item);
                                     analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                     objHIS.P_C1 = doubleValue;


                                                                break;
                                                  case "CSC_VCN":
                                                     strTempdata = seprarateString(row_item);
                                                     analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                                     objHIS.P_C2 = doubleValue;

                                                     break;
                                                 case "PMAX":
                                                     //strTempdata = seprarateString(row_item);
                                                     //analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                                     //objCUR.P_C2 = doubleValue;
                                                     break;
                                                 case "HCT":
                                                     strTempdata = seprarateString(row_item);
                                                     analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                                     objCUR.P_IMPORTKWH = doubleValue;
                                                     objCUR.P_IMPBT = douValue1;
                                                     objCUR.P_IMPCD = douValue2;
                                                     objCUR.P_IMPTD = douValue3;
                                                     break;

                                                 case "HCN":
                                                     strTempdata = seprarateString(row_item);
                                                     analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                                     objCUR.P_EXPORTKWH = doubleValue;
                                                     objCUR.P_EXPBT = douValue1;
                                                     objCUR.P_EXPCD = douValue2;
                                                     objCUR.P_EXPTD = douValue3;
                                                     break;
                                                 case "VCT":
                                                     strTempdata = seprarateString(row_item);
                                                     analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                                     objCUR.P_C1 = doubleValue;
                                                     break;
                                                 case "VCN":
                                                     strTempdata = seprarateString(row_item);
                                                     analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                                     objHIS.P_C2 = doubleValue;
                                                     break;
                                                 case "DIENAP":
                                                     strTempdata = seprarateString(row_item);
                                                     analyzeValue_Dienap(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                                     obj.P_V_A = doubleValue;
                                                     obj.P_V_B = douValue1;
                                                     obj.P_V_C = douValue2;
                                                     break;
                                                 case "DONGDIEN":
                                                     strTempdata = seprarateString(row_item);
                                                     analyzeValue_Dongdien(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                                     obj.P_A_A = doubleValue;
                                                     obj.P_A_B = douValue1;
                                                     obj.P_A_C = douValue2;
                                                     break;
                                                 case "P_HC":
                                                     strTempdata = seprarateString(row_item);
                                                     analyzeValue_P(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                                     objCUR.P_EXPORTKWH = doubleValue;
                                                     objCUR.P_EXPBT = doubleValue;
                                                     objCUR.P_EXPCD = doubleValue;
                                                     objCUR.P_EXPTD = doubleValue;
                                                     break;


                        case "P_VC":
                            strTempdata = seprarateString(row_item);
                            analyzeValue_P(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                            objCUR.P_EXPORTKWH = doubleValue;
                            objCUR.P_EXPBT = doubleValue;
                            objCUR.P_EXPCD = doubleValue;
                            objCUR.P_EXPTD = doubleValue;
                            break;
                            */
                            case "GIO_CT":
                                //strTempdata = seprarateString(row_item);
                                //analyzeValue_GIO(strTempdata, ref datetimetemp);
                                //objCUR.P_CLOCK_METER = datetimetemp;

                                break;
                            default:
                                break;
                        }
                        //obj.P_SERIALID = row_item[0];
                        //obj.P_DATETIME = Convert.ToDateTime(row_item[8]);

                        //obj.P_ROUTING_PATH = row_item[3];
                        //obj.P_RF_RSSI = int.Parse(row_item[4]);
                        //obj.P_RF_LEVER = int.Parse(row_item[5]);
                        //obj.P_RF_POWER_MAX = int.Parse(row_item[6]);
                        //obj.P_READ_PATH_SUCCESS = row_item[7];
                        //string[] item_para = row_item[2].Split(new char[] { '_' });
                        //if (item_para.Count() > 3)
                        //{
                        //    obj.P_IMPORTKWH = double.Parse(item_para[0].Replace(",", "."));
                        //    obj.P_EXPORTKWH = double.Parse(item_para[1].Replace(",", "."));
                        //    obj.P_VOLPHASE = double.Parse(item_para[3].Replace(",", "."));
                        //    obj.P_CURPHASE = double.Parse(item_para[2].Replace(",", "."));
                        //    obj.P_ACTIVE_POWER_PHASE = double.Parse(item_para[4].Replace(",", "."));
                        //}
                        //else
                        //{
                        //    obj.P_IMPORTKWH = double.Parse(item_para[0].Replace(",", "."));
                        //}

                    }
                }
            }
            catch (Exception ex)
            {
                //   MessageBox.Show(ex.ToString()+ objLoad.P_MA_DIEMDO);
                //  File.WriteAllText("C:\\LOGMDMS\\" + objLoad.P_MA_DIEMDO  + ".txt", objLoad.P_MA_DIEMDO + "\r\t");
                //using (StreamWriter sw = File.AppendText("C:\\LOGMDMS\\log_2.txt"))
                //{
                //    sw.WriteLine(objLoad.P_MA_DIEMDO + "\r\t");

                //}
                ketqua = "007";

            }

        }






        public void analyze3GHUUHONG_CSC(string[] value_item, ref dynamic objLoad, ref dynamic obj, ref dynamic objHIS, ref dynamic objCUR, ref string ketqua)
        {
            string strTempdata;
            string strCheckTypedata;
            double doubleValue = 0;
            double douValue1 = 0;
            double douValue2 = 0;
            double douValue3 = 0;
            DateTime datetimetemp = DateTime.Now;

            try
            {

                for (int i = 4; i < value_item.Length; i++)
                {
                    string[] row_item = value_item[i].Split(new char[] { '\t' });
                    {
                        // Kiểm tra loại giá trị dữ liệu
                        if (row_item[0].IndexOf(":") > 0)
                            strCheckTypedata = row_item[0].Substring(0, row_item[0].IndexOf(":"));
                        else
                            strCheckTypedata = "";



                        switch (strCheckTypedata)
                        {
                            /* case "LOADPROFILE":
                                strTempdata = seprarateString(row_item);
                                analyzeLoadProfile(strTempdata, ref objLoad);
                                break;
                               */

                            case "CSC_HCT":
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objHIS.P_IMPORTKWH = doubleValue;
                                objHIS.P_IMPBT = douValue1;
                                objHIS.P_IMPCD = douValue2;
                                objHIS.P_IMPTD = douValue3;


                                break;
                            case "CSC_HCN":
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objHIS.P_EXPORTKWH = doubleValue;
                                objHIS.P_EXPBT = douValue1;
                                objHIS.P_EXPCD = douValue2;
                                objHIS.P_EXPTD = douValue3;


                                break;
                            case "CSC_VCT":
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objHIS.P_C1 = doubleValue;


                                break;
                            case "CSC_VCN":
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objHIS.P_C2 = doubleValue;

                                break;
                            case "PMAX":
                                //strTempdata = seprarateString(row_item);
                                //analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                //objCUR.P_C2 = doubleValue;
                                break;
                            /*   case "HCT":
                                   strTempdata = seprarateString(row_item);
                                   analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                   objCUR.P_IMPORTKWH = doubleValue;
                                   objCUR.P_IMPBT = douValue1;
                                   objCUR.P_IMPCD = douValue2;
                                   objCUR.P_IMPTD = douValue3;
                                   break;

                               case "HCN":
                                   strTempdata = seprarateString(row_item);
                                   analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                   objCUR.P_EXPORTKWH = doubleValue;
                                   objCUR.P_EXPBT = douValue1;
                                   objCUR.P_EXPCD = douValue2;
                                   objCUR.P_EXPTD = douValue3;
                                   break;
                               case "VCT":
                                   strTempdata = seprarateString(row_item);
                                   analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                   objCUR.P_C1 = doubleValue;
                                   break;
                               case "VCN":
                                   strTempdata = seprarateString(row_item);
                                   analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                   objHIS.P_C2 = doubleValue;
                                   break;
                               case "DIENAP":
                                   strTempdata = seprarateString(row_item);
                                   analyzeValue_Dienap(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                   obj.P_V_A = doubleValue;
                                   obj.P_V_B = douValue1;
                                   obj.P_V_C = douValue2;
                                   break;
                               case "DONGDIEN":
                                   strTempdata = seprarateString(row_item);
                                   analyzeValue_Dongdien(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                   obj.P_A_A = doubleValue;
                                   obj.P_A_B = douValue1;
                                   obj.P_A_C = douValue2;
                                   break;
                               case "P_HC":
                                   strTempdata = seprarateString(row_item);
                                   analyzeValue_P(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                   objCUR.P_EXPORTKWH = doubleValue;
                                   objCUR.P_EXPBT = doubleValue;
                                   objCUR.P_EXPCD = doubleValue;
                                   objCUR.P_EXPTD = doubleValue;
                                   break;


                               case "P_VC":
                                   strTempdata = seprarateString(row_item);
                                   analyzeValue_P(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                   objCUR.P_EXPORTKWH = doubleValue;
                                   objCUR.P_EXPBT = doubleValue;
                                   objCUR.P_EXPCD = doubleValue;
                                   objCUR.P_EXPTD = doubleValue;
                                   break;
                                   */
                            case "GIO_CT":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objCUR.P_CLOCK_METER = datetimetemp;

                                break;
                            default:
                                break;
                        }


                    }
                }
            }
            catch (Exception ex)
            {
                //    MessageBox.Show(ex.ToString() + objHIS.P_MA_DIEMDO);
                //  File.WriteAllText("C:\\LOGMDMS\\"+ objHIS.P_MA_DIEMDO + ".txt", objHIS.P_MA_DIEMDO + "\r\t");
                //using (StreamWriter sw = File.AppendText("C:\\LOGMDMS\\log.txt"))
                //{
                //    sw.WriteLine(objHIS.P_MA_DIEMDO + "\r\t");

                //}

                ketqua = "009"; // lỗi không xác định 
            }

        }
        public void analyze3GHUUHONG_TSVH(string[] value_item, ref dynamic objLoad, ref dynamic obj, ref dynamic objHIS, ref dynamic objCUR, ref string ketqua)
        {
            string strTempdata;
            string strCheckTypedata;
            double doubleValue = 0;
            double douValue1 = 0;
            double douValue2 = 0;
            double douValue3 = 0;
            DateTime datetimetemp = DateTime.Now;

            try
            {

                for (int i = 4; i < value_item.Length; i++)
                {
                    string[] row_item = value_item[i].Split(new char[] { '\t' });
                    {
                        // Kiểm tra loại giá trị dữ liệu
                        if (row_item[0].IndexOf(":") > 0)
                            strCheckTypedata = row_item[0].Substring(0, row_item[0].IndexOf(":"));
                        else
                            strCheckTypedata = "";



                        switch (strCheckTypedata)
                        {
                            /*
                             case "LOADPROFILE":
                                strTempdata = seprarateString(row_item);
                                analyzeLoadProfile(strTempdata, ref objLoad);
                                break;
                            
                            
                             case "CSC_HCT":
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objHIS.P_IMPORTKWH = doubleValue;
                                objHIS.P_IMPBT = douValue1;
                                objHIS.P_IMPCD = douValue2;
                                objHIS.P_IMPTD = douValue3;


                                break;
                            case "CSC_HCN":
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objHIS.P_EXPORTKWH = doubleValue;
                                objHIS.P_EXPBT = douValue1;
                                objHIS.P_EXPCD = douValue2;
                                objHIS.P_EXPTD = douValue3;


                                break;
                            case "CSC_VCT":
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objHIS.P_C1 = doubleValue;


                                break;
                            case "CSC_VCN":
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objHIS.P_C2 = doubleValue;

                                break;
                                 case "VCN":
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objHIS.P_C2 = doubleValue;
                                break;
                                 */
                            case "PMAX":
                                //strTempdata = seprarateString(row_item);
                                //analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                //objCUR.P_C2 = doubleValue;
                                break;
                            case "HCT": // Chỉ số hữu công thuân
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objCUR.P_IMPORTKWH = doubleValue;
                                objCUR.P_IMPBT = douValue1;
                                objCUR.P_IMPCD = douValue2;
                                objCUR.P_IMPTD = douValue3;
                                break;

                            case "HCN": // Chỉ số hữu cộng ngươc
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objCUR.P_EXPORTKWH = doubleValue;
                                objCUR.P_EXPBT = douValue1;
                                objCUR.P_EXPCD = douValue2;
                                objCUR.P_EXPTD = douValue3;
                                break;
                            case "VCT":
                                strTempdata = seprarateString(row_item);
                                analyzeValue(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objCUR.P_C1 = doubleValue;
                                break;

                            case "DIENAP":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_Dienap(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objCUR.P_V_A = doubleValue;
                                objCUR.P_V_B = douValue1;
                                objCUR.P_V_C = douValue2;
                                break;
                            case "DONGDIEN":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_Dongdien(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objCUR.P_A_A = doubleValue;
                                objCUR.P_A_B = douValue1;
                                objCUR.P_A_C = douValue2;
                                break;
                            case "P_HC": // Công suất hữu công 
                                strTempdata = seprarateString(row_item);
                                analyzeValue_P(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objCUR.P_AP_T = doubleValue;
                                objCUR.P_AP_A = douValue1;
                                objCUR.P_AP_B = douValue2;
                                objCUR.P_AP_C = douValue3;
                                break;


                            case "P_VC": // Công suất vô công 
                                strTempdata = seprarateString(row_item);
                                analyzeValue_P(strTempdata, ref doubleValue, ref douValue1, ref douValue2, ref douValue3);
                                objCUR.P_RP_T = doubleValue;
                                objCUR.P_RP_A = douValue1;
                                objCUR.P_RP_B = douValue2;
                                objCUR.P_RP_C = douValue3;
                                break;
                            case "GIO_CT":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objCUR.P_CLOCK_METER = datetimetemp;

                                break;
                            default:
                                break;
                        }


                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString() + objCUR.P_MA_DIEMDO);
                //   File.WriteAllText("C:\\LOGMDMS\\" + objCUR.P_MA_DIEMDO  + "log.txt", objCUR.P_MA_DIEMDO + "\r\t");
                //using (StreamWriter sw = File.AppendText("C:\\LOGMDMS\\log.txt"))
                //{
                //    sw.WriteLine(objLoad.P_MA_DIEMDO + "\r\t");

                //}
                ketqua = "009";
            }

        }



        public void analyze3GHUUHONG_EVENT(string[] value_item, ref dynamic objEvent, ref string ketqua)
        {
            string strTempdata;
            string strCheckTypedata;
            double doubleValue = 0;
            double douValue1 = 0;
            double douValue2 = 0;
            double douValue3 = 0;
            DateTime datetimetemp = DateTime.Now;

            try
            {

                for (int i = 4; i < value_item.Length; i++)
                {
                    string[] row_item = value_item[i].Split(new char[] { '\t' });
                    {
                        // Kiểm tra loại giá trị dữ liệu
                        if (row_item[0].IndexOf(":") > 0)
                            strCheckTypedata = row_item[0].Substring(0, row_item[0].IndexOf(":"));
                        else
                            strCheckTypedata = "";



                        switch (strCheckTypedata)
                        {
                            /*
                             case "LOADPROFILE":
                                strTempdata = seprarateString(row_item);
                                analyzeLoadProfile(strTempdata, ref objLoad);
                                break;
                             */

                            case "LAN_MATDIEN":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.p_SOLAN = doubleValue;

                                break;
                            case "THOIGIAN_MATDIEN":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIAN = datetimetemp;




                                break;
                            case "LAN_MATDIEN_A":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANA = doubleValue;


                                break;
                            case "LAN_MATDIEN_B":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANB = doubleValue;

                                break;
                            case "LAN_MATDIEN_C":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANC = doubleValue; break;
                                break;
                            case "THOIGIAN_MATDIEN_A":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANA = datetimetemp;
                                break;
                            case "THOIGIAN_MATDIEN_B": // Chỉ số hữu công thuân
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANB = datetimetemp;
                                break;

                            case "THOIGIAN_MATDIEN_C": // Chỉ số hữu cộng ngươc
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANC = datetimetemp;
                                break;
                            case "LAN_DUOIAP_A":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANDUOIAPA = doubleValue; break;
                                break;

                            case "LAN_DUOIAP_B":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANDUOIAPB = doubleValue; break;
                                break;
                            case "LAN_DUOIAP_C":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANDUOIAPC = doubleValue; break;
                                break;
                            case "THOIGIAN_DUOIAP_A": // Công suất hữu công 
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANDUOIAPA = datetimetemp;
                                break;


                            case "THOIGIAN_DUOIAP_B": // Công suất vô công 
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANDUOIAPB = datetimetemp;
                                break;
                            case "THOIGIAN_DUOIAP_C":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANDUOIAPC = datetimetemp;

                                break;
                            case "LAN_QUAAP_A":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANQUAAPA = doubleValue; break;

                                break;
                            case "LAN_QUAAP_B":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANQUAAPB = doubleValue; break;

                                break;
                            case "LAN_QUAAP_C":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANQUAAPC = doubleValue; break;

                                break;
                            case "THOIGIAN_QUAAP_A":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANQUAAPA = datetimetemp;

                                break;
                            case "THOIGIAN_QUAAP_B":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANQUAAPB = datetimetemp;

                                break;
                            case "THOIGIAN_QUAAP_C":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANQUAAPC = datetimetemp;

                                break;
                            case "LAN_QUADONG_A":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANQUADONGA = doubleValue; break;

                                break;
                            case "LAN_QUADONG_B":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANQUADONGB = doubleValue; break;


                                break;
                            case "LAN_QUADONG_C":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_3byte(strTempdata, ref doubleValue);
                                objEvent.P_SOLANQUADONGC = doubleValue; break;


                                break;
                            case "THOIGIAN_QUADONG_A":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANQUADONGA = datetimetemp;

                                break;
                            case "THOIGIAN_QUADONG_B":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANQUADONGB = datetimetemp;

                                break;
                            case "THOIGIAN_QUADONG_C":
                                strTempdata = seprarateString(row_item);
                                analyzeValue_GIO(strTempdata, ref datetimetemp);
                                objEvent.p_THOIGIANQUADONGC = datetimetemp;

                                break;
                            default:
                                break;
                        }


                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString() + objCUR.P_MA_DIEMDO);
                //   File.WriteAllText("C:\\LOGMDMS\\" + objCUR.P_MA_DIEMDO  + "log.txt", objCUR.P_MA_DIEMDO + "\r\t");
                //using (StreamWriter sw = File.AppendText("C:\\LOGMDMS\\log.txt"))
                //{
                //    sw.WriteLine(objLoad.P_MA_DIEMDO + "\r\t");

                //}
                ketqua = "009";
            }

        }


        // kết thúc hàm 

    }
}