using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MMS.AnalyzeFile
{
    public class General
    {
        public static string resPath = "";
        public static string connectString = "SERVER=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.42.90)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=MDMSCT)));User Id=EVNHES;Password=EVNHES;";

        public static string connectionstr = "SERVER=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.42.90)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=MDMSCT)));uid=EVNHES;pwd=EVNHES";
        public static int meterType = 1;
        public static int channelCount = 20;
        public static bool read_config()
        {
            try
            {
                DataSet ds = new DataSet();
                ds.ReadXml(Environment.CurrentDirectory + "\\Config.xml");
                DataRow row = ds.Tables[0].Rows[0];
                resPath = row["RES_PATH"].ToString();
                connectString = row["CONNECT_STRING"].ToString().Trim();
                int.TryParse(row["METER_TYPE"].ToString().Trim(), out meterType);
                int.TryParse(row["CHANNEL"].ToString().Trim(), out channelCount);

                return true;
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
                return false;
            }
        }
        private delegate void AppendRichTextBox_Callback(ref RichTextBox rtblog, string text, Color color);
        public static void AppendRichTextBox(ref RichTextBox rtblog, string text, Color color)
        {

            if (rtblog != null)
            {
                if (rtblog.InvokeRequired)
                {
                    AppendRichTextBox_Callback d = new AppendRichTextBox_Callback(AppendRichTextBox);
                    rtblog.Invoke(d, new object[] { rtblog, text, color });
                }
                else
                {
                    rtblog.SelectionColor = color;
                    rtblog.SelectedText = text + "\n";
                    rtblog.SelectionStart = rtblog.Text.Length;
                }
            }
        }
        public static void ListParms()
        {
            OracleConnection conn = new OracleConnection("SERVER=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.3.1.11)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=MDMSPMIS)));User Id=MDMS;Password=MDMS;");
            OracleCommand cmd = new OracleCommand("PKG_WEBSERVICE_MDMS.Get_DiemDo", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            conn.Open();
            OracleCommandBuilder.DeriveParameters(cmd);
            foreach (OracleParameter p in cmd.Parameters)
            {
                string acb = p.ParameterName + " " + p.OracleType + " " + p.Value + " " + p.Size;
                var dir = p.Direction;
                var dirf = p.Offset;
            }
        }

    }
}
