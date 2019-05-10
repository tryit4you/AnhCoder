using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConnectToOracle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OracleConnection connection = new OracleConnection();
            connection.ConnectionString = "Data Source=LOCALHOST:1521\\NEWDB; User Id = SYSTEM; Password=SYSTEM";           try
            {
                connection.Open();
                MessageBox.Show("Kết nối csdl thành công~");
                connection.Close();
            }catch(Exception ex)
            {
                MessageBox.Show("Lỗi kết nối :" +ex.Message);
            }
        }
    }
}
