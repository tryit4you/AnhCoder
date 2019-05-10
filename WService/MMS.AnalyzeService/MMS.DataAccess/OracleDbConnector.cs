using System;
using System.Collections.Generic;
using System.Data.OracleClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MMS.DataAccess
{
    public class OracleDbConnector
    {
        private OracleConnection oracleConnection;
        private OracleConnectionStringBuilder stringBuilder;
        public OracleDbConnector(string dbUsername, string dbPassword, string dbServer)
        {
            stringBuilder = new OracleConnectionStringBuilder();
            stringBuilder.DataSource = dbServer;
            stringBuilder.UserID = dbUsername;
            stringBuilder.Password = dbPassword;
            oracleConnection = new OracleConnection();
            oracleConnection.ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.42.90)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=MDMSCT))); User Id = EVNHES; Password = EVNHES; ";

        }
        public bool UploadBulkData(List<CustomerDTO> bulkData)
        {
            bool returnValue = false;
            try
            {
                oracleConnection.Open();

                // var cmd = oracleConnection.CreateCommand();

                // cmd.CommandText = "SELECT NVL (A.ASSETID, ' '), NVL (Z.KIEU_CTO, ' '),A.SERIALNUM FROM A_ASSET A JOIN ZAG_MDMS_DIEMDO Z ON A.ASSETID = Z.OBJID";
                // cmd.CommandType = CommandType.Text;
                // OracleDataReader dr = cmd.ExecuteReader();
                // dr.Read();
                //string test = dr.GetString(0);




                string query = @"insert into EVNHES.customers ( surname, firstName, emailAddress) values (:surname, :firstName, :emailAddress)";

                using (var command = oracleConnection.CreateCommand())
                {
                    command.CommandText = query;
                    command.CommandType = CommandType.Text;
                    command.BindByName = true;
                    // In order to use ArrayBinding, the ArrayBindCount property
                    // of OracleCommand object must be set to the number of records to be inserted
                    command.ArrayBindCount = bulkData.Count;
                    command.Parameters.Add(":surname", OracleDbType.Varchar2, bulkData.Select(c => c.Surname).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":firstName", OracleDbType.Varchar2, bulkData.Select(c => c.FirstName).ToArray(), ParameterDirection.Input);
                    command.Parameters.Add(":emailAddress", OracleDbType.Varchar2, bulkData.Select(c => c.EmailAddress).ToArray(), ParameterDirection.Input);
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
                oracleConnection.Close();
            }
            return returnValue;
        }
    }
}
