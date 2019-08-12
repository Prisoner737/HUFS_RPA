using System;
using System.Text;
using MySql.Data.MySqlClient;

namespace RPA_SummerProj
{
    class MySqlManager : IDisposable
    {
        MySqlConnection conn;
        public MySqlManager(params string[] usrInfo)
        {
            //usrInfo[0] = sever name, usrInfo[1] = user name, usrInfo[2] = DB name
            //usrInfo[3] = port number, usrInfo[4] = password

            StringBuilder cntInfo = new StringBuilder();

            cntInfo.Append("server=");
            cntInfo.Append(usrInfo[0] + ";");
            cntInfo.Append("user=");
            cntInfo.Append(usrInfo[1] + ";");
            cntInfo.Append("database=");
            cntInfo.Append(usrInfo[2] + ";");
            cntInfo.Append("port=");
            cntInfo.Append(usrInfo[3] + ";");
            cntInfo.Append("password=");
            cntInfo.Append(usrInfo[4] + ";");

            conn = new MySqlConnection(cntInfo.ToString());
            try
            {
                Console.WriteLine("Connecting to MySQL...");
                Console.WriteLine(" - Connection Info - ");
                Console.WriteLine(cntInfo.ToString());
                conn.Open();
                // Perform database operations
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }

        public MySqlCommand MySqlCommandSend(string sqlInfo)
        {
            try
            {
                MySqlCommand cmd = new MySqlCommand(sqlInfo, conn);
                cmd.ExecuteNonQuery();
                return cmd;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            MySqlCommand dcmd = null;

            return dcmd;
        }

        public string MySqlRDCommand(string sqlInfo)
        {
            StringBuilder sqlResult = new StringBuilder();

            MySqlDataReader rdr = MySqlCommandSend(sqlInfo).ExecuteReader();
            while (rdr.Read())
            {
                sqlResult.Append(rdr[0] + " -- " + rdr[1]);
                Console.WriteLine(rdr[0] + " -- " + rdr[1]);
            }
            rdr.Close();

            return sqlResult.ToString();
        }

        public void MySqlEXCommand(string sqlInfo)
        {
            MySqlCommandSend(sqlInfo).ExecuteNonQuery();
        }

        public T MySqlSGCommand<T>(string sqlInfo)
        {
            object result = MySqlCommandSend(sqlInfo).ExecuteScalar();
            if (result != null)
            {
                int r = Convert.ToInt32(result);
                Console.WriteLine("Number of countries in the world database is: " + r);
            }

            return (T)Convert.ChangeType(result, typeof(T));
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);

        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {

            }
            conn.Close();
            Console.WriteLine("DB Closed");
        }

        ~MySqlManager()
        {
            Dispose(false);
        }
    }
}

