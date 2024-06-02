using MySql.Data.MySqlClient;

namespace ReportGeneration_Kylosov.Classes.Common
{
    public class Connection
    {
        public static string config = "server=127.0.0.1;uid=root;pwd=;database=journal;";
        public static MySqlConnection OpenConnection()
        {
            MySqlConnection conn = new MySqlConnection(config);
            conn.Open();
            return conn;
        }
        public static MySqlDataReader Query(string SQL, MySqlConnection conn)
        {
            return new MySqlCommand(SQL, conn).ExecuteReader();
        }
        public static void CloseConnection(MySqlConnection conn)
        {
            conn.Close();
            MySqlConnection.ClearPool(conn);
        }
    }
}
