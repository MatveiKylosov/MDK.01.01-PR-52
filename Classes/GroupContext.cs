using MySql.Data.MySqlClient;
using ReportGeneration_Kylosov.Classes.Common;
using ReportGeneration_Kylosov.Models;
using System.Collections.Generic;

namespace ReportGeneration_Kylosov.Classes
{
    public class GroupContext : Group
    {
        public GroupContext(int Id, string Name) : base(Id, Name) { }
        public static List<GroupContext> AllGroups()
        {
            List<GroupContext> allGroups = new List<GroupContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader DBGroups = Connection.Query("SELECT * FROM `group` ORDER BY `Name`", connection);
            while (DBGroups.Read())
            {
                allGroups.Add(new GroupContext(
                    DBGroups.GetInt32(0),
                    DBGroups.GetString(1)));
            }
            Connection.CloseConnection(connection);
            return allGroups;
        }
    }
}
