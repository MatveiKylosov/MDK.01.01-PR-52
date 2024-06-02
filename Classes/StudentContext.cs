using MySql.Data.MySqlClient;
using ReportGeneration_Kylosov.Classes.Common;
using ReportGeneration_Kylosov.Models;
using System;
using System.Collections.Generic;

namespace ReportGeneration_Kylosov.Classes
{
    public class StudentContext : Student
    {
        public StudentContext(int Id, string Firstname, string Lastname, int IdGroup, bool Expelled, DateTime DateExpelled) :
            base(Id, Firstname, Lastname, IdGroup, Expelled, DateExpelled)
        { }
        public static List<StudentContext> AllStudents()
        {
            List<StudentContext> allStudents = new List<StudentContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader DBStudents = Connection.Query("SELECT * FROM `student` ORDER BY `LastName`", connection);
            while (DBStudents.Read())
            {
                allStudents.Add(new StudentContext(
                    DBStudents.GetInt32(0),
                    DBStudents.GetString(1),
                    DBStudents.GetString(2),
                    DBStudents.GetInt32(3),
                    DBStudents.GetBoolean(4),
                    DBStudents.IsDBNull(5) ? DateTime.Now : DBStudents.GetDateTime(5)));
            }
            Connection.CloseConnection(connection);
            return allStudents;
        }
    }
}
