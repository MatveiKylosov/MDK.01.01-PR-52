using System;

namespace ReportGeneration_Kylosov.Models
{
    public class Work
    {
        public int Id { get; set; }
        public int IdDiscipline { get; set; }
        public int IdType { get; set; }
        public DateTime Date { get; set; }
        public string Name { get; set; }
        public int Semester { get; set; }
        public Work(int Id, int IdDiscipline, int IdType, DateTime Date, string Name, int Semester)
        {
            this.Id = Id;
            this.IdDiscipline = IdDiscipline;
            this.IdType = IdType;
            this.Date = Date;
            this.Name = Name;
            this.Semester = Semester;
        }
    }
}
