namespace ReportGeneration_Kylosov.Models
{
    public class Group
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public Group(int Id, string Name)
        {
            this.Id = Id;
            this.Name = Name;
        }
    }
}
