using ReportGeneration_Kylosov.Classes;
using ReportGeneration_Kylosov.Items;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Input;

namespace ReportGeneration_Kylosov.Pages
{
    /// <summary>
    /// Логика взаимодействия для Main.xaml
    /// </summary>
    public partial class Main : Page
    {
        public List<GroupContext> AllGroups = GroupContext.AllGroups();
        public List<StudentContext> AllStudents = StudentContext.AllStudents();
        public List<WorkContext> AllWorks = WorkContext.AllWorks();
        public List<EvaluationContext> AllEvaluations = EvaluationContext.AllEvaluations();
        public List<DisciplineContext> AllDisciplines = DisciplineContext.AllDisciplines();

        public Main()
        {
            InitializeComponent();
            CreateGroupUI();
            CreateStudents(AllStudents);
        }

        public void CreateGroupUI()
        {
            foreach (GroupContext group in AllGroups)
                CBGroups.Items.Add(group.Name);

            CBGroups.Items.Add("Выберите");
            CBGroups.SelectedIndex = CBGroups.Items.Count - 1;
        }


        public void CreateStudents(List<StudentContext> AllStudents)
        {
            Parent.Children.Clear();
            foreach (StudentContext student in AllStudents)
                Parent.Children.Add(new Items.Student(student, this));
        }

        private void SelectGroup(object sender, SelectionChangedEventArgs e)
        {
            if (CBGroups.SelectedIndex != CBGroups.Items.Count - 1)
            {
                int IdGroup = AllGroups.Find(x => x.Name == CBGroups.SelectedItem).Id;
                CreateStudents(AllStudents.FindAll(x => x.IdGroup == IdGroup));
            }
        }
        private void SelectStudents(object sender, KeyEventArgs e)
        {
            List<StudentContext> SearchStudent = AllStudents;
            if (CBGroups.SelectedIndex != CBGroups.Items.Count - 1)
            {
                int IdGroup = AllGroups.Find(x => x.Name == CBGroups.SelectedItem).Id;
                SearchStudent = AllStudents.FindAll(x => x.IdGroup == IdGroup);
            }
            CreateStudents(SearchStudent.FindAll(x => $"{x.Lastname} {x.Firstname}".Contains(TBFIO.Text)));
        }
    }
}
