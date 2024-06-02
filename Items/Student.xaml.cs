using ReportGeneration_Kylosov.Classes;
using System;
using System.Collections.Generic;
using System.Windows.Controls;

namespace ReportGeneration_Kylosov.Items
{
    /// <summary>
    /// Логика взаимодействия для Student.xaml
    /// </summary>
    public partial class Student : UserControl
    {
        public Student(StudentContext student, Pages.Main main)
        {
            InitializeComponent();
            TBFio.Text = $"{student.Lastname} {student.Firstname}";
            CBExpelled.IsChecked = student.Expelled;
            List<DisciplineContext> StudentDisciplines = main.AllDisciplines.FindAll(x => x.IdGroup == student.IdGroup);
            int NecessarilyCount = 0; int WorksCount = 0; int DoneCount = 0; int MissedCount = 0;
            foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
            {
                List<WorkContext> StudentWorks = main.AllWorks.FindAll(x =>
                (x.IdType == 1 || x.IdType == 2 || x.IdType == 3) &&
                x.IdDiscipline == StudentDiscipline.Id);
                NecessarilyCount += StudentWorks.Count;
                foreach (WorkContext StudentWork in StudentWorks)
                {
                    EvaluationContext Evaulation = main.AllEvaluations.Find(x =>
                    x.IdWork == StudentWork.Id &&
                    x.IdStudent == student.Id);
                    if (Evaulation != null && Evaulation.Value.Trim() != "" && Evaulation.Value.Trim() != "2")
                        DoneCount++;
                }
                StudentWorks = main.AllWorks.FindAll(x =>
                x.IdType != 4 && x.IdType != 3);
                WorksCount += StudentWorks.Count;
                foreach (WorkContext StudentWork in StudentWorks)
                {
                    EvaluationContext Evaluation = main.AllEvaluations.Find(x =>
                    x.IdWork == StudentWork.Id &&
                    x.IdStudent == student.Id);
                    if (Evaluation != null && Evaluation.Lateness.Trim() != "")
                        MissedCount += Convert.ToInt32(Evaluation.Lateness);
                }
            }
            doneWorks.Value = (100f / (float)NecessarilyCount) * ((float)DoneCount);
            missedCount.Value = (100f / ((float)WorksCount * 90f)) * ((float)MissedCount);
            TBGroup.Text = main.AllGroups.Find(x => x.Id == student.IdGroup).Name;
        }
    }
}
