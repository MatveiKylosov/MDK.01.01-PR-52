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

            // Присваиваем значение в текстовое поле с ФИО
            TBFio.Text = $"{student.Lastname} {student.Firstname}";

            // Галочка "Отчислен" отмечена, если студент отчислен
            CBExpelled.IsChecked = student.Expelled;

            // Получаем дисциплины, в которых участвует студент
            List<DisciplineContext> StudentDisciplines = main.AllDisciplines.FindAll(x => x.IdGroup == student.IdGroup);

            // Создаем переменные отвечающие за расчёты
            int NecessarilyCount = 0; // обязательных работ
            int WorksCount = 0; // всего занятий
            int DoneCount = 0; // выполненных работ
            int MissedCount = 0; // пропущенных минут

            // Перебираем дисциплины
            foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
            {
                // Получаем кол-во работ принадлежащих к группе студента
                // К обязательным работам относятся [теоретические тесты], [экзамены] и [практические работы]
                List<WorkContext> StudentWorks = main.AllWorks.FindAll(x =>
                    (x.IdType == 1 || x.IdType == 2 || x.IdType == 3) &&
                    x.IdDiscipline == StudentDiscipline.Id);

                // Увеличиваем кол-во обязательных работ
                NecessarilyCount += StudentWorks.Count;

                // Перебор обязательных работы
                foreach (WorkContext StudentWork in StudentWorks)
                {
                    // Получаем оценки по работам
                    EvaluationContext Evaulation = main.AllEvaluations.Find(x =>
                        x.IdWork == StudentWork.Id &&
                        x.IdStudent == student.Id);

                    // Проверяем если есть оценка за занятие и она не пустая, и не стоит оценка 2
                    if (Evaulation != null && Evaulation.Value.Trim() != "" && Evaulation.Value.Trim() != "2")
                        // Значит работа сдана
                        DoneCount++;
                }

                // Получаем все занятия, кроме экзамена и оценки за месяц
                StudentWorks = main.AllWorks.FindAll(x =>
                    x.IdType != 4 && x.IdType != 3);

                // Увеличиваем количество занятий
                WorksCount += StudentWorks.Count;

                // Перебираем занятия
                foreach (WorkContext StudentWork in StudentWorks)
                {
                    // Получаем оценки к занятиям с пропусками
                    EvaluationContext Evaluation = main.AllEvaluations.Find(x =>
                        x.IdWork == StudentWork.Id &&
                        x.IdStudent == student.Id);

                    // Если оценка не пустая, и есть прогулы
                    if (Evaluation != null && Evaluation.Lateness.Trim() != "")
                        // Добавляем её в общее кол-во пропущенных минут
                        MissedCount += Convert.ToInt32(Evaluation.Lateness);
                }
            }

            // Выводим в процесс бар по формуле 100/(кол-во занятий)*выполненные
            doneWorks.Value = (100f / (float)NecessarilyCount) * ((float)DoneCount);

            // Выводим в процесс бар по формуле 100/(кол-во занятий * 90 (пара))*пропущенное кол-во минут
            missedCount.Value = (100f / ((float)WorksCount * 90f)) * ((float)MissedCount);

            // Присваиваем значение в текстовое поле с названием группы
            TBGroup.Text = main.AllGroups.Find(x => x.Id == student.IdGroup).Name;
        }
    }
}
