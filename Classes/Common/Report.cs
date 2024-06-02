using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using ReportGeneration_Kylosov.Pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportGeneration_Kylosov.Classes.Common
{
    public class Report
    {
        public static void Group(int IdGroup, Main main)
        {
            SaveFileDialog SFD = new SaveFileDialog()
            {
                InitialDirectory = @"C:\",
                Filter = "Excel file (*.xlsx)|*.xlsx",
            };

            SFD.ShowDialog();
            if (SFD.FileName != "")
            {
                GroupContext Group = main.AllGroups.Find(x => x.Id == IdGroup);
                var ExcelApp = new Excel.Application();

                try
                {
                    ExcelApp.Visible = false;
                    Excel.Workbook Workbook = ExcelApp.Workbooks.Add(Type.Missing);
                    Excel.Worksheet Worksheet = Workbook.ActiveSheet as Excel.Worksheet;
                    Worksheet.Name = Group.Name;
                    (Worksheet.Cells[1, 1] as Excel.Range).Value = $"Отчёт о группе {Group.Name}";
                    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 5]].Merge();
                    Styles((Range)Worksheet.Cells[1, 1], 18);
                    (Worksheet.Cells[3, 1] as Excel.Range).Value = $"Список группы:";
                    Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, 5]].Merge();
                    Styles((Range)Worksheet.Cells[3, 1], 12, Excel.XlHAlign.xlHAlignLeft);
                    (Worksheet.Cells[4, 1] as Excel.Range).Value = $"ФИО";
                    Styles((Range)Worksheet.Cells[4, 1], 12, XlHAlign.xlHAlignCenter, true);
                    (Worksheet.Cells[4, 1] as Excel.Range).ColumnWidth = 35.0f;
                    (Worksheet.Cells[4, 2] as Excel.Range).Value = $"Кол-во не сданных практических";
                    Styles((Range)Worksheet.Cells[4, 2], 12, XlHAlign.xlHAlignCenter, true);
                    (Worksheet.Cells[4, 3] as Excel.Range).Value = $"Кол-во не сданных теоретических";
                    Styles((Range)Worksheet.Cells[4, 3], 12, XlHAlign.xlHAlignCenter, true);
                    (Worksheet.Cells[4, 4] as Excel.Range).Value = $"Отсутствовал на паре";
                    Styles((Range)Worksheet.Cells[4, 4], 12, XlHAlign.xlHAlignCenter, true);
                    (Worksheet.Cells[4, 5] as Excel.Range).Value = $"Опоздал";
                    Styles((Range)Worksheet.Cells[4, 5], 12, XlHAlign.xlHAlignCenter, true);
                    int Height = 5;
                    List<StudentContext> Students = main.AllStudents.FindAll(x => x.IdGroup == IdGroup);

                    var bestStudent = Students
                    .Select(student =>
                    {
                        List<DisciplineContext> StudentDisciplines = main.AllDisciplines.FindAll(x => x.IdGroup == student.IdGroup);
                        int PracticeCount = 0; int TheoryCount = 0; int AbsenteeismCount = 0; int LateCount = 0;
                        foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
                        {
                            List<WorkContext> StudentWorks = main.AllWorks.FindAll(x => x.IdDiscipline == StudentDiscipline.Id);
                            foreach (WorkContext StudentWork in StudentWorks)
                            {
                                EvaluationContext Evaluation = main.AllEvaluations.Find(x => x.IdWork == StudentWork.Id && x.IdStudent == student.Id);
                                if ((Evaluation != null && (Evaluation.Value.Trim() == "" || Evaluation.Value.Trim() == "2")) || Evaluation == null)
                                {
                                    if (StudentWork.IdType == 1)
                                        PracticeCount++;
                                    else if (StudentWork.IdType == 2)
                                        TheoryCount++;
                                }
                                if (Evaluation != null && Evaluation.Lateness.Trim() != "")
                                {
                                    if (Convert.ToInt32(Evaluation.Lateness) == 90)
                                        AbsenteeismCount++;
                                    else LateCount++;
                                }
                            }
                        }
                        return new
                        {
                            Student = student,
                            PracticeCount,
                            TheoryCount,
                            AbsenteeismCount,
                            LateCount,
                            Total = PracticeCount + TheoryCount,
                            Attendance = 100 - (AbsenteeismCount * 100.0 / StudentDisciplines.Count)
                        };
                    })
                    .OrderBy(x => x.Total)
                    .ThenByDescending(x => x.Attendance)
                    .FirstOrDefault();

                    foreach (StudentContext student in Students)
                    {
                        List<DisciplineContext> StudentDisciplines = main.AllDisciplines.FindAll(x => x.IdGroup == student.IdGroup);
                        int PracticeCount = 0; int TheoryCount = 0; int AbsenteeismCount = 0; int LateCount = 0;
                        foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
                        {
                            List<WorkContext> StudentWorks = main.AllWorks.FindAll(x => x.IdDiscipline == StudentDiscipline.Id);
                            foreach (WorkContext StudentWork in StudentWorks)
                            {
                                EvaluationContext Evaluation = main.AllEvaluations.Find(x => x.IdWork == StudentWork.Id && x.IdStudent == student.Id);
                                if ((Evaluation != null && (Evaluation.Value.Trim() == "" || Evaluation.Value.Trim() == "2"))
                                    || Evaluation == null)
                                {
                                    if (StudentWork.IdType == 1)
                                        PracticeCount++;
                                    else if (StudentWork.IdType == 2)
                                        TheoryCount++;
                                }
                                if (Evaluation != null && Evaluation.Lateness.Trim() != "")
                                {
                                    if (Convert.ToInt32(Evaluation.Lateness) == 90)
                                        AbsenteeismCount++;
                                    else LateCount++;
                                }
                            }
                        }
                        (Worksheet.Cells[Height, 1] as Excel.Range).Value = $"{student.Lastname} {student.Firstname}";
                        Styles(Worksheet.Cells[Height, 1] as Excel.Range, 12, XlHAlign.xlHAlignLeft, true);
                        (Worksheet.Cells[Height, 2] as Excel.Range).Value = PracticeCount.ToString();
                        Styles(Worksheet.Cells[Height, 2] as Excel.Range, 12, XlHAlign.xlHAlignCenter, true);
                        (Worksheet.Cells[Height, 3] as Excel.Range).Value = TheoryCount.ToString();
                        Styles(Worksheet.Cells[Height, 3] as Excel.Range, 12, XlHAlign.xlHAlignCenter, true);
                        (Worksheet.Cells[Height, 4] as Excel.Range).Value = AbsenteeismCount.ToString();
                        Styles(Worksheet.Cells[Height, 4] as Excel.Range, 12, XlHAlign.xlHAlignCenter, true);
                        (Worksheet.Cells[Height, 5] as Excel.Range).Value = LateCount.ToString();
                        Styles(Worksheet.Cells[Height, 5] as Excel.Range, 12, XlHAlign.xlHAlignCenter, true);

                        if (bestStudent != null && student.Id == bestStudent.Student.Id)
                        {
                            Excel.Range row = Worksheet.Range[Worksheet.Cells[Height, 1], Worksheet.Cells[Height, 5]];
                            row.Interior.Color = Excel.XlRgbColor.rgbLightGreen;
                        }

                        Height++;
                    }
                    Workbook.SaveAs(SFD.FileName);
                    Workbook.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                ExcelApp.Quit();
            }
        }
        public static void Styles(Excel.Range Cell,
            int FontSize, Excel.XlHAlign Position = Excel.XlHAlign.xlHAlignCenter,
            bool Border = false)
        {
            Cell.Font.Name = "Bahnschhrift Light Condensed";
            Cell.Font.Size = FontSize;
            Cell.HorizontalAlignment = Position;
            Cell.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
            if (Border)
            {
                Excel.Borders border = Cell.Borders;
                border.LineStyle = Excel.XlLineStyle.xlDouble;
                border.Weight = XlBorderWeight.xlThin;
                Cell.WrapText = true;
            }
        }
    }
}
