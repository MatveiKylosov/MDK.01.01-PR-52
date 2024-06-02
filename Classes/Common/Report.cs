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
        // Метод для генерации отчета по группе
        public static void Group(int IdGroup, Main main)
        {
            // Открываем диалоговое окно для сохранения файла
            SaveFileDialog SFD = new SaveFileDialog()
            {
                InitialDirectory = @"C:\",
                Filter = "Excel file (*.xlsx)|*.xlsx",
            };
            SFD.ShowDialog();
            if (SFD.FileName != "")
            {
                // Находим группу по Id
                GroupContext Group = main.AllGroups.Find(x => x.Id == IdGroup);
                // Создаем экземпляр Excel приложения
                var ExcelApp = new Excel.Application();

                try
                {
                    // Делаем Excel невидимым
                    ExcelApp.Visible = false;
                    // Создаем новую рабочую книгу
                    Excel.Workbook Workbook = ExcelApp.Workbooks.Add(Type.Missing);
                    // Получаем активный лист и называем его именем группы
                    Excel.Worksheet Worksheet = Workbook.ActiveSheet as Excel.Worksheet;
                    Worksheet.Name = Group.Name;

                    // Устанавливаем заголовок отчета
                    SetCellValueAndStyle(Worksheet, 1, 1, $"Отчёт о группе {Group.Name}", 18, Excel.XlHAlign.xlHAlignCenter, true);
                    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 5]].Merge();
                    // Устанавливаем подзаголовок "Список группы"
                    SetCellValueAndStyle(Worksheet, 3, 1, "Список группы:", 12, Excel.XlHAlign.xlHAlignLeft, true);
                    Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, 5]].Merge();

                    // Заголовки столбцов
                    string[] headers = { "ФИО", "Кол-во не сданных практических", "Кол-во не сданных теоретических", "Отсутствовал на паре", "Опоздал" };
                    double[] columnWidths = { 35.0f, 20.0f, 20.0f, 20.0f, 20.0f };

                    // Устанавливаем заголовки и ширину столбцов
                    for (int i = 0; i < headers.Length; i++)
                    {
                        SetCellValueAndStyle(Worksheet, 4, i + 1, headers[i], 12, Excel.XlHAlign.xlHAlignCenter, true, columnWidths[i]);
                    }

                    int row = 5;  // Начальная строка для данных студентов
                    StudentContext bestStudent = null;  // Лучший студент
                    int bestStudentRow = -1;  // Строка лучшего студента
                    int bestScore = int.MaxValue;  // Наименьшее количество не сданных работ
                    int bestAttendance = int.MinValue;  // Наименьшее количество пропусков
                    int sheetNumber = 1;  // Номер листа

                    // Перебираем всех студентов группы
                    foreach (var student in main.AllStudents.FindAll(x => x.IdGroup == IdGroup))
                    {
                        // Подсчитываем метрики студента
                        int[] counts = CountStudentMetrics(main, student);
                        int totalMissed = counts[0] + counts[1];  // Сумма не сданных работ
                        int attendance = counts[2];  // Количество пропусков

                        // Определяем самого успешного студента
                        if (totalMissed < bestScore || (totalMissed == bestScore && attendance < bestAttendance))
                        {
                            bestStudent = student;
                            bestScore = totalMissed;
                            bestAttendance = attendance;
                            bestStudentRow = row;
                        }

                        // Заполняем данные по студенту
                        SetCellValueAndStyle(Worksheet, row, 1, $"{student.Lastname} {student.Firstname}", 12, Excel.XlHAlign.xlHAlignLeft, true);

                        for (int col = 0; col < counts.Length; col++)
                        {
                            SetCellValueAndStyle(Worksheet, row, col + 2, counts[col].ToString(), 12, Excel.XlHAlign.xlHAlignCenter, true);
                        }
                        row++;

                        // Создаем отдельный лист для каждого студента
                        Excel.Worksheet studentWorksheet = Workbook.Worksheets.Add(Type.Missing, Workbook.Worksheets[Workbook.Worksheets.Count]);
                        studentWorksheet.Name = $"{sheetNumber}. {student.Lastname} {student.Firstname}";

                        // Заголовок листа студента
                        SetCellValueAndStyle(studentWorksheet, 1, 1, $"Отчёт о студенте {student.Lastname} {student.Firstname}", 18, Excel.XlHAlign.xlHAlignCenter, true);
                        studentWorksheet.Range[studentWorksheet.Cells[1, 1], studentWorksheet.Cells[1, 5]].Merge();

                        // Заголовки столбцов для листа студента
                        string[] studentHeaders = { "Тип работы", "Название работы", "Статус", "Оценка", "Дата" };
                        double[] studentColumnWidths = { 20.0f, 35.0f, 15.0f, 10.0f, 15.0f };

                        for (int j = 0; j < studentHeaders.Length; j++)
                        {
                            SetCellValueAndStyle(studentWorksheet, 3, j + 1, studentHeaders[j], 12, Excel.XlHAlign.xlHAlignCenter, true, studentColumnWidths[j]);
                        }

                        int studentRow = 4;  // Начальная строка для данных студента
                        // Перебираем все дисциплины студента
                        foreach (var discipline in main.AllDisciplines.FindAll(x => x.IdGroup == student.IdGroup))
                        {
                            // Перебираем все работы по дисциплине
                            foreach (var work in main.AllWorks.FindAll(x => x.IdDiscipline == discipline.Id))
                            {
                                var evaluation = main.AllEvaluations.Find(x => x.IdWork == work.Id && x.IdStudent == student.Id);

                                // Определяем статус работы и оценку
                                string status = (evaluation != null && !string.IsNullOrWhiteSpace(evaluation.Value) && evaluation.Value.Trim() != "2") ? "Сдано" : "Не сдано";
                                string grade = evaluation != null ? evaluation.Value : "Нет оценки";

                                // Заполняем данные по работе
                                SetCellValueAndStyle(studentWorksheet, studentRow, 1, work.IdType == 1 ? "Практическая" : "Теоретическая", 12, Excel.XlHAlign.xlHAlignLeft, true);
                                SetCellValueAndStyle(studentWorksheet, studentRow, 2, work.Name, 12, Excel.XlHAlign.xlHAlignLeft, true);
                                SetCellValueAndStyle(studentWorksheet, studentRow, 3, status, 12, Excel.XlHAlign.xlHAlignCenter, true);
                                SetCellValueAndStyle(studentWorksheet, studentRow, 4, grade, 12, Excel.XlHAlign.xlHAlignCenter, true);

                                studentRow++;
                            }
                        }
                        sheetNumber++;  // Увеличиваем номер листа
                    }

                    // Выделяем строку самого успешного студента
                    if (bestStudentRow != -1)
                    {
                        Worksheet.Range[Worksheet.Cells[bestStudentRow, 1], Worksheet.Cells[bestStudentRow, 5]].Interior.Color = Excel.XlRgbColor.rgbLightGreen;
                    }

                    // Сохраняем рабочую книгу
                    Workbook.SaveAs(SFD.FileName);
                    Workbook.Close();
                }
                catch (Exception ex)
                {
                    // Отображаем сообщение об ошибке
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    // Закрываем Excel приложение
                    ExcelApp.Quit();
                }
                MessageBox.Show("Конец");
            }
        }

        // Метод для установки значения и стиля ячейки
        private static void SetCellValueAndStyle(Excel.Worksheet worksheet, int row, int column, string value, int fontSize, Excel.XlHAlign alignment, bool border = false, double columnWidth = 0)
        {
            var cell = worksheet.Cells[row, column] as Excel.Range;
            cell.Value = value;
            cell.Font.Name = "Bahnschrift Light Condensed";
            cell.Font.Size = fontSize;
            cell.HorizontalAlignment = alignment;
            cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cell.WrapText = true;

            if (border)
            {
                var borders = cell.Borders;
                borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                borders.Weight = Excel.XlBorderWeight.xlThin;
            }

            if (columnWidth > 0)
            {
                cell.ColumnWidth = columnWidth;
            }
        }

        // Метод для подсчета метрик студента
        private static int[] CountStudentMetrics(Main main, StudentContext student)
        {
            int practiceCount = 0, theoryCount = 0, absenteeismCount = 0, lateCount = 0;

            // Перебираем все дисциплины студента
            foreach (var discipline in main.AllDisciplines.FindAll(x => x.IdGroup == student.IdGroup))
            {
                // Перебираем все работы по дисциплине
                foreach (var work in main.AllWorks.FindAll(x => x.IdDiscipline == discipline.Id))
                {
                    var evaluation = main.AllEvaluations.Find(x => x.IdWork == work.Id && x.IdStudent == student.Id);

                    // Увеличиваем счетчики в зависимости от типа работы и статуса оценки
                    if ((evaluation != null && (string.IsNullOrWhiteSpace(evaluation.Value) || evaluation.Value.Trim() == "2")) || evaluation == null)
                    {
                        if (work.IdType == 1) practiceCount++;
                        else if (work.IdType == 2) theoryCount++;
                    }

                    // Увеличиваем счетчики пропусков и опозданий
                    if (evaluation != null && !string.IsNullOrWhiteSpace(evaluation.Lateness))
                    {
                        if (Convert.ToInt32(evaluation.Lateness) == 90) absenteeismCount++;
                        else lateCount++;
                    }
                }
            }

            return new[] { practiceCount, theoryCount, absenteeismCount, lateCount };
        }
    }
}
