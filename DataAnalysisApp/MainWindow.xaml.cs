﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using RDotNet;
using Xceed.Wpf.Toolkit;

using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;


namespace DataAnalysisApp
{
    public partial class MainWindow : Window
    {
        private Word.Application WordApp;
        private Word.Document WordDoc;
        private Excel.Application ExcelApp;
        private Excel.Workbook ExcelBook;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            REngine engine = REngine.GetInstance();
            engine.Initialize();
            engine.Evaluate("library(Hmisc)");
            engine.Evaluate("require(qgraph)");
            engine.Evaluate("data <- read.table('student-mat.csv',sep=',',header=TRUE)");

            object begin = 0;
            object end = 0;
            WordApp = new Word.Application();
            WordApp.Visible = true;
            string ReportPath = @"C:\DataAnalysisApp\report.docx";
            WordDoc = WordApp.Documents.Add(ReportPath);
            Word.Range wordrange = WordDoc.Range(ref begin, ref end);
            wordrange = WordDoc.Bookmarks["Название"].Range;
            wordrange.Text = watemark1.Text;
            wordrange = WordDoc.Bookmarks["Автор"].Range;
            wordrange.Text = watemark1_Copy.Text;
            wordrange = WordDoc.Bookmarks["Дата"].Range;
            wordrange.Text = dateTimePicker1.SelectedDate.Value.Date.ToShortDateString();
            CheckBox[] CheckboxArr = new CheckBox[7] { checkBox, checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox7 };
            Word.Range wordcellrange;
            CharacterVector CountOfRows = engine.Evaluate("nrow(data)").AsCharacter();

            if (checkBox6.IsChecked == true)
            {
                string ExcelReportPath = @"C:\DataAnalysisApp\reportExcel.xlsx";
                ExcelApp = new Excel.Application();
                ExcelApp.Visible = true;
                ExcelBook = ExcelApp.Workbooks.Add(ExcelReportPath);
                ExcelBook.Worksheets.Add();
                ExcelBook.Worksheets.Add();
                ExcelBook.Worksheets.Add();
            }
            if (checkBox8.IsChecked == false)
            {

                wordrange = WordDoc.Bookmarks["Опис_Данные"].Range;
                wordrange.InsertAfter("Анализировались данные о студентах, обучающихся в старших классах двух школ на математическом курсе. \r");
                wordrange.InsertAfter("Данные содержат 395 записей и 33 переменных.");

                foreach (CheckBox CurrentCheckbox in CheckboxArr)
                {
                    if (CurrentCheckbox.IsChecked == true)
                    {
                        switch (CurrentCheckbox.Name)
                        {
                            case "checkBox":

                                wordrange = WordDoc.Bookmarks["Опис"].Range;
                                wordrange.InsertAfter("Характеристики метрических переменных исследуемого набора данных представлены в таблице:");

                                wordrange = WordDoc.Bookmarks["Опис_Таблица"].Range;
                                Word.Table wordTable = WordDoc.Tables.Add(wordrange, 4, 5);

                                WordDoc.Tables[1].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                                WordDoc.Tables[1].Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                                string[] columnNames = new string[5] { "Переменная", "Минимум", "Среднее", "Медиана", "Максимум" };


                                for (int i = 1; i < 6; i++)
                                {
                                    wordcellrange = WordDoc.Tables[1].Cell(1, i).Range;
                                    wordcellrange.Text = columnNames[i - 1];
                                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                }

                                string[] rowNames = new string[3] { "Возраст", "Учебное время", "Прогулы" };

                                for (int i = 2; i < 5; i++)
                                {
                                    wordcellrange = WordDoc.Tables[1].Cell(i, 1).Range;
                                    wordcellrange.Text = rowNames[i - 2];
                                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                }

                                CharacterVector Age = engine.Evaluate("summary(data$age)").AsCharacter();
                                CharacterVector StudyTime = engine.Evaluate("summary(data$studytime)").AsCharacter();
                                CharacterVector Absences = engine.Evaluate("summary(data$absences)").AsCharacter();
                                CharacterVector[] VectorArr = new CharacterVector[3] { Age, StudyTime, Absences };
                                string[] AgeRow = new string[4];
                                string[] StudyTimeRow = new string[4];
                                string[] AbsencesRow = new string[4];
                                string[][] RowsArr = new string[3][] { AgeRow, StudyTimeRow, AbsencesRow };
                                AgeRow[0] = Age[0];

                                for (int i = 0; i < 4; i++)
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            for (int j = 0; j < 3; j++)
                                                RowsArr[j][0] = VectorArr[j][0];
                                            break;
                                        case 1:
                                            for (int j = 0; j < 3; j++)
                                                RowsArr[j][1] = Math.Round(Convert.ToDouble(VectorArr[j][3])).ToString();
                                            break;
                                        case 2:
                                            for (int j = 0; j < 3; j++)
                                                RowsArr[j][2] = Math.Round(Convert.ToDouble(VectorArr[j][2])).ToString();
                                            break;
                                        case 3:
                                            for (int j = 0; j < 3; j++)
                                                RowsArr[j][3] = VectorArr[j][5];
                                            break;
                                    }
                                }
                                for (int i = 2; i < 5; i++)
                                    for (int j = 2; j < 6; j++)
                                    {
                                        wordcellrange = WordDoc.Tables[1].Cell(i, j).Range;
                                        wordcellrange.Text = RowsArr[i - 2][j - 2];
                                        wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                    }

                                wordrange = WordDoc.Bookmarks["Опис_Текст"].Range;

                                double InterquartileRange = double.Parse(Absences[4]) - double.Parse(Absences[1]);
                                double[] InternalBoundary = new double[2];
                                InternalBoundary[0] = InterquartileRange * 1.5 - double.Parse(Absences[0]);
                                InternalBoundary[1] = InterquartileRange * 1.5 + double.Parse(Absences[4]);

                                double[] OuterBoundary = new double[2];
                                OuterBoundary[0] = InterquartileRange * 3.0 - double.Parse(Absences[0]);
                                OuterBoundary[1] = InterquartileRange * 3.0 + double.Parse(Absences[4]);

                                engine.Evaluate("expValues <- data.frame(data$absences)");
                                CharacterVector ExperienceValues = engine.Evaluate("data$absences").AsCharacter();

                                bool Flag = false;
                                for (int i = 0; i < ExperienceValues.Length - 1; i++)
                                {
                                    if (double.Parse(ExperienceValues[i]) < InternalBoundary[0] ||
                                        double.Parse(ExperienceValues[i]) > InternalBoundary[1] ||
                                        double.Parse(ExperienceValues[i]) < OuterBoundary[0] ||
                                        double.Parse(ExperienceValues[i]) > OuterBoundary[1])
                                        Flag = true;
                                }
                                string FlagStr = "";
                                if (Flag == true) FlagStr = "наличии выбросов";
                                else FlagStr = "несмещенности данных";

                                wordrange.InsertAfter("Так, переменная 'Прогулы' изменяется от " + Absences[0] + " до " + Absences[5] + " со средним значением равным " + Math.Round(Convert.ToDouble(Absences[3])).ToString() + " и медианой равной " + Math.Round(Convert.ToDouble(Absences[2])).ToString() + ", что говорит о " + FlagStr + ".");

                                wordrange.InsertParagraphAfter();
                                engine.Evaluate("internet <- summary(data$internet)");
                                CharacterVector Internet = engine.Evaluate("internet").AsCharacter();
                                double ProcentOfInternetAcceced = Math.Round(double.Parse(Internet[1]) / double.Parse(CountOfRows[0]) * 100);
                                wordrange.InsertAfter("Количество студентов с доступом к интернету дома составляет " + Convert.ToInt32(ProcentOfInternetAcceced).ToString() + "% выборки.");

                                wordrange.InsertParagraphAfter();
                                engine.Evaluate("romantic <- summary(data$romantic)");
                                CharacterVector Romantic = engine.Evaluate("romantic").AsCharacter();
                                double ProcentOfRomantic = Math.Round(double.Parse(Romantic[1]) / double.Parse(CountOfRows[0]) * 100);
                                wordrange.InsertAfter("Количество студентов, состоящих в романтических отношениях составляет " + Convert.ToInt32(ProcentOfRomantic).ToString() + "% выборки.");

                                wordrange.InsertParagraphAfter();
                                engine.Evaluate("higher <- summary(data$higher)");
                                CharacterVector Edu = engine.Evaluate("higher").AsCharacter();
                                double ProcentOfEdu = Math.Round(double.Parse(Edu[1]) / double.Parse(CountOfRows[0]) * 100);
                                wordrange.InsertAfter("Количество студентов, которые хотят получить высшее образование составляет " + Convert.ToInt32(ProcentOfEdu).ToString() + "% выборки.");

                                wordrange.InsertParagraphAfter();
                                engine.Evaluate("Mjob <- summary(data$Mjob)");
                                CharacterVector MotherJob = engine.Evaluate("Mjob").AsCharacter();
                                double[] MotherJobProcent = new double[5];
                                for (int i = 0; i < 5; i++)
                                {
                                    MotherJobProcent[i] = Math.Round(double.Parse(MotherJob[i]) / double.Parse(CountOfRows[0]) * 100);
                                }

                                wordrange.InsertAfter("По показателю 'Место работы матери' выборка распределена следующим образом: 'Учитель' -- " + MotherJobProcent[4].ToString() + "%; 'Врач' -- " + MotherJobProcent[1].ToString() + "%; 'Государственная служба' -- " + MotherJobProcent[3].ToString() + "%; 'Домохозяйка' -- " + MotherJobProcent[0].ToString() + "%; 'Другое' --  " + MotherJobProcent[2].ToString() + "%.");

                                wordrange.InsertParagraphAfter();
                                engine.Evaluate("reason <- summary(data$reason)");
                                CharacterVector Reason = engine.Evaluate("reason").AsCharacter();
                                double[] ReasonProcent = new double[4];
                                for (int i = 0; i < 4; i++)
                                {
                                    ReasonProcent[i] = Math.Round(double.Parse(Reason[i]) / double.Parse(CountOfRows[0]) * 100);
                                }

                                wordrange.InsertAfter("По показателю 'Причина выбора данной школы' выборка распределена следующим образом: 'Репутация' -- " + ReasonProcent[3].ToString() + "%; 'Интерес к курсу' -- " + ReasonProcent[0].ToString() + "%; 'Близость к дому' -- " + ReasonProcent[1].ToString() + "%; 'Другое' -- " + ReasonProcent[2].ToString() + "%.");

                                wordrange = WordDoc.Bookmarks["Разрыв0"].Range;
                                wordrange.InsertBreak(Word.WdBreakType.wdPageBreak);

                                if (checkBox6.IsChecked == true)
                                {

                                    ExcelBook.Worksheets[1].Name = "Descriptive statistics";
                                    for (int i = 1; i < 5; i++)
                                        for (int j = 1; j < 6; j++)
                                        {
                                            (ExcelBook.Worksheets[1].Cells(i, j) as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                            ExcelBook.Worksheets[1].Cells[i, j].Value = wordTable.Cell(i, j).Range.Text.Substring(0, wordTable.Cell(i, j).Range.Text.Length - 1);
                                            ExcelBook.Worksheets[1].Cells[i, j].EntireColumn.ColumnWidth = 20;
                                        }
                                }

                                break;

                            case "checkBox1":

                                engine.Evaluate("data <- as.data.frame(data)");
                                wordrange = WordDoc.Bookmarks["Т_Тест"].Range;
                                wordrange.InsertAfter("Рассмотрим зависимость числа прогулов от того, состоит ли студент в романтических отношениях, учился в дестком саду, помогают ли ему в семье с учебой и где студент живет. \r");

                                wordrange.InsertParagraphAfter();
                                double TTestPValue = engine.Evaluate("t.test(data[,30] ~ data$romantic)$p.value").AsNumeric()[0];
                                double TTestTValue = engine.Evaluate("t.test(data[,30] ~ data$romantic)$statistic").AsNumeric()[0];
                                if (TTestPValue <= 0.05) wordrange.InsertAfter("Согласно критерию Стьюлента выявлены статистически значимые различия между 'Прогулы' и 'Отношения' (p.value =  " + Math.Round(TTestPValue, 5) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");
                                else wordrange.InsertAfter("Согласно критерию Стьюлента не выявлены статистически значимые различия между 'Прогулы' и 'Отношения' (p.value =  " + Math.Round(TTestPValue, 5) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");

                                wordrange.InsertParagraphAfter();
                                TTestPValue = engine.Evaluate("t.test(data[,30] ~ data$nursery)$p.value").AsNumeric()[0];
                                TTestTValue = engine.Evaluate("t.test(data[,30] ~ data$nursery)$statistic").AsNumeric()[0];
                                if (TTestPValue <= 0.05) wordrange.InsertAfter("Согласно критерию Стьюлента выявлены статистически значимые различия между 'Прогулы' и 'Детский сад' (p.value =  " + Math.Round(TTestPValue, 5) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");
                                else wordrange.InsertAfter("Согласно критерию Стьюлента не выявлены статистически значимые различия между 'Прогулы' и 'Детский сад' (p.value =  " + Math.Round(TTestPValue, 5) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");

                                wordrange.InsertParagraphAfter();
                                TTestPValue = engine.Evaluate("t.test(data[,30] ~ data$famsup)$p.value").AsNumeric()[0];
                                TTestTValue = engine.Evaluate("t.test(data[,30] ~ data$famsup)$statistic").AsNumeric()[0];
                                if (TTestPValue <= 0.05) wordrange.InsertAfter("Согласно критерию Стьюлента выявлены статистически значимые различия между 'Прогулы' и 'Помощь с учебой в семье' (p.value =  " + Math.Round(TTestPValue, 5) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");
                                else wordrange.InsertAfter("Согласно критерию Стьюлента не выявлены статистически значимые различия между 'Прогулы' и 'Помощь с учебой в семье' (p.value =  " + Math.Round(TTestPValue, 5) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");

                                wordrange.InsertParagraphAfter();
                                TTestPValue = engine.Evaluate("t.test(data[,30] ~ data$address)$p.value").AsNumeric()[0];
                                TTestTValue = engine.Evaluate("t.test(data[,30] ~ data$address)$statistic").AsNumeric()[0];
                                if (TTestPValue <= 0.05) wordrange.InsertAfter("Согласно критерию Стьюлента выявлены статистически значимые различия между 'Прогулы' и 'Место проживания' (p.value =  " + Math.Round(TTestPValue, 5) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");
                                else wordrange.InsertAfter("Согласно критерию Стьюлента не выявлены статистически значимые различия между 'Прогулы' и 'Место проживания' (p.value =  " + Math.Round(TTestPValue, 5) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");

                                wordrange = WordDoc.Bookmarks["Разрыв"].Range;
                                wordrange.InsertBreak(Word.WdBreakType.wdPageBreak);

                                break;

                            case "checkBox2":

                                wordrange = WordDoc.Bookmarks["Анова"].Range;
                                wordrange.InsertParagraphAfter();
                                wordrange.InsertAfter("С помощью теста ANOVA были изучены различия по количеству прогулов для переменных: 'Работа матери', 'Опекун студента', 'Количество употребляемого алкоголя в будние дни'.\r");

                                wordrange.InsertAfter(" Работа матери (5 групп) : p value = " + Math.Round(engine.Evaluate("(summary(aov(absences ~ Mjob, data = data))[[1]][[5]])[1]").AsNumeric()[0], 5) + "; F value = " +
                                    Math.Round(engine.Evaluate("(summary(aov(absences ~ Mjob, data = data))[[1]][[4]])[1]").AsNumeric()[0], 5) + ". ");
                                if (Math.Round(engine.Evaluate("(summary(aov(absences ~ Mjob, data = data))[[1]][[5]])[1]").AsNumeric()[0], 5) < 0.05)
                                    wordrange.InsertAfter("Выявлены зависимости в группах по данному параметру.");
                                else wordrange.InsertAfter("Не выявлены различия в группах по данному параметру.");

                                wordrange.InsertParagraphAfter();
                                wordrange.InsertAfter("'Опекун студента (3 группы)': p value = " + Math.Round(engine.Evaluate("(summary(aov(absences ~ guardian, data = data))[[1]][[5]])[1]").AsNumeric()[0], 5) + "; F value = " +
                                    Math.Round(engine.Evaluate("(summary(aov(absences ~ guardian, data = data))[[1]][[4]])[1]").AsNumeric()[0], 5) + ". ");
                                if (Math.Round(engine.Evaluate("(summary(aov(absences ~ guardian, data = data))[[1]][[5]])[1]").AsNumeric()[0], 5) < 0.05)
                                    wordrange.InsertAfter("Выявлены зависимости в группах по данному параметру.");
                                else wordrange.InsertAfter("Не выявлены различия в группах по данному параметру.");

                                wordrange.InsertParagraphAfter();
                                wordrange.InsertAfter("'Количество употребляемого алкгоголя в будние дни (5 групп)': p value = " + Math.Round(engine.Evaluate("(summary(aov(absences ~ Dalc, data = data))[[1]][[5]])[1]").AsNumeric()[0], 5) + "; F value = " +
                                     Math.Round(engine.Evaluate("(summary(aov(absences ~ Dalc, data = data))[[1]][[4]])[1]").AsNumeric()[0], 5) + ". ");
                                if (Math.Round(engine.Evaluate("(summary(aov(absences ~ Dalc, data = data))[[1]][[5]])[1]").AsNumeric()[0], 5) < 0.05)
                                    wordrange.InsertAfter("Выявлены зависимости в группах по данному параметру.");
                                else wordrange.InsertAfter("Не выявлены различия в группах по данному параметру.");
                                break;

                            case "checkBox3":

                                wordrange = WordDoc.Bookmarks["Манна_Уитни"].Range;
                                wordrange.InsertParagraphAfter();

                                wordrange.InsertAfter("Переменная " + engine.Evaluate("colnames(data)[27]").AsCharacter()[0] + " принимает всего " + engine.Evaluate("length(levels(factor(data$Dalc)))").AsCharacter()[0] + " значений, поэтому для выявления различий по " + engine.Evaluate("colnames(data)[27]").AsCharacter()[0] + " по показателю " + engine.Evaluate("colnames(data)[2]").AsCharacter()[0] + " используем критерий Манна-Уитни.");
                                string ResFlag = "";
                                if (engine.Evaluate("wilcox.test(Dalc~sex, data = data)[3]").AsNumeric()[0] > 0.05) ResFlag = "не ";
                                wordrange.InsertAfter(" Согласно этому критерию " + ResFlag + "выявлены статистически значимые различия между " + engine.Evaluate("levels(data$sex)").AsCharacter()[0] + " и " + engine.Evaluate("levels(data$sex)").AsCharacter()[1] + "");
                                wordrange.InsertAfter(" (p = " + engine.Evaluate("wilcox.test(Dalc~sex, data = data)").AsCharacter()[2] + ", W = " + engine.Evaluate("wilcox.test(Dalc~sex, data = data)").AsCharacter()[0] + ").");

                                wordrange.InsertParagraphAfter();

                                wordrange.InsertAfter("Переменная " + engine.Evaluate("colnames(data)[25]").AsCharacter()[0] + " принимает всего " + engine.Evaluate("length(levels(factor(data$freetime)))").AsCharacter()[0] + " значений, поэтому для выявления различий по " + engine.Evaluate("colnames(data)[25]").AsCharacter()[0] + " по показателю " + engine.Evaluate("colnames(data)[22]").AsCharacter()[0] + " используем критерий Манна-Уитни.");
                                ResFlag = "";
                                if (engine.Evaluate("wilcox.test(freetime~internet, data = data)[3]").AsNumeric()[0] > 0.05) ResFlag = "не ";
                                wordrange.InsertAfter(" Согласно этому критерию " + ResFlag + "выявлены статистически значимые различия между " + engine.Evaluate("levels(data$internet)").AsCharacter()[0] + " и " + engine.Evaluate("levels(data$internet)").AsCharacter()[1] + "");
                                wordrange.InsertAfter(" (p = " + engine.Evaluate("wilcox.test(freetime~internet, data = data)").AsCharacter()[2] + ", W = " + engine.Evaluate("wilcox.test(freetime~internet, data = data)").AsCharacter()[0] + ").");

                                wordrange.InsertParagraphAfter();
                                wordrange.InsertAfter("Переменная " + engine.Evaluate("colnames(data)[27]").AsCharacter()[0] + " принимает всего " + engine.Evaluate("length(levels(factor(data$Dalc)))").AsCharacter()[0] + " значений, поэтому для выявления различий по " + engine.Evaluate("colnames(data)[27]").AsCharacter()[0] + " по показателю " + engine.Evaluate("colnames(data)[23]").AsCharacter()[0] + " используем критерий Манна-Уитни.");
                                ResFlag = "";
                                if (engine.Evaluate("wilcox.test(Dalc~romantic, data = data)[3]").AsNumeric()[0] > 0.05) ResFlag = "не ";
                                wordrange.InsertAfter(" Согласно этому критерию " + ResFlag + "выявлены статистически значимые различия между " + engine.Evaluate("levels(data$romantic)").AsCharacter()[0] + " и " + engine.Evaluate("levels(data$romantic)").AsCharacter()[1] + "");
                                wordrange.InsertAfter(" (p = " + engine.Evaluate("wilcox.test(Dalc~romantic, data = data)").AsCharacter()[2] + ", W = " + engine.Evaluate("wilcox.test(Dalc~romantic, data = data)").AsCharacter()[0] + ").");

                                wordrange.InsertParagraphAfter();

                                wordrange.InsertAfter("Переменная " + engine.Evaluate("colnames(data)[26]").AsCharacter()[0] + " принимает всего " + engine.Evaluate("length(levels(factor(data$goout)))").AsCharacter()[0] + " значений, поэтому для выявления различий по " + engine.Evaluate("colnames(data)[26]").AsCharacter()[0] + " по показателю " + engine.Evaluate("colnames(data)[22]").AsCharacter()[0] + " используем критерий Манна-Уитни.");
                                ResFlag = "";
                                if (engine.Evaluate("wilcox.test(goout~internet, data = data)[3]").AsNumeric()[0] > 0.05) ResFlag = "не ";
                                wordrange.InsertAfter(" Согласно этому критерию " + ResFlag + "выявлены статистически значимые различия между " + engine.Evaluate("levels(data$internet)").AsCharacter()[0] + " и " + engine.Evaluate("levels(data$internet)").AsCharacter()[1] + "");
                                wordrange.InsertAfter(" (p = " + engine.Evaluate("wilcox.test(goout~internet, data = data)").AsCharacter()[2] + ", W = " + engine.Evaluate("wilcox.test(goout~internet, data = data)").AsCharacter()[0] + ").");


                                wordrange.InsertParagraphAfter();

                                wordrange.InsertAfter("Переменная " + engine.Evaluate("colnames(data)[15]").AsCharacter()[0] + " принимает всего " + engine.Evaluate("length(levels(factor(data$failures)))").AsCharacter()[0] + " значений, поэтому для выявления различий по " + engine.Evaluate("colnames(data)[15]").AsCharacter()[0] + " по показателю " + engine.Evaluate("colnames(data)[21]").AsCharacter()[0] + " используем критерий Манна-Уитни.");
                                ResFlag = "";
                                if (engine.Evaluate("wilcox.test(failures~higher, data = data)[3]").AsNumeric()[0] > 0.05) ResFlag = "не ";
                                wordrange.InsertAfter(" Согласно этому критерию " + ResFlag + "выявлены статистически значимые различия между " + engine.Evaluate("levels(data$higher)").AsCharacter()[0] + " и " + engine.Evaluate("levels(data$higher)").AsCharacter()[1] + "");
                                wordrange.InsertAfter(" (p = " + engine.Evaluate("wilcox.test(failures~higher, data = data)").AsCharacter()[2] + ", W = " + engine.Evaluate("wilcox.test(failures~higher, data = data)").AsCharacter()[0] + ").");

                                wordrange.InsertParagraphAfter();

                                wordrange.InsertAfter("Переменная " + engine.Evaluate("colnames(data)[13]").AsCharacter()[0] + " принимает всего " + engine.Evaluate("length(levels(factor(data$traveltime)))").AsCharacter()[0] + " значений, поэтому для выявления различий по " + engine.Evaluate("colnames(data)[13]").AsCharacter()[0] + " по показателю " + engine.Evaluate("colnames(data)[4]").AsCharacter()[0] + " используем критерий Манна-Уитни.");
                                ResFlag = "";
                                if (engine.Evaluate("wilcox.test(traveltime~address, data = data)[3]").AsNumeric()[0] > 0.05) ResFlag = "не ";
                                wordrange.InsertAfter(" Согласно этому критерию " + ResFlag + "выявлены статистически значимые различия между " + engine.Evaluate("levels(data$address)").AsCharacter()[0] + " и " + engine.Evaluate("levels(data$address)").AsCharacter()[1] + "");
                                wordrange.InsertAfter(" (p = " + engine.Evaluate("wilcox.test(traveltime~address, data = data)").AsCharacter()[2] + ", W = " + engine.Evaluate("wilcox.test(traveltime~address, data = data)").AsCharacter()[0] + ").");

                                break;

                            case "checkBox4":

                                wordrange = WordDoc.Bookmarks["Хи_Квадрат"].Range;
                                wordrange.InsertParagraphAfter();
                                string AnsFlag = "";
                                wordrange.InsertAfter("Для выявления взаимосвязей между номинальными переменными используется критерий хи-квадрат.\r");

                                double PValue = engine.Evaluate("chisq.test(data$Dalc, data$famrel)[3]").AsNumeric()[0];
                                double HiValue = engine.Evaluate("chisq.test(data$Dalc, data$famrel)[1]").AsNumeric()[0];
                                if (PValue > 0.05) AnsFlag = "отсутствие";
                                else AnsFlag = "наличие";
                                wordrange.InsertAfter("Так, в рассматриваемых данных показано " + AnsFlag + " статистически значимой взаимосвязи между " + engine.Evaluate("colnames(data)[27]").AsCharacter()[0] + " и " + engine.Evaluate("colnames(data)[24]").AsCharacter()[0] + " (p = " + PValue.ToString() + ", X = " + HiValue.ToString() + ").\r");

                                PValue = engine.Evaluate("chisq.test(data$Dalc, data$studytime)[3]").AsNumeric()[0];
                                HiValue = engine.Evaluate("chisq.test(data$Dalc, data$studytime)[1]").AsNumeric()[0];
                                if (PValue > 0.05) AnsFlag = "отсутствие";
                                else AnsFlag = "наличие";
                                wordrange.InsertAfter("Так, в рассматриваемых данных показано " + AnsFlag + " статистически значимой взаимосвязи между " + engine.Evaluate("colnames(data)[27]").AsCharacter()[0] + " и " + engine.Evaluate("colnames(data)[13]").AsCharacter()[0] + " (p = " + PValue.ToString() + ", X = " + HiValue.ToString() + ").\r");

                                PValue = engine.Evaluate("chisq.test(data$Dalc, data$Walc)[3]").AsNumeric()[0];
                                HiValue = engine.Evaluate("chisq.test(data$Dalc, data$Walc)[1]").AsNumeric()[0];
                                if (PValue > 0.05) AnsFlag = "отсутствие";
                                else AnsFlag = "наличие";
                                wordrange.InsertAfter("Так, в рассматриваемых данных показано " + AnsFlag + " статистически значимой взаимосвязи между " + engine.Evaluate("colnames(data)[27]").AsCharacter()[0] + " и " + engine.Evaluate("colnames(data)[28]").AsCharacter()[0] + " (p = " + PValue.ToString() + ", X = " + HiValue.ToString() + ").\r");

                                PValue = engine.Evaluate("chisq.test(data$Dalc, data$health)[3]").AsNumeric()[0];
                                HiValue = engine.Evaluate("chisq.test(data$Dalc, data$health)[1]").AsNumeric()[0];
                                if (PValue > 0.05) AnsFlag = "отсутствие";
                                else AnsFlag = "наличие";
                                wordrange.InsertAfter("Так, в рассматриваемых данных показано " + AnsFlag + " статистически значимой взаимосвязи между " + engine.Evaluate("colnames(data)[27]").AsCharacter()[0] + " и " + engine.Evaluate("colnames(data)[29]").AsCharacter()[0] + " (p = " + PValue.ToString() + ", X = " + HiValue.ToString() + ").\r");

                                PValue = engine.Evaluate("chisq.test(data$Dalc, data$age)[3]").AsNumeric()[0];
                                HiValue = engine.Evaluate("chisq.test(data$Dalc, data$age)[1]").AsNumeric()[0];
                                if (PValue > 0.05) AnsFlag = "отсутствие";
                                else AnsFlag = "наличие";
                                wordrange.InsertAfter("Так, в рассматриваемых данных показано " + AnsFlag + " статистически значимой взаимосвязи между " + engine.Evaluate("colnames(data)[27]").AsCharacter()[0] + " и " + engine.Evaluate("colnames(data)[3]").AsCharacter()[0] + " (p = " + PValue.ToString() + ", X = " + HiValue.ToString() + ").\r");

                                break;

                            case "checkBox5":

                                wordrange = WordDoc.Bookmarks["Корр_Анализ"].Range;
                                wordrange.InsertAfter("Корреляционный анализ позволяет определить взаимосвязь между метрическими переменными. Значения коэффициентов корреляции представлены в таблице, статистически значимые взаимосвязи выделены полужирным шрифтом.\r");

                                engine.Evaluate("rc <- rcorr(as.matrix(data[,c(3, 30)]))");

                                wordrange = WordDoc.Bookmarks["Корр_Таблица"].Range;
                                Word.Table wordTable1 = WordDoc.Tables.Add(wordrange, 3, 3);

                                wordTable1.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                                wordTable1.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                                wordTable1.Cell(1, 2).Range.Text = "age";
                                wordTable1.Cell(2, 1).Range.Text = "age";
                                wordTable1.Cell(1, 3).Range.Text = "absences";
                                wordTable1.Cell(3, 1).Range.Text = "absences";
                                wordTable1.Cell(3, 2).Range.Text = Math.Round(engine.Evaluate("rc$P[2][1]").AsNumeric()[0], 5).ToString();
                                wordTable1.Cell(2, 3).Range.Text = Math.Round(engine.Evaluate("rc$P[2][1]").AsNumeric()[0], 5).ToString();
                                wordTable1.Cell(3, 3).Range.Text = "NULL";
                                wordTable1.Cell(2, 2).Range.Text = "NULL";

                                for (int i = 1; i < 4; i++)
                                    for (int j = 1; j < 4; j++)
                                    {
                                        wordcellrange = wordTable1.Cell(i, j).Range;
                                        wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                    }

                                wordTable1.Cell(2, 3).Range.Bold = 1;
                                wordTable1.Cell(3, 2).Range.Bold = 1;

                                if (checkBox6.IsChecked == true)
                                {

                                    ExcelBook.Worksheets[2].Name = "Correlation analysis";
                                    for (int i = 1; i < 4; i++)
                                        for (int j = 1; j < 4; j++)
                                        {
                                            if (wordTable1.Cell(i, j).Range.Bold == -1) (ExcelBook.Worksheets[2].Cells(i, j) as Excel.Range).Font.Bold = true;
                                            (ExcelBook.Worksheets[2].Cells(i, j) as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                            ExcelBook.Worksheets[2].Cells[i, j].Value = wordTable1.Cell(i, j).Range.Text.Substring(0, wordTable1.Cell(i, j).Range.Text.Length - 1);
                                            ExcelBook.Worksheets[2].Cells[i, j].EntireColumn.ColumnWidth = 20;
                                        }
                                }

                                wordrange = WordDoc.Bookmarks["Разрыв1"].Range;
                                wordrange.InsertBreak(Word.WdBreakType.wdPageBreak);

                                break;
                            case "checkBox7":

                                wordrange = WordDoc.Bookmarks["Регр_Анализ"].Range;
                                wordrange.InsertAfter("Регрессионный анализ позволяет определить зависимость между absenses и такими переменными, как Dalc, Walc, goout, health, freetime, famrel, romantic, internet, guardian, schoolsup. Значения коэффициентов регрессионного уравнения и уровни значимости представлены в таблице:\r");

                                engine.Evaluate("reg <- lm(data$absences ~ data$Dalc+data$Walc+data$goout+data$health+data$freetime+data$famrel+data$romantic+data$internet+data$guardian+data$schoolsup)");
                                CharacterMatrix RegressionMatrix = engine.Evaluate("summary(reg)$coefficients").AsCharacterMatrix();

                                wordrange = WordDoc.Bookmarks["Регр_Анализ_Таблица"].Range;
                                Word.Table wordTable2 = WordDoc.Tables.Add(wordrange, 13, 5);

                                wordTable2.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                                wordTable2.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                                string[] columnRegrNames = new string[4] { "Estimate", "Std. Error", "t Value", "Pr(>|t|)" };
                                string[] rowRegrNames = new string[12] { "Intercept", "Dalc", "Walc", "goout", "health", "freetime", "famrel", "romanticyes", "internetyes", "guardianmother", "guardianother", "schoolsupyes"};

                                for (int i = 2; i < 6; i++)
                                {
                                    wordTable2.Cell(1, i).Range.Text = columnRegrNames[i - 2];
                                    wordTable2.Cell(1, i).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                }

                                for (int i = 2; i < 14; i++)
                                {
                                    wordTable2.Cell(i, 1).Range.Text = rowRegrNames[i - 2];
                                    wordTable2.Cell(i, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                }

                                for (int j = 2; j < 6; j++)
                                    for(int i = 2; i < 14; i++)
                                    {
                                        wordTable2.Cell(i, j).Range.Text = Math.Round(double.Parse(RegressionMatrix[i - 2, j - 2]),5).ToString();
                                        wordTable2.Cell(i, j).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                        if (j == 5 && double.Parse(RegressionMatrix[i - 2, j - 2]) <= 0.05)
                                        {
                                            wordTable2.Cell(i, j).Range.Bold = 1;
                                            wordTable2.Cell(i, 1).Range.Bold = 1;
                                        }
                                    }

                                if (checkBox6.IsChecked == true)
                                {

                                    ExcelBook.Worksheets[3].Name = "Regression analysis";
                                    for (int j = 1; j < 6; j++)
                                        for (int i = 1; i < 14; i++)
                                        {
                                            if (wordTable2.Cell(i, j).Range.Bold == -1) (ExcelBook.Worksheets[3].Cells(i, j) as Excel.Range).Font.Bold = true;
                                            (ExcelBook.Worksheets[3].Cells(i, j) as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                            ExcelBook.Worksheets[3].Cells[i, j].Value = wordTable2.Cell(i, j).Range.Text.Substring(0, wordTable2.Cell(i, j).Range.Text.Length - 1);
                                            ExcelBook.Worksheets[3].Cells[i, j].EntireColumn.ColumnWidth = 20;
                                            
                                        }
                                }

                                break;

                            default:
                                break;
                        }
                    }
                }
            }
        }

    }
}
