using System;
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

using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;


namespace DataAnalysisApp
{
    public partial class MainWindow : Window
    {
        //Переменные для отчетов
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
            //engine.Evaluate("library(foreign)");
            //engine.Evaluate("library(Hmisc)");
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
            wordrange.Text = textBox.Text;
            wordrange = WordDoc.Bookmarks["Автор"].Range;
            wordrange.Text = textBox1.Text;
            wordrange = WordDoc.Bookmarks["Дата"].Range;
            wordrange.Text = dateTimePicker1.SelectedDate.ToString();
            CheckBox[] CheckboxArr = new CheckBox[6] { checkBox, checkBox1, checkBox22, checkBox3, checkBox4, checkBox5, };
            Word.Range wordcellrange;
            CharacterVector CountOfRows = engine.Evaluate("nrow(data)").AsCharacter();

            foreach (CheckBox CurrentCheckbox in CheckboxArr)
            {
                if (CurrentCheckbox.IsChecked == true)
                {
                    switch (CurrentCheckbox.Name)
                    {
                        case "checkBox":

                            wordrange = WordDoc.Bookmarks["Опис_Таблица"].Range;
                            Word.Table wordTable = WordDoc.Tables.Add(wordrange, 4, 5);

                            wordTable = WordDoc.Tables[1];

                            WordDoc.Tables[1].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            WordDoc.Tables[1].Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                            string[] columnNames = new string[5] { "Переменная", "Минимум", "Среднее", "Медиана", "Максимум" } ;


                            for (int i = 1; i < 6; i++) 
                                {
                                    wordcellrange = WordDoc.Tables[1].Cell(1, i).Range;
                                    wordcellrange.Text = columnNames[i - 1];
                                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }

                            string[] rowNames = new string[3] { "Возраст", "Учебное время", "Прогулы"};

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
                                switch(i)
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
                            break;

                        case "checkBox1":

                            engine.Evaluate("data <- as.data.frame(data)");
                            wordrange = WordDoc.Bookmarks["Т_Тест"].Range;
                            wordrange.InsertAfter("Рассмотрим различия по кол-ву лет обучения, зарплате и тем, была ли установлена зарплата путем коллективных переговоров.\r");
                            double TTestPValue = engine.Evaluate("t.test(data[,3] ~ data$union)$p.value").AsNumeric()[0];
                            double TTestTValue = engine.Evaluate("t.test(data[,3] ~ data$union)$statistic").AsNumeric()[0];
                            if (TTestPValue <= 0.05) wordrange.InsertAfter("Согласно критерию Стьюлента выявлены статистически значимые различия между Schooling и Union (p.value =  " + Math.Round(TTestPValue, 5) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");
                            else wordrange.InsertAfter("Согласно критерию Стьюлента не выявлены статистически значимые различия между Schooling и Union (p.value =  " + Math.Round(TTestPValue, 5) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");
                            TTestPValue = engine.Evaluate("t.test(data[,9] ~ data$union)$p.value").AsNumeric()[0];
                            TTestTValue = engine.Evaluate("t.test(data[,9] ~ data$union)$statistic").AsNumeric()[0];
                            if (TTestPValue <= 0.05) wordrange.InsertAfter("Согласно критерию Стьюлента выявлены статистически значимые различия между Wage и Union (p.value =  " + Math.Round(TTestPValue, 15) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");
                            else wordrange.InsertAfter("Согласно критерию Стьюлента не выявлены статистически значимые различия между Wage и Union (p.value =  " + Math.Round(TTestPValue, 15) + ", t = " + Math.Round(TTestTValue, 5) + ")\r");
                                  
                            break;

                        case "checkbox22":

                            wordrange = WordDoc.Bookmarks["Анова"].Range;
                            wordrange.InsertParagraphAfter();
                            wordrange.InsertAfter("С помощью теста ANOVA были изучены различия для групп по расе (black - темнокожие, hisp - латиноамериканцы, other - остальные) для года рождения, опыта и зарплаты.\r");
                            wordrange.InsertAfter("Year: p value = " + Math.Round(engine.Evaluate("aov(data$ethn ~ data$year)[[1]][[2]]").AsNumeric()[0], 5) + "; F = " +
                                Math.Round(engine.Evaluate("aov(data$ethn ~ data$year)[[1]][[1]]").AsNumeric()[0], 5));
                            wordrange.InsertAfter("Year: p value = " + Math.Round(engine.Evaluate("aov(data$ethn ~ data$exper)[[1]][[2]]").AsNumeric()[0], 5) + "; F = " +
                                Math.Round(engine.Evaluate("aov(data$ethn ~ data$exper)[[1]][[1]]").AsNumeric()[0], 5));
                            wordrange.InsertAfter("Year: p value = " + Math.Round(engine.Evaluate("aov(data$ethn ~ data$wage)[[1]][[2]]").AsNumeric()[0], 5) + "; F = " +
                                Math.Round(engine.Evaluate("aov(data$ethn ~ data$wage)[[1]][[1]]").AsNumeric()[0], 5));
                            break;

                        case "checkbox3":

                            //engine.Evaluate("cor(data[,c(2, 3, 4, 9)])");
                            //engine.Evaluate("rc <- rcorr(as.matrix(data[,c(2, 3, 4, 9)]))");
                            //NumericMatrix CorrelationMatrix = engine.Evaluate("rc").AsNumericMatrix();

                            //wordrange = WordDoc.Bookmarks["aaf"].Range;
                            //wordrange.InsertAfter("Year");

                            //wordrange = WordDoc.Bookmarks["Корр_Таблица_13"].Range;
                            //wordrange.InsertAfter("School");

                            //wordrange = WordDoc.Bookmarks["Корр_Таблица_14"].Range;
                            //wordrange.InsertAfter("Exper");

                            //wordrange = WordDoc.Bookmarks["Корр_Таблица_15"].Range;
                            //wordrange.InsertAfter("Wage");

                            //wordrange = WordDoc.Bookmarks["Корр_Таблица_21"].Range;
                            //wordrange.InsertAfter("Year");

                            //wordrange = WordDoc.Bookmarks["Корр_Таблица_22"].Range;
                            //wordrange.InsertAfter(CorrelationMatrix[2, 2].ToString());

                            //wordrange = WordDoc.Bookmarks["Корр_Таблица_23"].Range;
                            //wordrange.InsertAfter(CorrelationMatrix[2, 3].ToString());

                            //wordrange = WordDoc.Bookmarks["Корр_Таблица_24"].Range;
                            //wordrange.InsertAfter(CorrelationMatrix[2, 4].ToString());
                            //for (int i = 0; i <= 3; i++)
                            //{
                            //    for (int j = 0; j <= 3; j++)
                            //    {
                            //        if (CorrelationMatrix[i, j] <= 0.05)
                            //        {
                            //            WordDoc.Tables[1].Cell(i + 2, j + 2).Range.Bold = 1;
                            //        }
                            //        wordTable.Cell(i + 2, j + 2).Range.Text = Math.Round(CorrelationMatrix[i, j], 5).ToString();
                            //    }
                            //}

                            break;
                        
                        default:
                            break;
                    }
                }
            }
        }

    }
}
