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
            engine.Evaluate("library(Ecdat)");
            engine.Evaluate("library(Hmisc)");
            engine.Evaluate("data <- data.frame(Males)");

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
            CheckBox[] CheckboxArr = new CheckBox[6] { checkBox, checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, };
            CharacterVector CountOfRows = engine.Evaluate("nrow(data)").AsCharacter();

            foreach (CheckBox CurrentCheckbox in CheckboxArr)
            {
                if (CurrentCheckbox.IsChecked == true)
                {
                    switch (CurrentCheckbox.Name)
                    {
                        case "checkBox":

                            CharacterVector YearSummary = engine.Evaluate("summary(data$year)").AsCharacter();

                            wordrange = WordDoc.Bookmarks["Опис_Год"].Range;
                            wordrange.InsertAfter("Year");

                            wordrange = WordDoc.Bookmarks["Опис_Год_Минимум"].Range;
                            wordrange.InsertAfter(YearSummary[0]);

                            wordrange = WordDoc.Bookmarks["Опис_Год_Среднее"].Range;
                            wordrange.InsertAfter(Math.Round(Convert.ToDouble(YearSummary[3])).ToString());

                            wordrange = WordDoc.Bookmarks["Опис_Год_Медиана"].Range;
                            wordrange.InsertAfter(Math.Round(Convert.ToDouble(YearSummary[2])).ToString());

                            wordrange = WordDoc.Bookmarks["Опис_Год_Максимум"].Range;
                            wordrange.InsertAfter(YearSummary[5]);

                            CharacterVector SchoolYears = engine.Evaluate("summary(data$school)").AsCharacter();

                            wordrange = WordDoc.Bookmarks["Опис_Обучение"].Range;
                            wordrange.InsertAfter("School");

                            wordrange = WordDoc.Bookmarks["Опис_Обучение_Минимум"].Range;
                            wordrange.InsertAfter(SchoolYears[0]);

                            wordrange = WordDoc.Bookmarks["Опис_Обучение_Среднее"].Range;
                            wordrange.InsertAfter(Math.Round(Convert.ToDouble(SchoolYears[3])).ToString());

                            wordrange = WordDoc.Bookmarks["Опис_Обучение_Медиана"].Range;
                            wordrange.InsertAfter(Math.Round(Convert.ToDouble(SchoolYears[2])).ToString());

                            wordrange = WordDoc.Bookmarks["Опис_Обучение_Максимум"].Range;
                            wordrange.InsertAfter(SchoolYears[5]);

                            CharacterVector Experience = engine.Evaluate("summary(data$exper)").AsCharacter();

                            wordrange = WordDoc.Bookmarks["Опис_Опыт"].Range;
                            wordrange.InsertAfter("Experience");

                            wordrange = WordDoc.Bookmarks["Опис_Опыт_Минимум"].Range;
                            wordrange.InsertAfter(Experience[0]);

                            wordrange = WordDoc.Bookmarks["Опис_Опыт_Среднее"].Range;
                            wordrange.InsertAfter(Math.Round(Convert.ToDouble(Experience[3])).ToString());

                            wordrange = WordDoc.Bookmarks["Опис_Опыт_Медиана"].Range;
                            wordrange.InsertAfter(Math.Round(Convert.ToDouble(Experience[2])).ToString());

                            wordrange = WordDoc.Bookmarks["Опис_Опыт_Максимум"].Range;
                            wordrange.InsertAfter(Experience[5]);

                            wordrange = WordDoc.Bookmarks["Опис_Опыт_Строка"].Range;

                            double InterquartileRange = double.Parse(Experience[4]) - double.Parse(Experience[1]);
                            double[] InternalBoundary = new double[2];
                            InternalBoundary[0] = InterquartileRange * 1.5 - double.Parse(Experience[0]);
                            InternalBoundary[1] = InterquartileRange * 1.5 + double.Parse(Experience[4]);

                            double[] OuterBoundary = new double[2];
                            OuterBoundary[0] = InterquartileRange * 3.0 - double.Parse(Experience[0]);
                            OuterBoundary[1] = InterquartileRange * 3.0 + double.Parse(Experience[4]);

                            engine.Evaluate("expValues <- data.frame(data$exper)");
                            CharacterVector ExperienceValues = engine.Evaluate("data$exper").AsCharacter();

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

                            wordrange.InsertAfter("Так, Experience (опыт работы) изменяется от " + Experience[0] + " до " + Experience[5] + " со средним значением равным " + Math.Round(Convert.ToDouble(Experience[3])).ToString() + " и медианой равной " + Math.Round(Convert.ToDouble(Experience[2])).ToString() + ", что говорит о " + FlagStr);

                            wordrange = WordDoc.Bookmarks["Опис_Женатые"].Range;
                            engine.Evaluate("married <- summary(data$maried)");
                            CharacterVector Married = engine.Evaluate("married").AsCharacter();
                            double ProcentOfMarried = Math.Round(double.Parse(Married[1]) / double.Parse(CountOfRows[0]) * 100);                            
                            wordrange.InsertAfter("Married (женатые) составляет " + Convert.ToInt32(ProcentOfMarried).ToString() + "% выборки.");

                            wordrange = WordDoc.Bookmarks["Опис_Жительство"].Range;
                            engine.Evaluate("residence <- summary(data$residence)");
                            CharacterVector Residence = engine.Evaluate("residence").AsCharacter();
                            double[] ResidenceProcent = new double[5];
                            for (int i = 0; i < 5; i++)
                            {
                                ResidenceProcent[i] = Math.Round(double.Parse(Residence[i]) / double.Parse(CountOfRows[0]) * 100);
                            }

                            wordrange.InsertAfter("По показателю Residense (место проживания) выборка распределена следующим образом: rural_area -- " + ResidenceProcent[0].ToString() + "%; north_east -- " + ResidenceProcent[1].ToString() + "%; nothern_central -- " + ResidenceProcent[2].ToString() + "%; south -- " + ResidenceProcent[3].ToString() + "%; NA --  " + ResidenceProcent[4].ToString() + "%.");

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

                        default:
                            break;
                    }
                }
            }
        }

    }
}
