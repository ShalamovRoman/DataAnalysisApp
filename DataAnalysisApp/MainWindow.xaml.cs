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
            engine.Evaluate("data <- data.frame(Males)");

            object begin = 0;
            object end = 0;
            WordApp = new Word.Application();
            WordApp.Visible = true;
            string ReportPath = @"C:\DataAnalysisApp\report.docx";
            WordDoc = WordApp.Documents.Add(ReportPath);
            Word.Range wordrange = WordDoc.Range(ref begin, ref end);
            wordrange = WordDoc.Bookmarks["Отчет"].Range;
            wordrange.Text = textBox.Text;
            wordrange = WordDoc.Bookmarks["Автор"].Range;
            wordrange.Text = textBox1.Text;
            wordrange = WordDoc.Bookmarks["Дата"].Range;
            wordrange.Text = dateTimePicker1.SelectedDate.ToString();
            CheckBox[] CheckboxArr = new CheckBox[6] { checkBox, checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, };

            foreach (CheckBox CurrentCheckbox in CheckboxArr)
            {
                if (CurrentCheckbox.IsChecked == true)
                {
                    switch(CurrentCheckbox.Name)
                    {
                        case "checkBox":
                            wordrange = WordDoc.Bookmarks["описстат"].Range;
                            wordrange.InsertParagraphAfter();
                            wordrange.InsertAfter("Описательные статистики\r" + "Описание данных...:\r" + "...для метрических:\r" + "Год рождения:\r");

                            CharacterVector str = engine.Evaluate("summary(data$year)").AsCharacter();
                            wordrange.InsertAfter("Минимум = " + str[0] + "; ");

                            wordrange.InsertAfter("1 квартиль = " + str[1] + "; ");
                            wordrange.InsertAfter("Медиана = " + str[2] + "; ");
                            wordrange.InsertAfter("Среднее = " + str[3] + "; ");
                            wordrange.InsertAfter("3 квартиль = " + str[4] + "; ");
                            wordrange.InsertAfter("Максимум = " + str[5] + "\r");

                            str = engine.Evaluate("summary(data$school)").AsCharacter();

                            wordrange.InsertAfter("Время обучения:\rМинимум = " + str[0] + "; ");
                            wordrange.InsertAfter("1 квартиль = " + str[1] + "; ");
                            wordrange.InsertAfter("Медиана = " + str[2] + "; ");
                            wordrange.InsertAfter("Среднее = " + str[3] + "; ");
                            wordrange.InsertAfter("3 квартиль = " + str[4] + "; ");
                            wordrange.InsertAfter("Максимум = " + str[5] + "\r");

                            str = engine.Evaluate("summary(data$exper)").AsCharacter();

                            wordrange.InsertAfter("Опыт:\rМинимум = " + str[0] + "; ");
                            wordrange.InsertAfter("1 квартиль = " + str[1] + "; ");
                            wordrange.InsertAfter("Медиана = " + str[2] + "; ");
                            wordrange.InsertAfter("Среднее = " + str[3] + "; ");
                            wordrange.InsertAfter("3 квартиль = " + str[4] + "; ");
                            wordrange.InsertAfter("Максимум = " + str[5] + "\r");

                            break;

                        default:
                            break;
                    }
                }
            }
        }

    }
}
