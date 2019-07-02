using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel =  Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace re
{
    public partial class Form1 : Form
    {
        //Пути к файлам
        static string pathTemplate = "";
        static string pathTable = "";
        static string pathSave = "";

        //Свойства кнопки "Сгенерировать"
        private void Properties(Button button3)
        {
            button3.Enabled = true;
            button3.BackColor = Color.FromArgb(144, 238, 144);
            button3.ForeColor = Color.FromArgb(0, 0, 0);
        }

        public Form1()
        {
            InitializeComponent();
        }

        //Получить путь до файла эксель
        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Файлы xlsx|*.xlsx";
            ofd.ShowDialog();

            pathTable = ofd.FileName;
            //Если все пути получены: включаем кнопку "Сгенерировать"
            if (pathTable != "" && pathTemplate != "" && pathSave != "")
            {
                Properties(button3);
            }
        }

        //Получить путь до файла ворд
        private void Button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Файлы docx|*.docx|Файлы doc|*.doc";
            ofd.ShowDialog();

            pathTemplate = ofd.FileName;
            if (pathTable != "" && pathTemplate != "" && pathSave != "")
            {
                Properties(button3);
            }
        }

        //Получить директорию для сохранения писем
        private void Button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowDialog();

            pathSave = fbd.SelectedPath;
            if (pathTable != "" && pathTemplate != "" && pathSave != "")
            {
                Properties(button3);
            }
        }

        //Кнопка "Сгенерировать"
        private void Button3_Click(object sender, EventArgs e)
        {
            var WordApp = new Word.Application(); 
            WordApp.Visible = false;

            //Открыть ворд && открыть эксель, получить 1 лист
            var WordDocument = WordApp.Documents.Open(@"" + pathTemplate);
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"" + pathTable);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            //Получить кол-во заполненных ячеек эксель
            int nInLastRow = ObjWorkSheet.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            int nInLastCol = ObjWorkSheet.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            //Св-ва прогрессбара и лейбла
            progressBar1.Visible = true;
            label1.Visible = false;
            label1.Text = "";
            progressBar1.Value = 0;
            progressBar1.Step = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = nInLastRow;

            string[,] list = new string[nInLastRow, nInLastCol]; //Равен по размеру листу

            //Данные с листа в массив
            for (int i = 0; i < nInLastRow; i++)
            {
                for (int j = 0; j < nInLastCol; j++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString();

                }
            }

            //Данные из массива в ворд
            try
            {
                for (int i = 0; i < nInLastRow; i++)
                {
                    for (int j = 0; j < nInLastCol; j++)
                    {

                        string tmp = list[i, j]; //Значение текущей ячейки в переменную
                        ReplaceWordStub("{" + j + "}", tmp, WordDocument); //Меняем метку в шаблоне ворд на значение

                        //Сохранить, когда будет обрабатываться последняя ячейка в строке
                        if (j == nInLastCol-1)
                        {
                            WordDocument.SaveAs(@"" + pathSave + "\\" + list[i, j] + ".docx");
                            //Закрыть, иначе будет множество экземпляров ворд в фоне, открыть шаблон снова для следующей строки
                            WordDocument.Close(false, Type.Missing, Type.Missing);
                            WordDocument = WordApp.Documents.Open(@"" + pathTemplate);
                        }

                    }
                    
                    //Показать прогресс
                    progressBar1.Value += 1;

                    if (progressBar1.Value == nInLastRow)
                    {
                        progressBar1.Visible = false;
                        label1.Visible = true;
                        label1.Text = "Complete!";
                    }
                }
            }
            catch (Exception ex)
            {
                label1.Text = ex.ToString();
            }
            finally
            {
                //Закрыть после завершения работы
                WordDocument.Close(false, Type.Missing, Type.Missing);
                WordApp.Quit();
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
            }
        }

        //Метод, который меняет метку и данные из ячейки местами
        static void ReplaceWordStub(string StubToReplace, string Text, Word.Document WordDocument)
        {

            var Range = WordDocument.Content;
            Range.Find.Execute(FindText: StubToReplace, ReplaceWith: Text);

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.Collect(); // Убрать за собой
        }
    }
}