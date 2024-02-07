using System;
using System.Windows.Forms;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;

namespace InstructionsParser_v2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox1.Text = Properties.Settings.Default.old_path;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            button1.Text = "Обработка...";
            button1.Enabled = false;

            char[] charStr = textBox1.Text.ToCharArray();

            string tx_path = textBox1.Text.ToString();
            var tx_path_files = Directory.GetFiles(tx_path);
            List<string> paths = new List<string>();
            foreach (var path in tx_path_files)
            {
                if (path.Contains(".doc"))
                    paths.Add(path);

            }

            if (paths.Count == 0)
            {
                MessageBox.Show("Файлы отсутсвуют");
                button1.Text = "Запустить";
                button1.Enabled = true;
                return;
            }

            //Настройка ProgressBar
            progressBar1.Minimum = 0;
            progressBar1.Maximum = paths.Count;
            progressBar1.Step = 1;

            //Объявляем приложения
            Excel.Application excel = new Excel.Application
            {
                Visible = false,
                //Количество листов в рабочей книге
                SheetsInNewWorkbook = 1
            };
            Word.Application WordApp = new Word.Application();

            //Добавить рабочую книгу
            Excel.Workbook workBook = excel.Workbooks.Add(Type.Missing);
            //Отключить отображение окон с сообщениями
            excel.DisplayAlerts = false;
            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet sheet = (Excel.Worksheet)excel.Worksheets.get_Item(1);
            //Название листа (вкладки снизу)
            //sheet.Name = "Лист1";

            //Пример заполнения ячеек
            sheet.Range["A1"].Value = "№";
            sheet.Range["B1"].Value = "Инструкция";
            sheet.Range["C1"].Value = "Фирма-производитель";
            Excel.Range r1 = sheet.Cells[1, 1];
            Excel.Range r2 = sheet.Cells[1, 3];
            Excel.Range range = sheet.get_Range(r1, r2);
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

            //увеличиваем размер по ширине диапазон ячеек
            Excel.Range range1 = sheet.get_Range("A:A");
            Excel.Range range2 = sheet.get_Range("B:B");
            Excel.Range range3 = sheet.get_Range("C:C");
            range1.EntireColumn.ColumnWidth = 8;
            range2.EntireColumn.ColumnWidth = 50;
            range3.EntireColumn.ColumnWidth = 150;
            range1.EntireRow.WrapText = false;
            range2.EntireRow.WrapText = false;
            range3.EntireRow.WrapText = true;
            range1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            range2.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range2.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            range3.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            //Перебор файлов
            int row_excel = 2;
            foreach (var path in paths)
            {
                try
                {
                    Word.Document doc = WordApp.Documents.Open(path, ReadOnly: true, Visible: false);
                    Word.ContentControls contentControls = null;
                    Word.ContentControl contentControl = null;

                    string controlsList = string.Empty;

                    doc.Activate();
                    doc = WordApp.ActiveDocument;
                    contentControls = doc.ContentControls;
                    for (int i = 1; i <= contentControls.Count; i++)
                    {
                        contentControl = contentControls[i];
                        controlsList += String.Format("{0} : {1}{2}",
                            contentControl.Title, contentControl.Type, Environment.NewLine);

                        if (contentControl.Title.Contains("Производитель"))
                        {
                            sheet.Range["A" + (row_excel).ToString()].Value = row_excel - 1;
                            sheet.Range["B" + (row_excel).ToString()].Value = Path.GetFileName(path);
                            sheet.Range["C" + (row_excel).ToString()].Value = contentControl.Range.Text.Replace("\r", "\n");
                            r1 = sheet.Cells[1, 1];
                            r2 = sheet.Cells[row_excel, 3];
                            sheet.get_Range(r1, r2).Borders.Color = ColorTranslator.ToOle(Color.Black);
                            row_excel += 1;
                            Console.WriteLine(Path.GetFileName(path));
                            progressBar1.PerformStep();
                        }
                    }

                    for (int i = 1; i <= contentControls.Count; i++)
                    {

                        contentControl = contentControls[i];
                        controlsList += String.Format("{0} : {1}{2}",
                            contentControl.Title, contentControl.Type, Environment.NewLine);
                    }
                    doc.Close();
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    Console.WriteLine(ex);
                }
                finally
                {

                }
            }
            Console.WriteLine();
            excel.Application.ActiveWorkbook.SaveAs(tx_path + @"\word_to_excel_Список_фирм.xlsx", Type.Missing,
  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            progressBar1.Value = paths.Count;
            MessageBox.Show("Готово. Файл сохранен по пути" + tx_path + "word_to_excel_Список_фирм.xlsx");
            progressBar1.Value = 0;
            button1.Text = "Запустить";
            button1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }
    }
}