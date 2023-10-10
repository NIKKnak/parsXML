using static System.Windows.Forms.LinkLabel;
using OfficeOpenXml;
using System.Text;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
namespace parsXML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Метод для выбора исходного файла через диалоговое окно
            string ChooseFile()
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }

                return null;
            }


            // Выбор исходного файла
            string inputFile = ChooseFile();

            if (inputFile != null)
            {
                // Чтение файла
                List<string> lines = new List<string>();
                using (StreamReader sr = new StreamReader(inputFile))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        lines.Add(line);
                    }
                }

                // Создание нового Excel файла
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                // Заголовки столбцов
                worksheet.Cells[1, 1] = "Client ID";
                worksheet.Cells[1, 2] = "Name";
                worksheet.Cells[1, 3] = "User ID";

                int row = 2; // Строка, с которой начинается запись значений

                // Поиск строк, содержащих поля
                foreach (string line in lines)
                {
                    if (line.Contains("client id=") && line.Contains("name=") && line.Contains("user id="))
                    {
                        // Извлечение значений полей
                        string clientId = GetValue(line, "client id=");
                        string name = GetValue(line, "name=");
                        string userId = GetValue(line, "user id=");

                        // Запись значений в ячейки
                        worksheet.Cells[row, 1] = clientId;
                        worksheet.Cells[row, 2] = name;
                        worksheet.Cells[row, 3] = userId;

                        row++;
                    }
                }

                // Сохранение файла
                workbook.SaveAs("output.xlsx");
                workbook.Close();
                excelApp.Quit();
            }
        }

        // Метод для извлечения значения поля из строки
        static string GetValue(string line, string field)
        {
            int startIndex = line.IndexOf(field) + field.Length;
            int endIndex = line.IndexOf(" ", startIndex);
            return line.Substring(startIndex, endIndex - startIndex);
        }

        
        













    }



        }
