using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ClosedXML.Excel;

namespace ConverT
{
    class Program
    {
        [DllImport("kernel32.dll")]
        static extern uint GetPrivateProfileString(
        string lpAppName,
        string lpKeyName,
        string lpDefault,
        StringBuilder lpReturnedString,
        uint nSize,
        string lpFileName);

        static void Main(string[] args)
        {
            WriteLineColor("Приложение запущено.", "Yellow");

            analysis_directory();
        }

        static int total_file = 0;
        static int failed = 0;
        static int total_page = 0;
        static int con_true = 0;
        static int already_con = 0;
        static int read_only = 0;

        static Boolean backup_data(string soure,string name)
        {
            try
            {
                string EntryDate = DateTime.Today.ToShortDateString().Replace(".","_");

                if (!Directory.Exists(Environment.CurrentDirectory + "//backup" + "//" + EntryDate))
                    Directory.CreateDirectory(Environment.CurrentDirectory + "//backup" + "//" + EntryDate);

                if (File.Exists(Environment.CurrentDirectory + "//backup//" + EntryDate + "//" + "backup_" + name))
                    File.Delete(Environment.CurrentDirectory + "//backup//" + EntryDate + "//" + "backup_" + name);
                else
                    File.Copy(soure + "//" + name, Environment.CurrentDirectory + "//backup//" + EntryDate + "//" + "backup_" + name);
            }
            catch (System.Exception ex)
            {
                WriteLineColor(ex.Message, "Red");
                Log.log_write(ex.Message, "Exception", "Exception");
                Console.ReadKey();
                return false;
            }

            return true;
        }

        static void analysis_directory()
        {
            try
            {
                string searchPattern = "*.xlsx";

                DirectoryInfo di = new DirectoryInfo(Environment.CurrentDirectory);

                FileInfo[] files =
                    di.GetFiles(searchPattern, SearchOption.AllDirectories);

                foreach (FileInfo file in files)
                {
                    if (!file.Name.Contains("backup_"))
                    {
                        total_file++;

                        WriteLineColor("\n===================================================================", "Magenta");
                        WriteLineColor("Найденый файл: " + file.Name, "Yellow");
                        WriteLineColor("Путь к файлу: " + file.DirectoryName + "\\" + "\n", "Green");

                        int ro = 0;

                        while (file.IsReadOnly && ro < 2)
                        {
                            WriteLineColor("===================================================================", "Magenta");
                            WriteLineColor("Внимание файл доступен только для чтения!", "Red");
                            WriteLineColor("Возможно установлен флаг только чтение!", "Red");
                            WriteLineColor("Закройте все приложения блокирующие доступ к файлу и нажмите клавишу!\n", "Red");
                            WriteLineColor("Количество попыток разблокировать: " + (2-ro).ToString() + "\n", "Yellow");
                            WriteLineColor("===================================================================", "Magenta");
                            Console.ReadKey();
                            ro++;
                        }

                        if (file.IsReadOnly)
                        {
                            WriteLineColor("Файл остался заблокированым!", "Cyan");
                            read_only++;
                        }
                        else
                        {
                            if (file_processing(file.DirectoryName, file.Name))
                            {
                                if (file.Name.Contains("c") || file.Name.Contains("C"))
                                {
                                    WriteLineColor("\n" + file.Name + " содержит символ очистки количества и сумм бланка", "Yellow");
                                    WriteLineColor("Очистка... ", "Cyan");
                                    //todo function clean all numbers
                                }
                                WriteLineColor("Успешно!", "Cyan");
                                con_true++;
                            }
                            else
                            {
                                WriteLineColor("Отказ!", "Cyan");
                            }
                            WriteLineColor("===================================================================", "Magenta");
                        }
                    }
                }

                progress();
                Console.ReadKey();
            }
            catch (System.Exception ex)
            {
                WriteLineColor(ex.Message,"Red");
                Log.log_write(ex.Message, "Exception", "Exception");
                Console.ReadKey();
            }
        }

        static Boolean file_processing(string file_patch, string file_name)
        {
            var wb = new XLWorkbook(file_patch + "\\" + file_name);

            var ws = wb.Worksheets.Worksheet(1);

            if (ws.Cell(16, 1).Value.ToString() == "Инвентаризационная опись")
            {
                WriteLineColor("Файл " + file_name + " уже конвертирован!", "Red");
                Log.log_write("Файл " + file_name + " уже конвертирован!", "WARNING", "warning");
                already_con++;
                return false;
            }

            if (ws.Cell(2, 1).Value.ToString() != "Инвентаризационная опись")
            {
                WriteLineColor("Файл " + file_name + " формат не подходит!", "Red");
                Log.log_write("Файл " + file_name + " формат не подходит!", "ERROR", "warning");
                failed++;
                return false;
            }

            if (backup_data(file_patch, file_name))
            {
                
                WriteLineColor("Создана копия оригинального файла.", "Magenta");
            }
            else
            {
                WriteLineColor("Внимание копия файла не сделана!", "Red");
                Console.ReadKey();
            }

            WriteLineColor("Обработка...", "Cyan");

            //Удаляем ненужные колонки

            ws.Column(9).Delete();
            ws.Range("A1:I5").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            ws.Range("A7:I11").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            ws.Range("A11:I15").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            ws.Range("A15:I17").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            //Прячем ненужные строки
            int i = 1;

            while (i <= 15)
            {
                ws.Row(i).Hide();
                i++;
            }

            //Объединение ячеек 
            ws.Range("A16:G16").Row(1).Merge();
            ws.Range("A17:G17").Row(1).Merge();


            ws.Cell(16, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(16, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell(16, 1).Value = "Инвентаризационная опись";

            ws.Cell(17, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(17, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell(17, 1).Value = "товарно-материальных ценностей";

            int total_cell = 16;
            int total_coll = 1;

            while (total_coll <= 7)
            {
                total_cell = 16;

                while (ws.Cell(total_cell, 1).GetString() != "")
                {
                    ws.Cell(total_cell, total_coll).Style.Font.FontName = "Arial";

                    ws.Cell(total_cell, total_coll).Style.Font.FontSize = 9;

                    total_cell++;
                }

                total_coll++;
            }

            int cel1 = total_cell + 1;
            int cel3 = total_cell + 3;

            //Добавление в конец файла строк
            ws.Cell(cel1, 1).Value = "Материально-ответственное(ые) лицо(а) :";
            ws.Cell(cel1, 1).Style.Font.FontName = "Arial";
            ws.Cell(cel1, 1).Style.Font.FontSize = 9;
            ws.Range("A" + cel1 + ":G"+ cel1).Row(1).Merge();

            ws.Cell(cel3, 1).Value = "Начальник комиссии :";
            ws.Cell(cel3, 1).Style.Font.FontName = "Arial";
            ws.Cell(cel3, 1).Style.Font.FontSize = 9;
            ws.Range("A" + cel3 + ":G" + cel3).Row(1).Merge();


            //ширина колонок штрихкод,наименование,сумма.
            ws.Column(2).Width = 13;
            ws.Column(3).Width = 52;
            ws.Column(7).Width = 9;

            //считаем количество страниц на печать
            int pages = (total_cell+4) / 54;

            //минимальное количество страниц отправляемых на печать.
            if (pages == 0)
                pages = 1;

            total_page += pages;

            //устанавливаем параметры страниц печати (количество страниц в ширину,количество страниц в высоту)
            ws.PageSetup.FitToPages(1, pages);

            WriteLineColor("Всего строк: " + total_cell.ToString() + "  Всего колонок: " + total_coll.ToString() + "  Всего страниц на печать: " + pages.ToString(), "Cyan");

            wb.Save();

            return true;
        }

        static void WriteLineColor(string value, string color)
        {
            if (color == "Red")
                Console.ForegroundColor = ConsoleColor.Red;
            else if (color == "Green")
                Console.ForegroundColor = ConsoleColor.Green;
            else if (color == "Magenta")
                Console.ForegroundColor = ConsoleColor.Magenta;
            else if (color == "Yellow")
                Console.ForegroundColor = ConsoleColor.Yellow;
            else if (color == "Cyan")
                Console.ForegroundColor = ConsoleColor.Cyan;

            Console.WriteLine(value.PadRight(Console.WindowWidth - 1)); // <-- see note

            Console.ResetColor();
        }

        static void progress()
        {
            WriteLineColor("\n", "Red");
            WriteLineColor("\n", "Red");
            WriteLineColor("Всего файлов: " + total_file,"Green");
            WriteLineColor("Конвертировано: " + con_true, "Green");
            WriteLineColor("Уже Конвертированы: " + already_con, "Red");
            WriteLineColor("ReadOnly: " + read_only, "Red");
            WriteLineColor("Отказ: " + failed, "Red");
            WriteLineColor("\n", "Red");
            WriteLineColor("Понадобиться страниц на печать: " + total_page, "Green");
        }
    }
}
