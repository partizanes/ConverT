using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using System.Threading;

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

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        static IntPtr ConsoleHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

        private const int SW_MINIMIZE = 6;

        private const int SW_MAXIMIZE = 3;

        static void Main(string[] args)
        {
            WriteLineColor("Приложение запущено.", "Yellow");

            if (args.Length > 0)
            {
                int i = 0;
                while ( i < args.Length )
                {
                    WriteLineColor("\n", "Magenta");
                    WriteLineColor("\n===================================================================", "Magenta");
                    WriteLineColor("Внимание приложение запущено с аргументом: ", "Yellow");
                    WriteLineColor(args[i], "Red");

                    if (args[i] == "beep")
                    {
                        beep();
                    }

                    i++;
                }

                file_processing(args[0],"argument_check");
                Console.ReadKey();
                Environment.Exit(0);
            }

            ShowWindow(ConsoleHandle, SW_MAXIMIZE);

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

                if (!File.Exists(Environment.CurrentDirectory + "//backup//" + EntryDate + "//" + "backup_" + name))
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

                                    clean_num(file.DirectoryName, file.Name);
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

        static Boolean clean_num(string file_patch, string file_name)
        {
            var wb = new XLWorkbook(file_patch + "\\" + file_name);

            var ws = wb.Worksheets.Worksheet(1);

            //считаем длину таблицы
            int total_cell = 16;

            while (ws.Cell(total_cell, 1).GetString() != "")
            {
                total_cell++;
            }

            //очистка всех числовых данных
            ws.Range("C19:D"+(total_cell-1)).Column(2).Clear();

            ws.Range("F19:G" + (total_cell - 1)).Column(2).Clear();

            //обозначем пунктиром все поля
            int start_cell = 19;

            while (start_cell < total_cell )
            {
                ws.Cell(start_cell, 4).Style.Border.BottomBorder = XLBorderStyleValues.Hair;
                ws.Cell(start_cell, 4).Style.Border.BottomBorderColor = XLColor.Black;

                ws.Cell(start_cell, 7).Style.Border.LeftBorder = XLBorderStyleValues.Hair;
                ws.Cell(start_cell, 7).Style.Border.RightBorder = XLBorderStyleValues.Hair;
                ws.Cell(start_cell, 7).Style.Border.BottomBorder = XLBorderStyleValues.Hair ;
                ws.Cell(start_cell, 7).Style.Border.BottomBorderColor = XLColor.Black;

                start_cell++;
            }

            //очищаем линию в конце документа
            int i =1;

            while (i < 8)
            {
                ws.Cell(total_cell-1, i).Style.Border.BottomBorder = XLBorderStyleValues.Hair;
                ws.Cell(total_cell-1, i).Style.Border.BottomBorderColor = XLColor.Black;

                i++;
            }

            //очищяем конец документа для добавления доп.строк
            i = 1;

            while (i <= 9)
            {
                ws.Range("A" + (total_cell + i).ToString() + ":I" + (total_cell + i).ToString()).Delete(XLShiftDeletedCells.ShiftCellsLeft);
                i++;
            }

            //добавление 10 дополнительных строк для неучтенных сразу строк
            i = 0;

            while (i <= 10)
            {
                int z = 1;

                while (z < 8)
                {
                    ws.Cell(total_cell + i, z).Style.Border.LeftBorder = XLBorderStyleValues.Hair;
                    ws.Cell(total_cell + i, z).Style.Border.RightBorder = XLBorderStyleValues.Hair;
                    ws.Cell(total_cell + i, z).Style.Border.BottomBorder = XLBorderStyleValues.Hair;
                    ws.Cell(total_cell + i, z).Style.Border.BottomBorderColor = XLColor.Black;

                    z++;
                }

                i++;
            }


            //Линия обозначающая конец акта
            i = 1;

            while (i < 8)
            {
                ws.Cell(total_cell + 11, i).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                ws.Cell(total_cell + 11, i).Style.Border.BottomBorderColor = XLColor.Black;

                i++;
            }

            int first_insert = total_cell + 12;

            ws.Cell(first_insert, 1).Value = "Всего наименований:";
            ws.Cell(first_insert , 1).Style.Font.FontName = "Arial";
            ws.Cell(first_insert , 1).Style.Font.FontSize = 11;
            ws.Range("A" + first_insert + ":G" + first_insert).Row(1).Merge();

            first_insert = total_cell + 14;

            ws.Cell(first_insert, 1).Value = "Всего единиц товара:";
            ws.Cell(first_insert, 1).Style.Font.FontName = "Arial";
            ws.Cell(first_insert, 1).Style.Font.FontSize = 11;
            ws.Range("A" + first_insert + ":G" + first_insert).Row(1).Merge();

            first_insert = total_cell + 16;

            ws.Cell(first_insert, 1).Value = "На сумму:";
            ws.Cell(first_insert, 1).Style.Font.FontName = "Arial";
            ws.Cell(first_insert, 1).Style.Font.FontSize = 11;
            ws.Range("A" + first_insert + ":G" + first_insert).Row(1).Merge();

            first_insert = total_cell + 18;

            ws.Cell(first_insert, 1).Value = "Материально-ответственное(ые) лицо(а) :";
            ws.Cell(first_insert, 1).Style.Font.FontName = "Arial";
            ws.Cell(first_insert, 1).Style.Font.FontSize = 11;
            ws.Range("A" + first_insert + ":G" + first_insert).Row(1).Merge();

            first_insert = total_cell + 20;

            ws.Cell(first_insert, 1).Value = "Начальник комиссии :";
            ws.Cell(first_insert, 1).Style.Font.FontName = "Arial";
            ws.Cell(first_insert, 1).Style.Font.FontSize = 11;
            ws.Range("A" + first_insert + ":G" + first_insert).Row(1).Merge();

            wb.Save();

            return true;
        }

        static Boolean file_processing(string file_patch, string file_name)
        {
            string par = "\\";

            if (file_name == "argument_check")
            {
                par = "";
                file_name = "";
            }

            var wb = new XLWorkbook(file_patch + par + file_name);

            var ws = wb.Worksheets.Worksheet(1);

            ws.PageSetup.Margins.Top = 0.208;
            ws.PageSetup.Margins.Bottom = 0.208;
            ws.PageSetup.Margins.Left = 0.416;
            ws.PageSetup.Margins.Right = 0.208;
            ws.PageSetup.Margins.Footer = 0.333;
            ws.PageSetup.Margins.Header = 0.333;

            wb.SaveAs(file_patch + par + file_name);

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
            ws.Cell(16, 1).Value = "Инвентаризационная опись № " + file_name.ToLower().Replace("c", "").Replace("с", "").Replace("_", "").Replace(".xlsx", "");

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

                    ws.Cell(total_cell, total_coll).Style.Font.FontSize = 11;

                    total_cell++;
                }

                total_coll++;
            }


            int cel1 = total_cell + 1;
            int cel3 = total_cell + 3;

            //Добавление в конец файла строк
            ws.Cell(cel1, 1).Value = "Материально-ответственное(ые) лицо(а) :";
            ws.Cell(cel1, 1).Style.Font.FontName = "Arial";
            ws.Cell(cel1, 1).Style.Font.FontSize = 11;
            ws.Range("A" + cel1 + ":G"+ cel1).Row(1).Merge();

            ws.Cell(cel3, 1).Value = "Начальник комиссии :";
            ws.Cell(cel3, 1).Style.Font.FontName = "Arial";
            ws.Cell(cel3, 1).Style.Font.FontSize = 11;
            ws.Range("A" + cel3 + ":G" + cel3).Row(1).Merge();


            //ширина колонок штрихкод,наименование,сумма.
            ws.Column(2).Width = 14;
            ws.Column(3).Width = 70;
            ws.Column(4).Width = 7;
            ws.Column(7).Width = 10;


//             ws.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
//             ws.Column(2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
//
//             ws.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
//             ws.Column(3).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
//
//             ws.Column(7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
//             ws.Column(7).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //считаем количество страниц на печать
            int pages = (total_cell+4) / 35;

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

        static void beep()
        {
            Console.Beep(659, 125);
            Console.Beep(659, 125);
            Thread.Sleep(125);
            Console.Beep(659, 125);
            Thread.Sleep(167);
            Console.Beep(523, 125);
            Console.Beep(659, 125);
            Thread.Sleep(125);
            Console.Beep(784, 125);
            Thread.Sleep(375);
            Console.Beep(392, 125);
            Thread.Sleep(375);
            Console.Beep(523, 125);
            Thread.Sleep(250);
            Console.Beep(392, 125);
            Thread.Sleep(250);
            Console.Beep(330, 125);
            Thread.Sleep(250);
            Console.Beep(440, 125);
            Thread.Sleep(125);
            Console.Beep(494, 125);
            Thread.Sleep(125);
            Console.Beep(466, 125);
            Thread.Sleep(42);
            Console.Beep(440, 125);
            Thread.Sleep(125);
            Console.Beep(392, 125);
            Thread.Sleep(125);
            Console.Beep(659, 125);
            Thread.Sleep(125);
            Console.Beep(784, 125);
            Thread.Sleep(125);
            Console.Beep(880, 125);
            Thread.Sleep(125);
            Console.Beep(698, 125);
            Console.Beep(784, 125);
            Thread.Sleep(125);
            Console.Beep(659, 125);
            Thread.Sleep(125);
            Console.Beep(523, 125);
            Thread.Sleep(125);
            Console.Beep(587, 125);
            Console.Beep(494, 125);
            Thread.Sleep(125);
            Console.Beep(523, 125);
            Thread.Sleep(250);
            Console.Beep(392, 125);
            Thread.Sleep(250);
            Console.Beep(330, 125);
            Thread.Sleep(250);
            Console.Beep(440, 125);
            Thread.Sleep(125);
            Console.Beep(494, 125);
            Thread.Sleep(125);
            Console.Beep(466, 125);
            Thread.Sleep(42);
            Console.Beep(440, 125);
            Thread.Sleep(125);
            Console.Beep(392, 125);
            Thread.Sleep(125);
            Console.Beep(659, 125);
            Thread.Sleep(125);
            Console.Beep(784, 125);
            Thread.Sleep(125);
            Console.Beep(880, 125);
            Thread.Sleep(125);
            Console.Beep(698, 125);
            Console.Beep(784, 125);
            Thread.Sleep(125);
            Console.Beep(659, 125);
            Thread.Sleep(125);
            Console.Beep(523, 125);
            Thread.Sleep(125);
            Console.Beep(587, 125);
            Console.Beep(494, 125);
            Thread.Sleep(375);
            Console.Beep(784, 125);
            Console.Beep(740, 125);
            Console.Beep(698, 125);
            Thread.Sleep(42);
            Console.Beep(622, 125);
            Thread.Sleep(125);
            Console.Beep(659, 125);
            Thread.Sleep(167);
            Console.Beep(415, 125);
            Console.Beep(440, 125);
            Console.Beep(523, 125);
            Thread.Sleep(125);
            Console.Beep(440, 125);
            Console.Beep(523, 125);
            Console.Beep(587, 125);
            Thread.Sleep(250);
            Console.Beep(784, 125);
            Console.Beep(740, 125);
            Console.Beep(698, 125);
            Thread.Sleep(42);
            Console.Beep(622, 125);
            Thread.Sleep(125);
            Console.Beep(659, 125);
            Thread.Sleep(167);
            Console.Beep(698, 125);
            Thread.Sleep(125);
            Console.Beep(698, 125);
            Console.Beep(698, 125);
            Thread.Sleep(625);
            Console.Beep(784, 125);
            Console.Beep(740, 125);
            Console.Beep(698, 125);
            Thread.Sleep(42);
            Console.Beep(622, 125);
            Thread.Sleep(125);
            Console.Beep(659, 125);
            Thread.Sleep(167);
            Console.Beep(415, 125);
            Console.Beep(440, 125);
            Console.Beep(523, 125);
            Thread.Sleep(125);
            Console.Beep(440, 125);
            Console.Beep(523, 125);
            Console.Beep(587, 125);
            Thread.Sleep(250);
            Console.Beep(622, 125);
            Thread.Sleep(250);
            Console.Beep(587, 125);
            Thread.Sleep(250);
            Console.Beep(523, 125);
            Thread.Sleep(1125);
            Console.Beep(784, 125);
            Console.Beep(740, 125);
            Console.Beep(698, 125);
            Thread.Sleep(42);
            Console.Beep(622, 125);
            Thread.Sleep(125);
            Console.Beep(659, 125);
            Thread.Sleep(167);
            Console.Beep(415, 125);
            Console.Beep(440, 125);
            Console.Beep(523, 125);
            Thread.Sleep(125);
            Console.Beep(440, 125);
            Console.Beep(523, 125);
            Console.Beep(587, 125);
            Thread.Sleep(250);
            Console.Beep(784, 125);
            Console.Beep(740, 125);
            Console.Beep(698, 125);
            Thread.Sleep(42);
            Console.Beep(622, 125);
            Thread.Sleep(125);
            Console.Beep(659, 125);
            Thread.Sleep(167);
            Console.Beep(698, 125);
            Thread.Sleep(125);
            Console.Beep(698, 125);
            Console.Beep(698, 125);
            Thread.Sleep(625);
            Console.Beep(784, 125);
            Console.Beep(740, 125);
            Console.Beep(698, 125);
            Thread.Sleep(42);
            Console.Beep(622, 125);
            Thread.Sleep(125);
            Console.Beep(659, 125);
            Thread.Sleep(167);
            Console.Beep(415, 125);
            Console.Beep(440, 125);
            Console.Beep(523, 125);
            Thread.Sleep(125);
            Console.Beep(440, 125);
            Console.Beep(523, 125);
            Console.Beep(587, 125);
            Thread.Sleep(250);
            Console.Beep(622, 125);
            Thread.Sleep(250);
            Console.Beep(587, 125);
            Thread.Sleep(250);
            Console.Beep(523, 125);
            Thread.Sleep(625);

            WriteLineColor("by part!zanes!", "Cyan");

            beep();
        }
    }
}
