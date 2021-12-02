using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace DbToXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            args = new string[1];
            args[0] = "BudeBase.db";
            bool alreadyClosed = false;
            Range range = null;
            object misValue = System.Reflection.Missing.Value;
            if (File.Exists(args[0]) && args[0].EndsWith(".db"))
            {
                _Application application = new Application();
                Workbook wb = application.Workbooks.Add(misValue);
                Console.WriteLine("Соединение с Excel установлено.");
                using (var connection = new SQLiteConnection($"DataSource='{args[0]}';Version=3;"))
                {
                    connection.Open();
                    Console.WriteLine("Соединение с SQLite установлено.");
                    SQLiteCommand sqCom = connection.CreateCommand();
                    #region Getting the names of the tables
                    sqCom.CommandText = "SELECT name FROM sqlite_master WHERE type='table';";
                    SQLiteDataReader reader = sqCom.ExecuteReader();
                    List<string> tables = new List<string>();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            tables.Add(reader.GetValue(0).ToString());
                        }
                    }
                    reader.Close();
                    Console.WriteLine($"Найдено таблиц: {tables.Count}");
                    #endregion
                    int sheetCounter = 1;
                    foreach (string name in tables)
                    {
                        Worksheet ws;
                        if (sheetCounter != 1)
                        {
                            ws = (Worksheet)wb.Worksheets.Add();
                        }
                        else
                        {
                            ws = wb.Worksheets[1];
                        }
                        #region Counting rows
                        sqCom.CommandText = $"SELECT COUNT(*) FROM [{name}];";
                        reader = sqCom.ExecuteReader();
                        reader.Read();
                        int rows = reader.GetInt32(0);
                        reader.Close();
                        Console.WriteLine($"Строки: {rows}");
                        #endregion
                        #region Counting cols
                        sqCom.CommandText = $"SELECT * FROM [{name}];";
                        reader = sqCom.ExecuteReader();
                        int cols = reader.VisibleFieldCount;
                        reader.Close();
                        Console.WriteLine($"Столбцы: {cols}");
                        #endregion
                        range = ws.Range["A1", ws.Cells[rows + 1, cols].Address];
                        object[,] writeRange = range.Value2;
                        #region Reading headers
                        sqCom.CommandText = $"PRAGMA table_info([{name}]);";
                        reader = sqCom.ExecuteReader();
                        if (reader.HasRows)
                        {
                            for (int i = 1; i <= cols; i++)
                            {
                                reader.Read();
                                writeRange[1, i] = reader.GetValue(1).ToString();
                            }
                        }
                        reader.Close();
                        Console.WriteLine("Имена полей загружены");
                        #endregion
                        #region Reading data
                        sqCom.CommandText = $"SELECT * FROM [{name}];";
                        reader = sqCom.ExecuteReader();
                        if (reader.HasRows)
                        {
                            int currentRow = 2;
                            while (reader.Read())
                            {
                                for (int i = 1; i <= cols; i++)
                                {
                                    writeRange[currentRow, i] = reader.GetValue(i - 1);
                                }
                                currentRow++;
                            }
                            range.Value2 = writeRange;
                        }
                        reader.Close();
                        Console.WriteLine("Данные загружены");
                        #endregion
                        range.Columns.AutoFit();
                        range.RowHeight = 15;
                        sheetCounter++;
                    }
                    connection.Close();
                    try
                    {
                        wb.SaveAs($"{Environment.CurrentDirectory}\\{args[0].Split('.')[0]}.xlsx");
                        Console.WriteLine($"{args[0].Split('.')[0]}.xlsx сохранён");
                    }
                    catch
                    {
                        Console.WriteLine("Сохранение отменено.");
                        wb.Close();
                        application.Quit();
                        alreadyClosed = true;
                    }
                }
                if (!alreadyClosed)
                {
                    wb.Close();
                    application.Quit();
                }
            }
            else if (File.Exists(args[0]) && args[0].EndsWith(".xlsx"))
            {
                Console.WriteLine("Не забывайте, что первая строка в файле Excel должна состоять из названий столбцов.");
                _Application application = new Application();
                Workbook wb = application.Workbooks.Open(args[0]);
                Worksheet ws = wb.Worksheets[1];
                Console.WriteLine("Соединение с Excel установлено.");
                object[,] readRange = ws.UsedRange.Value2;
                using (var connection = new SQLiteConnection($"DataSource='{args[0].Split('.')[0]}.db';Version=3;"))
                {
                    connection.Open();
                    SQLiteCommand sqCom = connection.CreateCommand();
                    sqCom.CommandText = $"Create Table {args[0].Split('.')[0]}()";
                    sqCom.ExecuteNonQuery();
                    connection.Close();
                }
            }
            else
            {
                Console.WriteLine($"Не могу найти файл '{args[0]}' или неверно указано его расширение. Перезапустите и укажите верный аргумент программы.");
            }
        }
    }
}
