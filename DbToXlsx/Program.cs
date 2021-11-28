using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace DbToXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            Range range = null;
            object misValue = System.Reflection.Missing.Value;
            if (File.Exists(args[0]))
            {
                _Application application = new Application();
                Workbook wb = application.Workbooks.Add(misValue);
                Worksheet ws = wb.Worksheets[1];
                Console.WriteLine("Соединение с Excel установлено.");
                using (var connection = new SQLiteConnection($"DataSource='{args[0]}';Version=3;"))
                {
                    connection.Open();
                    Console.WriteLine("Соединение с SQLite установлено.");
                    SQLiteCommand sqCom = connection.CreateCommand();
                    #region Counting rows
                    sqCom.CommandText = $"SELECT COUNT(*) FROM infoBase;";
                    SQLiteDataReader reader = sqCom.ExecuteReader();
                    reader.Read();
                    int rows = reader.GetInt32(0);
                    Console.WriteLine($"Rows: {rows}");
                    reader.Close();
                    #endregion
                    #region Counting cols
                    sqCom.CommandText = "pragma table_info(infoBase);";
                    reader = sqCom.ExecuteReader();
                    int cols = 0;
                    while (reader.Read())
                    {
                        cols++;
                    }
                    Console.WriteLine($"Columns: {cols}");
                    reader.Close();
                    #endregion
                    range = ws.Range["A1", ws.Cells[rows + 1, cols].Address];
                    object[,] writeRange = range.Value2;
                    #region Reading headers
                    sqCom.CommandText = "PRAGMA table_info('infoBase');";
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
                    #endregion
                    #region Reading data
                    sqCom.CommandText = "SELECT * FROM infoBase;";
                    reader = sqCom.ExecuteReader();
                    reader.Read();
                    if (reader.HasRows)
                    {
                        int currentRow = 2;
                        while (reader.Read())
                        {
                            for (int i = 1; i <= cols; i++)
                            {
                                writeRange[currentRow, i] = reader.GetValue(i-1);
                            }
                            currentRow++;
                        }
                        range.Value2 = writeRange;
                    }
                    reader.Close();
                    connection.Close();
                    #endregion
                    range.Columns.AutoFit();
                    range.RowHeight = 15;
                    wb.SaveAs($"{Environment.CurrentDirectory}\\{args[0].Split('.')[0]}.xlsx");
                }
                wb.Close();
                application.Quit();
            }
            else
            {
                Console.WriteLine($"Не могу найти файл '{args[0]}'. Укажите верный аргумент программы и перезапустите.");
            }
        }
    }
}
