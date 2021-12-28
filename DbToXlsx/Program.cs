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
            bool exceptionClosed = false, sheetAddFlag = true;
            object missing = System.Reflection.Missing.Value;
            List<string> tableNames = new List<string>();
            int rows = 0, cols = 0;
            Worksheet sheet;
            Range range = null;

            if (File.Exists(args[0]) && args[0].EndsWith(".db"))
            {
                _Application application = new Application();
                Workbook wb = application.Workbooks.Add(missing);
                Console.WriteLine("Connected to EXCEL");
                using (var connection = new SQLiteConnection($"DataSource='{args[0]}';Version=3;Read Only = True;"))
                {
                    connection.Open();
                    Console.WriteLine($"Connected to SQLite");
                    using (SQLiteCommand sqCom = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table';",connection))
                    {
                        using (SQLiteDataReader reader = sqCom.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    tableNames.Add(reader.GetValue(0).ToString());
                                }
                            }
                        }
                    }

                    Console.WriteLine($"Tables count: {tableNames.Count}");

                    foreach (string name in tableNames)
                    {
                        Console.WriteLine($"Processing table: {name}");
                        
                        if (sheetAddFlag)
                        {
                            sheet = (Worksheet)wb.Worksheets.Add();
                            sheetAddFlag = false;
                        }
                        else
                        {
                            sheet = wb.Worksheets[1];
                        }
                        #region Counting rows and cols
                        using (SQLiteCommand sqCom = new SQLiteCommand($"SELECT COUNT(*) FROM [{name}];",connection))
                        {
                            using (SQLiteDataReader reader = sqCom.ExecuteReader())
                            {
                                reader.Read();
                                rows = reader.GetInt32(0);
                            }
                        }

                        using (SQLiteCommand sqCom = new SQLiteCommand($"SELECT * FROM [{name}];", connection))
                        {
                            using (SQLiteDataReader reader = sqCom.ExecuteReader())
                            {
                                cols = reader.VisibleFieldCount;
                            }
                        }

                        Console.WriteLine($"Rows count: {rows}");
                        Console.WriteLine($"Columns count: {cols}");
                        #endregion
                        range = sheet.Range["A1", sheet.Cells[rows + 1, cols].Address];
                        object[,] writeRange = range.Value2;
                        #region Reading headers
                        using (SQLiteCommand sqCom = new SQLiteCommand($"PRAGMA table_info([{name}]);",connection))
                        {
                            using (SQLiteDataReader reader = sqCom.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    for (int i = 1; i <= cols; i++)
                                    {
                                        reader.Read();
                                        writeRange[1, i] = reader.GetValue(1).ToString();
                                    }
                                }
                            }
                        }
                        #endregion
                        #region Reading data
                        using (SQLiteCommand sqCom = new SQLiteCommand($"SELECT * FROM [{name}];", connection))
                        {
                            using (SQLiteDataReader reader = sqCom.ExecuteReader())
                            {
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
                            }
                        }
                        #endregion
                        range.Columns.AutoFit();
                        range.RowHeight = 15;
                    }
                    try
                    {
                        wb.SaveAs($"{Environment.CurrentDirectory}\\{args[0].Split('.')[0]}.xlsx");
                        Console.WriteLine($"{args[0].Split('.')[0]}.xlsx successfully saved!");
                    }
                    catch
                    {
                        Console.WriteLine("Saving cancelled.");
                        wb.Close();
                        application.Quit();
                        exceptionClosed = true;
                    }
                }
                if (!exceptionClosed)
                {
                    wb.Close();
                    application.Quit();
                }
            }
            else
            {
                Console.WriteLine($"Can't find the file '{args[0]}' or its extension is incorrect. Restart and provide the correct program argument.");
            }
        }
    }
}
