﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;

namespace OpenXmlTest
{
    class Programm
    {
        static void Main(string[] args)
        {
            DataTable myTable = new DataTable();

            myTable.Columns.Add("Имя", typeof(string));
            myTable.Columns.Add("Возраст", typeof(int));
            myTable.Columns.Add("Город", typeof(string));
            myTable.Columns.Add("Дата рождения", typeof(DateTime));
            myTable.Columns.Add("Студент", typeof(bool));
            myTable.Columns.Add("Баллы", typeof(double));
            myTable.Columns.Add("Опыт (лет)", typeof(int));
            myTable.Columns.Add("Проекты", typeof(int));

            myTable.Rows.Add("Алексей", 30, "Минск", new DateTime(1994, 5, 10), true, 87.5, 8, 15);
            myTable.Rows.Add("Мария", 25, "Гомель", new DateTime(1999, 11, 23), false, 92.3, 3, 7);
            myTable.Rows.Add("Иван", 28, "Брест", new DateTime(1996, 2, 3), true, 78.0, 6, 12);
            myTable.Rows.Add("Ольга", 22, "Витебск", new DateTime(2002, 8, 17), false, 95.8, 2, 5);
            myTable.Rows.Add("Никита", 35, "Гродно", new DateTime(1989, 1, 30), true, 65.4, 12, 20);
            myTable.Rows.Add("Елена", 29, "Могилёв", new DateTime(1995, 3, 12), false, 88.1, 7, 11);
            myTable.Rows.Add("Дмитрий", 40, "Гомель", new DateTime(1983, 7, 5), true, 72.9, 15, 22);
            myTable.Rows.Add("Светлана", 27, "Минск", new DateTime(1996, 12, 1), false, 94.7, 5, 9);
            myTable.Rows.Add("Владимир", 31, "Брест", new DateTime(1992, 9, 14), true, 81.3, 10, 17);
            myTable.Rows.Add("Наталья", 24, "Витебск", new DateTime(2000, 6, 25), false, 90.5, 2, 4);
            myTable.Rows.Add("Андрей", 33, "Гродно", new DateTime(1990, 4, 18), true, 68.7, 11, 19);
            myTable.Rows.Add("Ирина", 26, "Минск", new DateTime(1997, 11, 30), true, 77.4, 4, 6);
            myTable.Rows.Add("Олег", 38, "Могилёв", new DateTime(1985, 1, 20), false, 83.6, 14, 18);

            //myTable.Columns.Add("Name", typeof(string));
            //myTable.Columns.Add("Возраст", typeof(int));

            //myTable.Rows.Add("Alice", 12);
            //myTable.Rows.Add("Alice", 12);
            //myTable.Rows.Add("Alice", 12);
            //myTable.Rows.Add("Bob");
            //myTable.Rows.Add("Charlie");
            //myTable.Rows.Add("Diana");
            ExcelExporter excelExporter = new ExcelExporter();

            excelExporter.ExportDataTableToExcel(myTable, @"C:\\Users\\zimnitskyaa\\Desktop\\test1.xlsx");

            Process.Start(new ProcessStartInfo()
            {
                FileName = @"C:\\Users\\zimnitskyaa\\Desktop\\test1.xlsx",
                UseShellExecute = true
            });
        }
    }

    public class ExcelExporter
    {
        public void ExportDataTableToExcel(DataTable table, string filePath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                Helpers helpers = new Helpers();
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Отчёт"
                };
                sheets.Append(sheet);

                // Создание стилей
                WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet stylesheet = helpers.InitStylesheet();
                stylesPart.Stylesheet = stylesheet;

                // Заголовок
                Row headerRow = new Row();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    Cell cell = new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(table.Columns[i].ColumnName),
                        StyleIndex = (uint)Helpers.GetBorderStyle(i, table.Columns.Count) + 1
                    };

                    headerRow.Append(cell);
                }
                sheetData.Append(headerRow);

                helpers.ApplyFilterAndFreezePane(worksheetPart.Worksheet, (uint)table.Columns.Count);

                // Данные
                for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
                {
                    Row dataRow = new Row();
                    for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                    {
                        Cell cell = new Cell();
                        object value = table.Rows[rowIndex][columnIndex];
                        DataColumn currentColumn = table.Columns[columnIndex];

                        CellType cellType = helpers.GetCellType(currentColumn);

                        BorderStyle borderStyle = Helpers.GetBorderStyle(columnIndex, table.Columns.Count, rowIndex, table.Rows.Count);
                        uint borderStyleIndex = (uint)borderStyle + 1;

                        switch (cellType)
                        {
                            case CellType.Integer:

                                cell.CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture));
                                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                cell.StyleIndex = helpers.GetStyle(stylesheet, (uint)CellType.Integer, borderStyleIndex); // для целых чисел
                                break;

                            case CellType.Float:

                                cell.CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture));
                                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                cell.StyleIndex = helpers.GetStyle(stylesheet, (uint)CellType.Float, borderStyleIndex); // для целых чисел
                                break;

                            case CellType.DateTime:

                                DateTime dt = (DateTime)value;
                                cell.CellValue = new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture));
                                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                cell.StyleIndex = helpers.GetStyle(stylesheet, (uint)CellType.DateTime, borderStyleIndex);
                                break;

                            case CellType.Boolean:

                                cell.CellValue = new CellValue((bool)value ? "Да" : "Нет");
                                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                cell.StyleIndex = borderStyleIndex;
                                break;

                            default:

                                cell.CellValue = new CellValue(value?.ToString());
                                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                cell.StyleIndex = borderStyleIndex;
                                break;
                        }

                        dataRow.Append(cell);
                    }
                    sheetData.Append(dataRow);
                }

                stylesheet.Save();
                workbookPart.Workbook.Save();
            }
        }
    }
}