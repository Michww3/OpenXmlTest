using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace OpenXmlTest
{
    internal class Helpers
    {
        //получение позиции ячейки
        public static BorderStyle GetBorderStyle(int colIndex, int totalColumns, int rowIndex = 0, int totalRows = 1)
        {
            BorderStyle style = (colIndex == 0 ? BorderStyle.Left : BorderStyle.None) |
                                (colIndex == totalColumns - 1 ? BorderStyle.Right : BorderStyle.None) |
                                (rowIndex == 0 ? BorderStyle.Top : BorderStyle.None) |
                                (rowIndex == totalRows - 1 ? BorderStyle.Bottom : BorderStyle.None);

            return style;
        }

        //получение типа данных ячейки
        public CellType GetCellType(DataColumn dataColumn)
        {
            switch (dataColumn.DataType.Name)
            {
                case "Int32":
                case "Int64":
                    return CellType.Integer;

                case "Float":
                case "Double":
                case "Decimal":
                    return CellType.Float;

                case "DateTime":
                    return CellType.DateTime;

                case "Boolean":
                    return CellType.Boolean;

                default:
                    return CellType.String;
            }
        }

        //закрепелние заголовка и применение фильтра
        public void ApplyFilterAndFreezePane(Worksheet worksheet, uint columnCount)
        {
            //Создаем панель заморозки (Freeze Pane)
            SheetViews sheetViews = new SheetViews();
            SheetView sheetView = new SheetView() { WorkbookViewId = 0 };

            // Закрепляем первую строку (панель сверху)
            Pane pane = new Pane()
            {
                VerticalSplit = 1,
                TopLeftCell = "A2",
                ActivePane = PaneValues.BottomLeft,
                State = PaneStateValues.Frozen
            };
            sheetView.Append(pane);
            sheetViews.Append(sheetView);
            worksheet.InsertAt(sheetViews, 0);

            //Устанавливаем автофильтр на первую строку (например, A1:F1)
            char lastColumnLetter = (char)('A' + columnCount - 1);
            string filterRange = $"A1:{lastColumnLetter}1";

            AutoFilter autoFilter = new AutoFilter() { Reference = filterRange };
            worksheet.Append(autoFilter);
        }

        public uint GetStyle(Stylesheet stylesheet, uint cellTypeId, uint borderId)
        {
            List<CellFormat> cellFormats = stylesheet.CellFormats.Elements<CellFormat>().ToList();

            CellFormat cellFormat = cellFormats.FirstOrDefault(x => x.NumberFormatId == cellTypeId && x.BorderId == borderId);

            if (cellFormat != null)
            {
                return (uint)cellFormats.IndexOf(cellFormat);
            }

            // создаём новый стиль
            cellFormat = new CellFormat
            {
                NumberFormatId = cellTypeId,
                ApplyNumberFormat = true,
                BorderId = borderId,
                ApplyBorder = true
            };

            stylesheet.CellFormats.Append(cellFormat);

            return (uint)stylesheet.CellFormats.ChildElements.Count - 1;
        }

        //начальная инициализация стилей
        public Stylesheet InitStylesheet()
        {
            CellFormats cellFormats = new CellFormats();
            Borders borders = new Borders();
            borders.Append(new Border());
            cellFormats.Append(new CellFormat { NumberFormatId = (uint)CellType.String, ApplyNumberFormat = true, BorderId = 0, ApplyBorder = true });

            for (int i = 0; i < 16; i++)
            {
                BorderStyleValues borderStyleLeft = (i & 1) != 0 ? BorderStyleValues.Thick : BorderStyleValues.Thin;
                BorderStyleValues borderStyleRight = (i & 2) != 0 ? BorderStyleValues.Thick : BorderStyleValues.Thin;
                BorderStyleValues borderStyleTop = (i & 4) != 0 ? BorderStyleValues.Thick : BorderStyleValues.Thin;
                BorderStyleValues borderStyleBottom = (i & 8) != 0 ? BorderStyleValues.Thick : BorderStyleValues.Thin;

                borders.Append(
                    new Border(
                        new LeftBorder { Style = borderStyleLeft, Color = new Color { Auto = true } },
                        new RightBorder { Style = borderStyleRight, Color = new Color { Auto = true } },
                        new TopBorder { Style = borderStyleTop, Color = new Color { Auto = true } },
                        new BottomBorder { Style = borderStyleBottom, Color = new Color { Auto = true } }
                    )
                );

                uint borderId = (uint)borders.ChildElements.Count - 1;

                cellFormats.Append(new CellFormat {NumberFormatId = (uint)CellType.String, ApplyNumberFormat = true, BorderId = borderId, ApplyBorder = true });
            }

            return new Stylesheet(
                // Определяем шрифты
                new Fonts(
                    new Font()
                ),

                // Определяем заливки
                new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None })
                ),

                borders,

                cellFormats
            );
        }
    }
}
