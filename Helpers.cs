using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Data;

namespace OpenXmlTest
{
    internal class Helpers
    {
        public struct StyleKey
        {
            uint NumberFormatId;
            uint BorderId;

            public StyleKey(uint NumberFormatId, uint BorderId)
            {
                this.NumberFormatId = NumberFormatId;
                this.BorderId = BorderId;
            }
        }

        //получение позиции ячейки
        public static BorderStyle GetBorderStyle(int colIndex, int totalColumns, int rowIndex = 0, int totalRows = 1)
        {
            BorderStyle style = ~((colIndex == 0 ? BorderStyle.Left : BorderStyle.None) |
                                (colIndex == totalColumns - 1 ? BorderStyle.Right : BorderStyle.None) |
                                (rowIndex == 0 ? BorderStyle.Top : BorderStyle.None) |
                                (rowIndex == totalRows - 1 ? BorderStyle.Bottom : BorderStyle.None)) &
                                (BorderStyle.Left | BorderStyle.Right | BorderStyle.Top | BorderStyle.Bottom);

            return style;
        }

        //получение типа данных ячейки
        public static CellType GetCellType(DataColumn dataColumn)
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
        public static void ApplyFilterAndFreezePane(Worksheet worksheet, uint columnCount)
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

        public static uint GetOrCreateStyle(Stylesheet stylesheet, Dictionary<StyleKey, uint> styleCache, uint excelNumberFormatId, uint borderId, ref uint styleIndexCounter)
        {
            var key = new StyleKey(excelNumberFormatId, borderId);

            if (styleCache.TryGetValue(key, out uint existingIndex))
                return existingIndex;

            // создаём новый стиль
            var cellFormat = new CellFormat
            {
                NumberFormatId = excelNumberFormatId,
                ApplyNumberFormat = true,
                BorderId = borderId,
                ApplyBorder = true
            };

            stylesheet.CellFormats.Append(cellFormat);
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.ChildElements.Count;

            uint newIndex = styleIndexCounter++;
            styleCache[key] = newIndex;

            return newIndex;
        }

        //начальная инициализация стилей
        public static Stylesheet InitStylesheet()
        {
            CellFormats cellFormats = new CellFormats();
            Borders borders = new Borders();

            for (int i = 0; i < 16; i++)
            {
                BorderStyleValues borderStyleLeft = (i & 1) != 0 ? BorderStyleValues.Thin : BorderStyleValues.Thick;
                BorderStyleValues borderStyleRight = (i & 2) != 0 ? BorderStyleValues.Thin : BorderStyleValues.Thick;
                BorderStyleValues borderStyleTop = (i & 4) != 0 ? BorderStyleValues.Thin : BorderStyleValues.Thick;
                BorderStyleValues borderStyleBottom = (i & 8) != 0 ? BorderStyleValues.Thin : BorderStyleValues.Thick;

                borders.Append(
                    new Border(
                        new LeftBorder { Style = borderStyleLeft, Color = new Color { Auto = true } },
                        new RightBorder { Style = borderStyleRight, Color = new Color { Auto = true } },
                        new TopBorder { Style = borderStyleTop, Color = new Color { Auto = true } },
                        new BottomBorder { Style = borderStyleBottom, Color = new Color { Auto = true } }
                    )
                );

                cellFormats.Append(new CellFormat { BorderId = (uint)i, ApplyBorder = true });
            }

            return new Stylesheet(
                // Определяем шрифты
                new Fonts(
                    new Font() // 0 - обычный шрифт
                ),

                // Определяем заливки (фоны ячеек)
                new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None })     // 0 - без заливки
                ),

                borders,

                cellFormats
            );
        }
    }
}
