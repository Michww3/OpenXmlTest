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
        public static CellPosition GetCellPosition(int colIndex, int totalCols, int rowIndex = 0, int totalRows = 0)
        {
            bool isTop = rowIndex == 0;
            bool isBottom = rowIndex == totalRows - 1;
            bool isLeft = colIndex == 0;
            bool isRight = colIndex == totalCols - 1;

            if (totalRows == 1 && totalCols == 1)
            {
                return CellPosition.OneRowOne;
            }
            if (totalRows == 1)
            {
                if (isLeft)
                {
                    return CellPosition.OneRowLeft;
                }
                if (isRight)
                {
                    return CellPosition.OneRowRight;
                }

                return CellPosition.OneRowInner;
            }
            if (totalCols == 1)
            {
                if (isTop)
                {
                    return CellPosition.OneColsTop;
                }
                if (isBottom)
                {
                    return CellPosition.OneColsBottom;
                }

                return CellPosition.OneColsInner;
            }

            if (isTop && isLeft)
            {
                return CellPosition.TopLeft;
            }
            if (isTop && isRight)
            {
                return CellPosition.TopRight;
            }
            if (isBottom && isLeft)
            {
                return CellPosition.BottomLeft;
            }
            if (isBottom && isRight)
            {
                return CellPosition.BottomRight;
            }
            if (isTop)
            {
                return CellPosition.Top;
            }
            if (isBottom)
            {
                return CellPosition.Bottom;
            }
            if (isLeft)
            {
                return CellPosition.Left;
            }
            if (isRight)
            {
                return CellPosition.Right;
            }

            return CellPosition.Inner;

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
            Border[] borders = new Border[]
            {
                // 0 - default
                new Border(),
                // 1 - толстые границы сверху и снизу
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                ),
                // 2 - толстые границы сверху снизу и слева
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                ),
                // 3 - толстые границы сверху снизу и справа
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                ),
                // 4 - толстые границы  со всех сторон
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                ),
                // 5 - верхняя ячейка
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                ),
                // 6 - правая верхняя ячейка
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                ),
                // 7 - правая средняя ячейка
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                ),
                // 8 - правая нижняя ячейка
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                ),
                // 9 - нижняя ячейка
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                ),
                // 10 - левая нижняя ячейка
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                ),
                // 11 - левая средняя ячейка
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                ),
                // 12 - левая верхняя ячейка
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                ),
                // 13 - тонкие границы со всех сторон
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                ),
                // 14 - толстые границы сверху и по бокам
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                ),
                // 15 - толстые границы снизу и по бокам
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                ),
                // 16 - толстые по бокам
                new Border(
                    new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                    new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                    new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                )
            };

            CellFormat[] cellFormats = new CellFormat[borders.Length];

            for (int i = 0; i < borders.Length; i++)
            {
                cellFormats[i] = new CellFormat { BorderId = (uint)i, ApplyBorder = true };
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

                // Определяем границы ячеек
                new Borders(borders),

                new CellFormats(cellFormats)
            );
        }
    }
}
