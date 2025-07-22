using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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


        //получение поизиции ячейки заголовка
        public static uint GetHeaderPosition(int columnsCount, int currentColumns)
        {
            if (columnsCount == 1)
            {
                return (uint)HeaderPosition.One;
            }
            if (currentColumns == 0)
            {
                return (uint)HeaderPosition.Left;
            }
            if (currentColumns < columnsCount - 1)
            {
                return (uint)HeaderPosition.Inner;
            }

            return (uint)HeaderPosition.Right;
        }
        //получение позиции ячейки
        public static CellPosition GetCellPosition(int rowIndex, int totalRows, int colIndex, int totalCols)
        {
            if (totalRows == 1)
            {
                return 0;
            }

            bool isTop = rowIndex == 0;
            bool isBottom = rowIndex == totalRows - 1;
            bool isLeft = colIndex == 0;
            bool isRight = colIndex == totalCols - 1;

            if (isTop && isLeft) return CellPosition.TopLeft;
            if (isTop && isRight) return CellPosition.TopRight;
            if (isBottom && isLeft) return CellPosition.BottomLeft;
            if (isBottom && isRight) return CellPosition.BottomRight;
            if (isTop) return CellPosition.Top;
            if (isBottom) return CellPosition.Bottom;
            if (isLeft) return CellPosition.Left;
            if (isRight) return CellPosition.Right;

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
                new Borders(

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
                    )
                ),


                new CellFormats(
                    // 0 - по умолчанию 
                    new CellFormat(),
                    //header
                    new CellFormat { BorderId = 1, ApplyBorder = true },      // 1 - заголовок (сверху и снизу толстая)
                    new CellFormat { BorderId = 2, ApplyBorder = true },      // 2 - заголовок: левая ячейка
                    new CellFormat { BorderId = 3, ApplyBorder = true },      // 3 - заголовок: правая ячейка
                    new CellFormat { BorderId = 4, ApplyBorder = true },      // 4 - заголовок: одна ячейка
                                                                              //default
                    new CellFormat { BorderId = 5, ApplyBorder = true },      // 5 - ячейка внутри таблицы (тонкие границы)
                    new CellFormat { BorderId = 6, ApplyBorder = true },      // 6 - левая верхняя ячейка
                    new CellFormat { BorderId = 7, ApplyBorder = true },      // 7 - левая средняя ячейка
                    new CellFormat { BorderId = 8, ApplyBorder = true },      // 8 - левая нижняя ячейка
                    new CellFormat { BorderId = 9, ApplyBorder = true },      // 9 - нижняя внутренняя 
                    new CellFormat { BorderId = 10, ApplyBorder = true },      // 10 - правая верхняя ячейка
                    new CellFormat { BorderId = 11, ApplyBorder = true },     // 11 - правая средняя ячейка
                    new CellFormat { BorderId = 12, ApplyBorder = true },     // 12 - правая нижняя ячейка
                    new CellFormat { BorderId = 13, ApplyBorder = true }     // 13 - верхняя внутренняя ячейка
                )
            );

        }
    }
}
