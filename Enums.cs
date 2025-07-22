namespace OpenXmlTest
{
    public enum CellPosition : uint
    {
        OneRowInner = 1,
        OneRowLeft,
        OneRowRight,
        OneRowOne,
        Top = 5,
        TopRight,
        Right,
        BottomRight,
        Bottom,
        BottomLeft,
        Left,
        TopLeft,
        Inner,
        OneColsTop,
        OneColsBottom,
        OneColsInner
    }

    public enum CellType : uint
    {
        Integer = 1,
        Float = 2,
        DateTime = 22,
        String,
        Boolean
    }

}
