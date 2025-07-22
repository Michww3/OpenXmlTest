using System;

namespace OpenXmlTest
{
    [Flags]
    public enum BorderStyle
    {
        None = 0,
        Left = 1 << 0,
        Right = 1 << 1,
        Top = 1 << 2,
        Bottom = 1 << 3
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
