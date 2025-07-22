using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlTest
{
    public enum HeaderPosition
    {
        Inner = 1,
        Left,
        Right,
        One
    }

    public enum CellPosition : uint
    {
        Top = 5,
        TopRight,
        Right,
        BottomRight,
        Bottom,
        BottomLeft,
        Left,
        TopLeft,
        Inner
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
