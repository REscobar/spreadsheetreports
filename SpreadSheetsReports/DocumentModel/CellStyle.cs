using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetsReports.DocumentModel
{
    public class CellStyle
    {
        
        public bool ShrinkToFit { get; set; }

        public FontStyle Font { get; set; }


        public bool IsHidden { get; set; }

        public bool IsLocked { get; set; }

        public HorizontalAlignment HorizontalAlignment { get; set; }

        public VerticalAlignment VerticalAlignment { get; set; }

        public byte Indent { get; set; }

        public bool WrapText { get; set; }

        public sbyte Rotation { get; set; }

        public BorderStyle BorderStyleTop { get; set; }
        public BorderStyle BorderStyleBottom { get; set; }
        public BorderStyle BorderStyleLeft { get; set; }
        public BorderStyle BorderStyleRight { get; set; }
        public BorderStyle BorderStyleDiagonalLeftToRight { get; set; }
        public BorderStyle BorderStyleDiagonalRightToLeft { get; set; }

        public FillPatternStyle FillPatternStyle { get; set; }

        public Color? FillPatternColor { get; set; }
    }
}
