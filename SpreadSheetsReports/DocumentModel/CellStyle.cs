namespace SpreadSheetsReports.DocumentModel
{
    public class CellStyle
    {
        public CellStyle()
        {
            //this.FontStyle = new FontStyle();
            //this.BorderStyleBottom = new BorderStyle();
            //this.BorderStyleDiagonalUpLeftToBottomRight = new BorderStyle();
            //this.BorderStyleDiagonalUpRightToBottomLeft = new BorderStyle();
            //this.BorderStyleLeft = new BorderStyle();
            //this.BorderStyleRight = new BorderStyle();
            //this.BorderStyleTop = new BorderStyle();
        }

        public bool ShrinkToFit { get; set; }

        public FontStyle FontStyle { get; set; }

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

        public BorderStyle BorderStyleDiagonalUpLeftToBottomRight { get; set; }

        public BorderStyle BorderStyleDiagonalUpRightToBottomLeft { get; set; }

        public FillPatternStyle FillPatternStyle { get; set; }

        public Color? FillPatternColor { get; set; }
    }
}
