namespace SpreadSheetsReports.DocumentModel
{
    /// <summary>
    /// Represents a cell style
    /// </summary>
    public class CellStyle
    {
        /// <summary>
        /// Gets or sets a value indicating whether the content of the cell should shrink to fill the area of the cell
        /// </summary>
        public bool ShrinkToFit { get; set; }

        /// <summary>
        /// Gets or sets the <see cref="DocumentModel.FontStyle"/>
        /// </summary>
        public FontStyle FontStyle { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the cell is hidden
        /// </summary>
        public bool IsHidden { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the cell is locked
        /// </summary>
        public bool IsLocked { get; set; }

        /// <summary>
        /// Gets or sets the <see cref="DocumentModel.HorizontalAlignment"/>
        /// </summary>
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

        public FillPatternStyle? FillPatternStyle { get; set; }

        public Color? FillPatternColor { get; set; }

        public Color? BackgroundColor { get; set; }
    }
}
