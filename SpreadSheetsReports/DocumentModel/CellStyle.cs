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

        /// <summary>
        /// Gets or sets the <see cref="DocumentModel.VerticalAlignment"/>
        /// </summary>
        public VerticalAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// Gets or sets the indent.
        /// </summary>
        public byte Indent { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the inner text in the cell is wraped.
        /// </summary>
        public bool WrapText { get; set; }

        /// <summary>
        /// Gets or sets the content rotation.
        /// </summary>
        public sbyte Rotation { get; set; }

        /// <summary>
        /// Gets or sets the top <see cref="BorderStyle"/>
        /// </summary>
        public BorderStyle BorderStyleTop { get; set; }

        /// <summary>
        /// Gets or sets the bottom <see cref="BorderStyle"/>
        /// </summary>
        public BorderStyle BorderStyleBottom { get; set; }

        /// <summary>
        /// Gets or sets the left <see cref="BorderStyle"/>
        /// </summary>
        public BorderStyle BorderStyleLeft { get; set; }

        /// <summary>
        /// Gets or sets the right <see cref="BorderStyle"/>
        /// </summary>
        public BorderStyle BorderStyleRight { get; set; }

        /// <summary>
        /// Gets or sets the diagonal <see cref="BorderStyle"/>
        /// </summary>
        public BorderStyle BorderStyleDiagonalUpLeftToBottomRight { get; set; }

        /// <summary>
        /// Gets or sets the diagonal <see cref="BorderStyle"/>
        /// </summary>
        public BorderStyle BorderStyleDiagonalUpRightToBottomLeft { get; set; }

        /// <summary>
        /// Gets or sets the <see cref="FillPatternStyle"/>
        /// </summary>
        public FillPatternStyle? FillPatternStyle { get; set; }

        /// <summary>
        /// Gets or sets the fill pattern <see cref="Color"/>
        /// </summary>
        public Color? FillPatternColor { get; set; }

        /// <summary>
        /// Gets or sets the background <see cref="Color"/>
        /// </summary>
        public Color? BackgroundColor { get; set; }
    }
}
