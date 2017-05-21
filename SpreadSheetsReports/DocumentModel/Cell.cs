namespace SpreadSheetsReports.DocumentModel
{
    using ReportModel;

    /// <summary>
    /// Represents a cell in a <see cref="Row"/>
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// Gets or sets the <see cref="CellStyle"/>
        /// </summary>
        public CellStyle? Style { get; set; }

        /// <summary>
        /// Gets or sets style class name
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// Gets or sets the cell value
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// Gets or sets the <see cref="CellType"/>
        /// </summary>
        public CellType Type { get; set; }
    }
}