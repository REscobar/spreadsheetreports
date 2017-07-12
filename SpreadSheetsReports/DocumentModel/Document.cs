namespace SpreadSheetsReports.DocumentModel
{
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Represents an spreadsheet document
    /// </summary>
    public class Document
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Document"/> class.
        /// </summary>
        public Document()
        {
            this.Sheets = Enumerable.Empty<Sheet>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Document"/> class with the specified sheets.
        /// </summary>
        /// <param name="sheets">The sheets to initialize the document with.</param>
        public Document(IEnumerable<Sheet> sheets)
        {
            this.Sheets = new List<Sheet>(sheets).AsReadOnly();
        }

        /// <summary>
        /// Gets or sets the <see cref="Document"/> sheets.
        /// </summary>
        public IEnumerable<Sheet> Sheets { get; set; }
    }
}
