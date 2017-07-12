namespace SpreadSheetsReports.DocumentModel
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.Serialization;

    [DataContract]
    public class Sheet
    {
        public Sheet()
        {
            this.Rows = Enumerable.Empty<Row>();
            this.AreSummaryRowsBelowDetail = false;
            this.AreSummaryColumnsRightOfDetail = false;
        }

        public Sheet(IEnumerable<Row> rows)
            : this()
        {
            this.Rows = new List<Row>(rows).AsReadOnly();
        }

        /// <summary>
        /// Gets or sets a value indicating whether the summary rows are below grouped rows.
        /// </summary>
        public bool AreSummaryRowsBelowDetail { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the summary columns are to the right of the grouped columns.
        /// </summary>
        public bool AreSummaryColumnsRightOfDetail { get; set; }

        [DataMember]
        public string Name { get; set; }

        [DataMember]
        public IEnumerable<Row> Rows { get; set; }
    }
}