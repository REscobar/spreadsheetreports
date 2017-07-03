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
        }

        public Sheet(IEnumerable<Row> rows)
        {
            this.Rows = new List<Row>(rows).AsReadOnly();
        }

        [DataMember]
        public string Name { get; set; }

        [DataMember]
        public IEnumerable<Row> Rows { get; set; }
    }
}