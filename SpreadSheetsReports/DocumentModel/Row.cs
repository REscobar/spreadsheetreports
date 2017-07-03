namespace SpreadSheetsReports.DocumentModel
{
    using System.Collections.Generic;
    using System.Runtime.Serialization;

    [DataContract]
    public class Row
    {
        public Row()
        {
            this.Cells = new List<Cell>();
        }

        [DataMember]
        public IEnumerable<Cell> Cells { get; set; }

        [DataMember]
        public float? Height { get; set; }

        [DataMember]
        public CellStyle Style { get; set; }
    }
}