namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using SpreadSheetsReports.DocumentModel;

    public class RowCollectionSection : ReportControl, IRowCollectionGenerator
    {
        public RowCollection Rows { get; set; }

        public virtual IEnumerable<DocumentModel.Row> Generate()
        {
            this.Databind();
            return this.Rows.Select(r => r?.Generate()).ToList();
        }
    }
}