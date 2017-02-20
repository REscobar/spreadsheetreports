using System;
using System.Collections.Generic;
using SpreadSheetsReports.DocumentModel;
using System.Linq;

namespace SpreadSheetsReports.ReportModel
{
    public class RowCollectionSection : ReportControl, IRowCollectionGenerator
    {
        public RowCollection Rows { get; set; }

        public IEnumerable<DocumentModel.Row> Generate()
        {
            this.Databind();
            return this.Rows.Select(r => r?.Generate()).ToList();
        }
    }
}