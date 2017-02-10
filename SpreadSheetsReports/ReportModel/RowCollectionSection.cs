using System;
using System.Collections.Generic;
using SpreadSheetsReports.DocumentModel;

namespace SpreadSheetsReports.ReportModel
{
    public class RowCollectionSection : ReportControl, IRowCollectionGenerator
    {
        public RowCollection Rows { get; set; }

        public IEnumerable<DocumentModel.Row> Generate()
        {
            foreach (var row in this.Rows)
            {
                yield return row.Generate();
            }
        }

        protected override void DoRender()
        {
            foreach (var row in this.Rows)
            {
                if (row == null)
                {
                    continue;
                }

                row.Render();
            }
        }
    }
}