namespace SpreadSheetsReports.Sandbox
{
    using SpreadSheetsReports.ReportModel;
    using System.Collections.Generic;
    using System.Linq;

    public class ReverseRowCollectionSection :
        RowCollectionSection,
        IRowCollectionGenerator
    {
        public override IEnumerable<DocumentModel.Row> Generate()
        {
            this.Databind();
            return this.Rows.AsEnumerable().Reverse().Select(r => r?.Generate());
        }
    }
}
