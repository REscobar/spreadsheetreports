namespace SpreadSheetsReports.ReportModel
{
    using System.Collections.Generic;

    public interface IRowCollectionGenerator : IGenerator<IEnumerable<DocumentModel.Row>>
    {
    }
}