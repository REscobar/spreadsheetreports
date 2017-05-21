namespace SpreadSheetsReports.ReportModel
{
    using System.Collections;
    using System.Collections.Generic;

    internal interface IRowCollectionGenerator : IGenerator<IEnumerable<DocumentModel.Row>>
    {
    }
}