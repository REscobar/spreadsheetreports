using System.Collections;
using System.Collections.Generic;

namespace SpreadSheetsReports.ReportModel
{
    internal interface IRowCollectionGenerator : IGenerator<IEnumerable<DocumentModel.Row>>
    {
    }
}