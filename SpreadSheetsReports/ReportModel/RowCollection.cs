namespace SpreadSheetsReports.ReportModel
{
    using System.Collections.Generic;

    public class RowCollection : List<IRowGenerator>, IList<IRowGenerator>
    {
    }
}
