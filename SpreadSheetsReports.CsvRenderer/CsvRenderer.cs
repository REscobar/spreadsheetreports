namespace SpreadSheetsReports.CsvRenderer
{
    using SpreadSheetsReports.Renderer;
    using System.Linq;
    using SpreadSheetsReports.DocumentModel;
    using System.IO;
    using SpreadSheetsReports.ReportModel;
    public class CsvRenderer : IReportRenderer
    {
        public Stream Render(ReportDefinition report)
        {
            var r = report.Generate();
            return this.RenderToStream(r); 
        }

        protected Stream RenderToStream(Document document)
        {
            MemoryStream sb = new MemoryStream();
            StreamWriter sw = new StreamWriter(sb);
            var sheetToExport = document.Sheets.First();
            foreach (var row in sheetToExport.Rows)
            {
                //Concatena los valores de las celdas con el separador ",", tal como lo requiere CSV
                var csvRow = string.Join(",", row.Cells.Select(s => s.Value.ToString()).ToArray());
                sw.WriteLine(csvRow);
            }
            sb.Seek(0, SeekOrigin.Begin);

            return sb;
        }
    }
}
