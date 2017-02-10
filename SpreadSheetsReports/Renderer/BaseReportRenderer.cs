namespace SpreadSheetsReports.Renderer
{
    using System.IO;
    using DocumentModel;
    using ReportModel;

    public abstract class BaseReportRenderer : IReportRenderer
    {
        public virtual Stream Render(ReportDefinition report)
        {
            var document = report.Generate();
            return this.RenderToStream(document);
        }

        protected abstract Stream RenderToStream(Document document);
    }
}
