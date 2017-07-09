namespace SpreadSheetsReports.WpfUi.Sheets
{
    internal interface ISheetConverter
    {
        ReportModel.Sheet ConvertToSheet(SheetBinder binder);


    }
}