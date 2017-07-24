namespace SpreadSheetsReports.WpfUi.DataSource
{
    using System.Collections.ObjectModel;

    public interface IDataSourceBindable
    {
        string DataMember { get; set; }

        ObservableCollection<DataSourceBinding> Bindings { get; set; }
    }
}