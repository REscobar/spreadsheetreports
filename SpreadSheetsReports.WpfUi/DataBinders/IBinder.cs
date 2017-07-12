namespace SpreadSheetsReports.WpfUi.DataBinders
{
    public interface IBinder<T>
    {
        T ConvertTo();

        void ConvertFrom(T obj);
    }
}