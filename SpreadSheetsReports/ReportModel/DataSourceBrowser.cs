namespace SpreadSheetsReports.ReportModel
{
    using System;

    public abstract class DataSourceBrowser : IDataSourceBrowser
    {
        public event Action OnCurrentChanging;

        public event Action OnCurrentChanged;

        public abstract object Current { get; }

        public abstract object GetValue(string property);

        public virtual bool MoveNext()
        {
            this.OnCurrentChanging?.Invoke();
            var returnValue = this.InnerMoveNext();
            this.OnCurrentChanged?.Invoke();

            return returnValue;
        }

        protected abstract bool InnerMoveNext();
    }
}
