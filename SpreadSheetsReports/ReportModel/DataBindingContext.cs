namespace SpreadSheetsReports.ReportModel
{
    using System.Collections.Generic;

    public static class DataBindingContext
    {
        static readonly Stack<object> contextStack = new Stack<object>();

        public static void Push(object dataSource)
        {
            contextStack.Push(dataSource);
        }

        public static object Peek()
        {
            return contextStack.Peek();
        }

        public static object Pop()
        {
            return contextStack.Pop();
        }
    }
}
