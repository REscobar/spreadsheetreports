namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq.Expressions;
    using System.Xml.Serialization;

    [Serializable]
    public class PropertyBinding : IPropertyBinding
    {
        private static readonly ConcurrentDictionary<Type, Action<object, string, object>> WriterCache;
        private static readonly ConcurrentDictionary<Type, Func<object, string, object>> ReaderCache;

        static PropertyBinding()
        {
            WriterCache = new ConcurrentDictionary<Type, Action<object, string, object>>();
            ReaderCache = new ConcurrentDictionary<Type, Func<object, string, object>>();
        }

        public PropertyBinding()
        {
        }

        public PropertyBinding(string propertyName, object dataSource, string dataMember)
        {
            this.PropertyName = propertyName;
            var dataBrowser = dataSource as DataSourceBrowser;

            if (dataBrowser == null)
            {
                dataBrowser = new ObjectDataSourceBrowser(dataSource);
            }

            this.DataSource = dataBrowser;
            this.DataMember = dataMember;
        }

        public string PropertyName { get; set; }

        public string DataMember { get; set; }

        [XmlIgnore]
        public DataSourceBrowser DataSource { get; set; }

        internal PropertyBindingCollection Owner { get; set; }

        internal void PerformBind(ReportControl obj)
        {
            if (obj == null)
            {
                throw new ArgumentNullException(nameof(obj));
            }

            object value = this.DataSource.GetValue(this.DataMember);
            var objectType = obj.GetType();
            var setter = GetWriter(objectType);

            setter(obj, this.PropertyName, value);
        }

        private static Action<object, string, object> GetWriter(Type parameterType)
        {
            return WriterCache.GetOrAdd(parameterType, (k) =>
            {
                Type objectType = typeof(object);
                Type voidType = typeof(void);

                ParameterExpression obj = Expression.Parameter(objectType, "obj");
                ParameterExpression property = Expression.Parameter(typeof(string), "property");

                ParameterExpression item = Expression.Parameter(parameterType, "item");
                ParameterExpression value = Expression.Parameter(objectType, "value");

                BinaryExpression cast = Expression.Assign(item, Expression.TypeAs(obj, parameterType));

                List<SwitchCase> cases = new List<SwitchCase>();
                foreach (var prop in parameterType.GetProperties())
                {
                    var targetProperty = Expression.Property(item, prop.Name);
                    BinaryExpression caseExpression = Expression.Assign(targetProperty, Expression.Convert(value, targetProperty.Type));
                    SwitchCase caseStatement = Expression.SwitchCase(caseExpression, Expression.Constant(prop.Name, typeof(string)));
                    cases.Add(caseStatement);
                }

                SwitchExpression switchStatment = Expression.Switch(voidType, property, null, null, cases.ToArray());

                BlockExpression block = Expression.Block(voidType, new[] { item }, cast, switchStatment);

                var lambda = Expression.Lambda<Action<object, string, object>>(block, obj, property, value);
                return lambda.Compile();
            });
        }
    }
}
