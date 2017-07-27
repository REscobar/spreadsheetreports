namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq.Expressions;
    using System.Runtime.Serialization;
    using System.Xml.Serialization;

    [DataContract]
    public class PropertyBinding : PropertyBindingBase
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
            this.Expression = dataMember;
        }

        [IgnoreDataMember]
        public DataSourceBrowser DataSource { get; set; }

        public virtual PropertyBindingCollection Owner { get; set; }

        protected internal override void PerformBind(ReportControl obj)
        {
            if (obj == null)
            {
                throw new ArgumentNullException(nameof(obj));
            }

            var source = this.DataSource;
            if (source == null)
            {
                var newSource = DataBindingContext.Peek();
                source = newSource as DataSourceBrowser;
                if (source == null)
                {
                    source = new ObjectDataSourceBrowser(newSource);
                }
            }

            object value = source.GetValue(this.Expression);
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

                ParameterExpression obj = System.Linq.Expressions.Expression.Parameter(objectType, "obj");
                ParameterExpression property = System.Linq.Expressions.Expression.Parameter(typeof(string), "property");

                ParameterExpression item = System.Linq.Expressions.Expression.Parameter(parameterType, "item");
                ParameterExpression value = System.Linq.Expressions.Expression.Parameter(objectType, "value");

                BinaryExpression cast = System.Linq.Expressions.Expression.Assign(item, System.Linq.Expressions.Expression.TypeAs(obj, parameterType));

                List<SwitchCase> cases = new List<SwitchCase>();
                foreach (var prop in parameterType.GetProperties())
                {
                    var targetProperty = System.Linq.Expressions.Expression.Property(item, prop.Name);
                    BinaryExpression caseExpression = System.Linq.Expressions.Expression.Assign(targetProperty, System.Linq.Expressions.Expression.Convert(value, targetProperty.Type));
                    SwitchCase caseStatement = System.Linq.Expressions.Expression.SwitchCase(caseExpression, System.Linq.Expressions.Expression.Constant(prop.Name, typeof(string)));
                    cases.Add(caseStatement);
                }

                SwitchExpression switchStatment = System.Linq.Expressions.Expression.Switch(voidType, property, null, null, cases.ToArray());

                BlockExpression block = System.Linq.Expressions.Expression.Block(voidType, new[] { item }, cast, switchStatment);

                var lambda = System.Linq.Expressions.Expression.Lambda<Action<object, string, object>>(block, obj, property, value);
                return lambda.Compile();
            });
        }
    }
}
