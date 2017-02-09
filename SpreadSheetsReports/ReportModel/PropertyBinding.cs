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
            this.DataSource = dataSource;
            this.DataMember = dataMember;
        }

        public string PropertyName { get; set; }

        public string DataMember { get; set; }

        [XmlIgnore]
        public object DataSource { get; set; }

        internal PropertyBindingCollection Owner { get; set; }

        internal void PerformBind(ReportControl obj)
        {
            if (obj == null)
            {
                throw new ArgumentNullException(nameof(obj));
            }

            object value = this.ResolveValue();
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

                //MethodCallExpression debugger = Expression.Call(null, typeof(System.Diagnostics.Debugger).GetMethod("Break"));

                BlockExpression block = Expression.Block(voidType,new[] { item }, cast, switchStatment);

                var lambda = Expression.Lambda<Action<object, string, object>>(block, obj, property, value);
                return lambda.Compile();
            });
        }

        private static Func<object, string, object> GetReader(Type t)
        {
            return ReaderCache.GetOrAdd(t, (k) =>
            {
                Type objectType = typeof(object);

                ParameterExpression obj = Expression.Parameter(objectType, "obj");
                ParameterExpression property = Expression.Parameter(typeof(string), "property");

                ParameterExpression item = Expression.Parameter(t, "item");
                ParameterExpression result = Expression.Parameter(objectType, "result");

                BinaryExpression cast = Expression.Assign(item, Expression.TypeAs(obj, t));
                BinaryExpression assign = Expression.Assign(result, Expression.TypeAs(Expression.Constant(null), objectType));
                List<SwitchCase> cases = new List<SwitchCase>();
                foreach (var prop in t.GetProperties())
                {
                    BinaryExpression caseExpression = Expression.Assign(result, Expression.TypeAs(Expression.Property(item, prop.Name), objectType));
                    SwitchCase caseStatement = Expression.SwitchCase(caseExpression, Expression.Constant(prop.Name, typeof(string)));
                    cases.Add(caseStatement);
                }

                SwitchExpression switchStatment = Expression.Switch(property, assign, cases.ToArray());

                BlockExpression block = Expression.Block(objectType, new[] { item, result }, cast, assign, switchStatment);

                var lambda = Expression.Lambda<Func<object, string, object>>(block, obj, property);
                return lambda.Compile();
            });
        }

        private object ResolveValue()
        {
            return GetReader(this.DataSource.GetType())(this.DataSource, this.DataMember);
        }
    }
}
