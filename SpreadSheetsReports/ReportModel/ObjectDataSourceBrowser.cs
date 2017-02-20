namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq;
    using System.Linq.Expressions;

    public class ObjectDataSourceBrowser : DataSourceBrowser
    {
        private static readonly ConcurrentDictionary<Type, Func<string, object, object>> ReaderCache = new ConcurrentDictionary<Type, Func<string, object, object>>();

        private readonly object datasource;
        private object currentValue;
        private Func<string, object, object> currentReader;

        private IEnumerator enumerator;

        public ObjectDataSourceBrowser(object datasource)
        {
            this.datasource = datasource;
        }

        public override object Current
        {
            get
            {
                return this.currentValue;
            }
        }

        public override bool MoveNext()
        {
            if (this.enumerator == null)
            {
                this.CreateEnumerator();
            }

            var res = this.enumerator.MoveNext();

            if (res)
            {
                this.currentValue = this.enumerator.Current;
            }
            else
            {
                this.currentValue = null;
            }

            return res;
        }

        public override object GetValue(string property)
        {
            if (this.currentReader == null)
            {
                this.CreateEnumerator();
                this.MoveNext();
            }

            return this.currentReader(property, this.currentValue);
        }

        private void CreateEnumerator()
        {
            if (this.datasource is IEnumerable)
            {
                var datasourceItemType = this.datasource.GetType().GetGenericArguments().First();
                this.currentReader = GetReader(datasourceItemType);
                this.enumerator = (this.datasource as IEnumerable).GetEnumerator();
            }
            else
            {
                this.currentReader = GetReader(this.datasource.GetType());
                this.enumerator = this.DummyIterator().GetEnumerator();
            }
        }

        private IEnumerable DummyIterator()
        {
            yield return this.datasource;
        }

        private static Func<string, object, object> GetReader(Type t)
        {
            return ReaderCache.GetOrAdd(t, k =>
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

                var lambda = Expression.Lambda<Func<string, object, object>>(block, property, obj);
                return lambda.Compile();
            });
        }

    }
}
