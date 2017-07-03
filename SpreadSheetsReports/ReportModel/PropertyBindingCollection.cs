namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Runtime.Serialization;

    [CollectionDataContract]
    public class PropertyBindingCollection : ICollection<PropertyBindingBase>
    {
        private readonly ReportControl owner;
        private readonly List<PropertyBindingBase> bindings;

        public PropertyBindingCollection()
        {
            this.bindings = new List<PropertyBindingBase>();
        }

        public PropertyBindingCollection(ReportControl owner)
            : this()
        {
            this.owner = owner;
        }

        public int Count
        {
            get
            {
                return this.bindings.Count;
            }
        }

        public bool IsReadOnly
        {
            get
            {
                return false;
            }
        }

        public void Add(PropertyBindingBase item)
        {
            item.Owner = this;
            this.bindings.Add(item);
        }

        public void Add(string propertyName, object dataSource, string dataMember)
        {
            if (string.IsNullOrWhiteSpace(propertyName))
            {
                throw new ArgumentException(nameof(propertyName));
            }

            if (string.IsNullOrEmpty(dataMember))
            {
                throw new ArgumentException(nameof(dataMember));
            }

            this.Add(new PropertyBinding(propertyName, dataSource, dataMember));
        }

        public void Clear()
        {
            this.bindings.Clear();
        }

        public bool Contains(PropertyBindingBase item)
        {
            return this.bindings.Contains(item);
        }

        public void CopyTo(PropertyBindingBase[] array, int arrayIndex)
        {
            this.bindings.CopyTo(array, arrayIndex);
        }

        public IEnumerator<PropertyBindingBase> GetEnumerator()
        {
            return this.bindings.GetEnumerator();
        }

        public bool Remove(PropertyBindingBase item)
        {
            return this.bindings.Remove(item);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
