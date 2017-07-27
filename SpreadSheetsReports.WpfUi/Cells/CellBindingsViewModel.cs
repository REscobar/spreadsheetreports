namespace SpreadSheetsReports.WpfUi.Cells
{
    using System.Linq;
    using DataSource;

    public class CellBindingsViewModel
    {
        private CellBinder binder;

        public CellBindingsViewModel(CellBinder binder)
        {
            this.binder = binder;
            var binding = this.binder.Bindings.FirstOrDefault(b => b.PropertyName == "Value");
            if (binding != null)
            {
                this.ValueBindingExpression = binding.Expression;
                this.ValueBindingType = binding.Type;
            }

            binding = this.binder.Bindings.FirstOrDefault(b => b.PropertyName == "Type");

            if (binding != null)
            {
                this.TypeBndingExpression= binding.Expression;
                this.TypeBindingType = binding.Type;
            }
        }

        public string ValueBindingExpression { get; set; }

        public string ValueBindingType { get; set; }

        public string TypeBndingExpression { get; set; }

        public string TypeBindingType { get; set; }

        internal void Copy()
        {
            var binding = this.binder.Bindings.FirstOrDefault(b => b.PropertyName == "Value");
            if (binding != null)
            {
                this.binder.Bindings.Remove(binding);
            }

            if (!string.IsNullOrWhiteSpace(this.ValueBindingExpression))
            {
                var newBinding = new DataSourceBinding
                {
                    Expression = this.ValueBindingExpression,
                    PropertyName = "Value",
                    Type = this.ValueBindingType
                };

                this.binder.Bindings.Add(newBinding);
            }

            binding = this.binder.Bindings.FirstOrDefault(b => b.PropertyName == "Type");
            if (binding != null)
            {
                this.binder.Bindings.Remove(binding);
            }

            if (!string.IsNullOrWhiteSpace(this.TypeBndingExpression))
            {
                var newBinding = new DataSourceBinding
                {
                    Expression = this.TypeBndingExpression,
                    PropertyName = "Type",
                    Type = this.TypeBindingType
                };

                this.binder.Bindings.Add(newBinding);
            }
        }
    }
}
