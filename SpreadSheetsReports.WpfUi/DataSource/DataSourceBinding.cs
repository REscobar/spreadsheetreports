﻿// Copyright (c). All rights reserved.

namespace SpreadSheetsReports.WpfUi.DataSource
{
    using System;
    using System.ComponentModel;
    using DataBinders;
    using ReportModel;

    public class DataSourceBinding : INotifyPropertyChanged, IBinder<PropertyBindingBase>
    {
        private string propertyName;
        private string expression;
        private string type;

        public event PropertyChangedEventHandler PropertyChanged;

        public string PropertyName
        {
            get
            {
                return this.propertyName;
            }

            set
            {
                if (this.propertyName == value)
                {
                    return;
                }

                this.propertyName = value;
                this.NotifyPropertyChanged(nameof(this.PropertyName));
            }
        }

        public string Expression
        {
            get
            {
                return this.expression;
            }

            set
            {
                if (this.expression == value)
                {
                    return;
                }

                this.expression = value;
                this.NotifyPropertyChanged(nameof(this.Expression));
            }
        }

        public string Type
        {
            get
            {
                return this.type;
            }

            set
            {
                if (this.type == value)
                {
                    return;
                }
                this.type = value;
                this.NotifyPropertyChanged(nameof(this.Type));
            }
        }

        private void NotifyPropertyChanged(string propertyName)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public PropertyBindingBase ConvertTo()
        {
            if (this.Type == typeof(PropertyBinding).Name)
            {
                return new PropertyBinding
                {
                    Expression = this.Expression,
                    PropertyName = this.PropertyName
                };
            }
            else
            {
                return new ExpressionBinding
                {
                    Expression = this.Expression,
                    PropertyName = this.PropertyName
                };
            }
        }

        public void ConvertFrom(PropertyBindingBase obj)
        {
            throw new NotImplementedException();
        }
    }
}