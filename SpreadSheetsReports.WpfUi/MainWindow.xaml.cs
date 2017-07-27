namespace SpreadSheetsReports.WpfUi
{
    using System.Collections.Generic;
    using System.Windows;
    using ReportModel;
    using Sheets;
    using System.Linq;
    using Microsoft.Win32;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.InitializeComponent();

            var binder = this.Editor.DataContext as SheetCollectionBinder;
            //binder.ConvertFrom(BorderType().Sheets);
        }

        private static ReportDefinition BorderType()
        {
            return new ReportDefinition
            {
                Sheets = new List<ISheetGenerator>
                {
                    new Sheet
                    {
                        Bindings = new PropertyBindingCollection
                        {
                            new PropertyBinding
                            {
                                DataSource = new ObjectDataSourceBrowser(new { Name = "Test Binding" }),
                                Expression = "Name",
                                PropertyName = "Name"
                            }
                        },
                        Name = "Border",
                        Content = new ReportSection
                        {
                            Header = new RowCollectionSection
                            {
                                Rows = new RowCollection
                                {
                                    new Row
                                    {
                                        Cells = new List<ICellGenerator>
                                        {
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.DashDot
                                                    }
                                                },
                                                Value = "DashDot"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.DashDotDot
                                                    }
                                                },
                                                Value = "DashDotDot"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.Dashed
                                                    }
                                                },
                                                Value = "Dashed"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.Dotted
                                                    }
                                                },
                                                Value = "Dotted"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.Double
                                                    }
                                                },
                                                Value = "Double"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.Hair
                                                    }
                                                },
                                                Value = "Hair"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.Medium
                                                    }
                                                },
                                                Value = "Medium"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.MediumDashDot
                                                    }
                                                },
                                                Value = "MediumDashDot"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.MediumDashDotDot
                                                    }
                                                },
                                                Value = "MediumDashDotDot"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.MediumDashed
                                                    }
                                                },
                                                Value = "MediumDashed"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.SlantedDashDot
                                                    }
                                                },
                                                Value = "SlantedDashDot"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.Thick
                                                    }
                                                },
                                                Value = "Thick"
                                            },
                                            new Cell
                                            {
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Color = DocumentModel.Color.BlackColor,
                                                        Type = DocumentModel.BorderType.Thin
                                                    }
                                                },
                                                Value = "Thin"
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };
        }

        private void New_Click(object sender, RoutedEventArgs e)
        {
            var binder = this.Editor.DataContext as SheetCollectionBinder;
            binder.ConvertFrom(new[] { new Sheet { Name = "Sheet1" } });
        }

        private void Open_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".ssrxml"; // Default file extension
            openFileDialog.Filter = "SpreadSheet Report Definition (.ssrxml)|*.ssrxml|Xml Files|*.xml|All Files|*.*"; // Filter files by extension

            if (openFileDialog.ShowDialog() == true)
            {
                var report = Serializer.Deserialize(openFileDialog.FileName);
                var binder = this.Editor.DataContext as SheetCollectionBinder;
                binder.ConvertFrom(report.Sheets);
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = ".ssrxml"; // Default file extension
            saveDialog.Filter = "SpreadSheet Report Definition (.ssrxml)|*.ssrxml|Xml Files|*.xml|All Files|*.*"; // Filter files by extension

            // Show save file dialog box
            bool? result = saveDialog.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                string filename = saveDialog.FileName;
                var binder = this.Editor.DataContext as SheetCollectionBinder;
                var definition = new ReportDefinition { Sheets = binder.ConvertTo().ToList() };
                Serializer.Serialize(definition, filename);
            }
        }
    }
}
