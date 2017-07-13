namespace SpreadSheetsReports.WpfUi
{
    using ReportModel;
    using Sheets;
    using System.Collections.Generic;
    using System.Windows;

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
    }
}
