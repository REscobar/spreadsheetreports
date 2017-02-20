using SpreadSheetsReports;
using SpreadSheetsReports.NPOIRenderer;
using SpreadSheetsReports.Renderer;
using SpreadSheetsReports.ReportModel;
using SpreadSheetsReports.SpreadsheetLightRenderer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetsReports.Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            //System.Xml.Serialization.XmlSerializer reader = new System.Xml.Serialization.XmlSerializer(typeof(ReportDefinition));
            //var definition = reader.Deserialize(File.OpenRead("TestData.xml")) as ReportDefinition;

            var definition = TestReport();

            IReportRenderer renderer = new NPOIRenderer.NPOIRenderer();
            var stream = File.Create("Workbook1.xlsx");
            renderer.Render(definition).CopyTo(stream);
            stream.Flush();
            stream.Close();

            definition = TestReport();
            renderer = new SpreadsheetLightRenderer.SpreadsheetLightRenderer();
            stream = File.Create("Workbook2.xlsx");
            renderer.Render(definition).CopyTo(stream);
            stream.Flush();
            stream.Close();


            definition = TestReport2();
            renderer = new NPOIRenderer.NPOIRenderer();
            stream = File.Create("BindingNpoi.xlsx");
            renderer.Render(definition).CopyTo(stream);
            stream.Flush();
            stream.Close();

            definition = TestReport2();
            renderer = new SpreadsheetLightRenderer.SpreadsheetLightRenderer();
            stream = File.Create("BindingSL.xlsx");
            renderer.Render(definition).CopyTo(stream);
            stream.Flush();
            stream.Close();

        }

        private static ReportDefinition TestReport()
        {
            var testing = new ReportDefinition
            {
                Sheets = new List<Sheet>
                {
                    new Sheet
                    {
                        Name= "ONE",
                        Content = new ReportSection
                        {
                            Header = new RowCollectionSection
                            {
                                Rows = new RowCollection
                                {
                                    new Row
                                    {
                                        Height = 125,
                                        Cells = new List<Cell>
                                        {
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced",
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Type = DocumentModel.BorderType.Thick,
                                                        Color = new DocumentModel.Color
                                                        {
                                                            Red = 255
                                                        }
                                                    },
                                                    FontStyle = new DocumentModel.FontStyle
                                                    {
                                                        Color = new DocumentModel.Color
                                                        {
                                                            Alpha = 255,
                                                            Red = 127,
                                                            Blue = 127,
                                                            Green = 255
                                                        },
                                                        FontName = "Algerian"
                                                    }
                                                }
                                            }
                                        }
                                    },
                                    new Row
                                    {
                                        Cells = new List<Cell>
                                        {
                                            null,
                                            new Cell
                                            {
                                                Value = "NOW()",
                                                Type = CellType.Formula,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    Indent=2,
                                                    Rotation= 35,
                                                    IsLocked= true
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "Testing",
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    IsHidden= true,
                                                    FillPatternStyle= DocumentModel.FillPatternStyle.VerticalStripe,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Red = 64,
                                                        Blue = 255,
                                                        Green = 0
                                                    },
                                                    BackgroundColor = new DocumentModel.Color
                                                    {
                                                        Red = 255,
                                                        Blue = 64,
                                                        Green = 0
                                                    },
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Type = DocumentModel.BorderType.Thick,
                                                        Color = new DocumentModel.Color
                                                        {
                                                            Red = 255
                                                        }
                                                    },
                                                    BorderStyleDiagonalUpLeftToBottomRight = new DocumentModel.BorderStyle
                                                    {
                                                        Type = DocumentModel.BorderType.DashDotDot,
                                                        Color = new DocumentModel.Color
                                                        {
                                                            Green = 127
                                                        }
                                                    },
                                                    FontStyle = new DocumentModel.FontStyle
                                                    {
                                                        Color = new DocumentModel.Color
                                                        {
                                                            Alpha = 255,
                                                            Red = 127,
                                                            Blue = 127,
                                                            Green = 255
                                                        },
                                                        FontName = "Arial",
                                                        Underline = DocumentModel.UnderLineStyle.Double
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            }
                                        }
                                    },
                                    new Row
                                    {
                                        Style = new DocumentModel.CellStyle
                                        {
                                            BorderStyleBottom = new DocumentModel.BorderStyle
                                            {
                                                Type = DocumentModel.BorderType.Thick,
                                                Color = new DocumentModel.Color
                                                {
                                                    Red = 255
                                                }
                                            },
                                            FontStyle = new DocumentModel.FontStyle
                                            {
                                                Color = new DocumentModel.Color
                                                {
                                                    Alpha = 255,
                                                    Red = 127,
                                                    Blue = 127,
                                                    Green = 255
                                                },
                                                FontName = "Algerian"
                                            }
                                        },
                                        Cells = new List<Cell>
                                        {
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            }
                                        }
                                    },
                                    new Row
                                    {
                                        Cells = new List<Cell>
                                        {
                                            new Cell
                                            {
                                                Value = "Solid",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.Solid,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "ThreeQuarters",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.ThreeQuarters,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "OneHalf",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.OneHalf,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "OneQuarter",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.OneQuarter,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "OneEight",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.OneEight,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "OneSixteenth",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.OneSixteenth,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "HorizontalStripe",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.HorizontalStripe,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "VerticalStripe",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.VerticalStripe,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "ReverseDiagonalStripe",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.ReverseDiagonalStripe,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "DiagonalCrosshatch",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.DiagonalCrosshatch,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "ThickDiagonalCrosshatch",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.ThickDiagonalCrosshatch,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "ThinHorizontalStripe",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.ThinHorizontalStripe,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "ThinVerticalStripe",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.ThinVerticalStripe,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "ThinReverseDiagonalStripe",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.ThinReverseDiagonalStripe,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "ThinDiagonalStripe",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.ThinDiagonalStripe,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "ThinHorizontalCrosshatch",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.ThinHorizontalCrosshatch,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "ThinDiagonalCrosshatch",
                                                Type = CellType.Text,
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle = DocumentModel.FillPatternStyle.ThinDiagonalCrosshatch,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Alpha = 255,
                                                        Red = 255
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    },new Sheet
                    {
                        Content = new ReportSection
                        {
                            Header = new RowCollectionSection
                            {
                                Rows = new RowCollection
                                {
                                    new Row
                                    {
                                        Height = 125,
                                        Cells = new List<Cell>
                                        {
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced",
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Type = DocumentModel.BorderType.Thick,
                                                        Color = new DocumentModel.Color
                                                        {
                                                            Red = 255
                                                        }
                                                    },
                                                    FontStyle = new DocumentModel.FontStyle
                                                    {
                                                        Color = new DocumentModel.Color
                                                        {
                                                            Alpha = 255,
                                                            Red = 127,
                                                            Blue = 127,
                                                            Green = 255
                                                        },
                                                        FontName = "Algerian"
                                                    }
                                                }
                                            }
                                        }
                                    },
                                    new Row
                                    {
                                        Cells = new List<Cell>
                                        {
                                            null,
                                            new Cell
                                            {
                                                Value = "Testing",
                                                Style = new DocumentModel.CellStyle
                                                {
                                                    FillPatternStyle= DocumentModel.FillPatternStyle.VerticalStripe,
                                                    FillPatternColor = new DocumentModel.Color
                                                    {
                                                        Red = 64,
                                                        Blue = 255,
                                                        Green = 0
                                                    },
                                                    BackgroundColor = new DocumentModel.Color
                                                    {
                                                        Red = 255,
                                                        Blue = 64,
                                                        Green = 0
                                                    },
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Type = DocumentModel.BorderType.Thick,
                                                        Color = new DocumentModel.Color
                                                        {
                                                            Red = 255
                                                        }
                                                    },
                                                    BorderStyleDiagonalUpLeftToBottomRight = new DocumentModel.BorderStyle
                                                    {
                                                        Type = DocumentModel.BorderType.DashDotDot,
                                                        Color = new DocumentModel.Color
                                                        {
                                                            Green = 127
                                                        }
                                                    },
                                                    FontStyle = new DocumentModel.FontStyle
                                                    {
                                                        Color = new DocumentModel.Color
                                                        {
                                                            Alpha = 255,
                                                            Red = 127,
                                                            Blue = 127,
                                                            Green = 255
                                                        },
                                                        FontName = "Arial",
                                                        Underline = DocumentModel.UnderLineStyle.Double
                                                    }
                                                }
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            }
                                        }
                                    },
                                    new Row
                                    {
                                        Cells = new List<Cell>
                                        {
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            var data = new
            {
                Height = 48f,
                Str1 = DateTime.Today,
                Str2 = "Test2"
            };


            testing.Sheets.First().Content.Header.Rows.First().Bindings.Add(nameof(Row.Height), data, nameof(data.Height));
            testing.Sheets.First().Content.Header.Rows.First().Cells.First().Bindings.Add(nameof(Cell.Value), data, nameof(data.Str1));
            testing.Sheets.First().Content.Header.Rows.First().Cells.Skip(1).First().Bindings.Add(nameof(Cell.Value), data, nameof(data.Str2));

            testing.Sheets.First().Content.Header.Rows.Skip(1).First().Bindings.Add(nameof(Row.Height), data, nameof(data.Height));
            foreach (var cell in testing.Sheets.First().Content.Header.Rows.Skip(2).First().Cells)
            {
                cell.Bindings.Add(nameof(Cell.Value), data, nameof(data.Str2));
            }
            //System.Xml.Serialization.XmlSerializer reader = new System.Xml.Serialization.XmlSerializer(typeof(ReportDefinition));
            //reader.Serialize(File.Create("Def.xml"), testing);
            return testing;
        }

        private static ReportDefinition TestReport2()
        {
            var headerStyle = new DocumentModel.CellStyle
            {
                BorderStyleBottom = new DocumentModel.BorderStyle
                {
                    Type = DocumentModel.BorderType.Medium
                }
            };

            var report = new ReportDefinition
            {
                Sheets = new List<Sheet>
                {
                    new Sheet
                    {
                        Content = new ReportSection
                        {
                            Header =  new RowCollectionSection
                            {
                                Rows = new RowCollection
                                {
                                    new Row
                                    {
                                        Cells = new List<Cell>
                                        {
                                            new Cell
                                            {
                                                Value = "Int1",
                                                Style = headerStyle
                                            },
                                            new Cell
                                            {
                                                Value = "String1",
                                                Style = headerStyle
                                            },
                                            new Cell
                                            {
                                                Value = "DateTime1",
                                                Style = headerStyle
                                            },
                                            new Cell
                                            {
                                                Value = "Float1",
                                                Style = headerStyle
                                            }
                                        }
                                    }
                                }
                            },
                            SubSection = new ReportSection
                            {
                                Header = new RowCollectionSection
                                {
                                    Rows = new RowCollection
                                    {
                                        new Row
                                        {
                                            Cells = new List<Cell>
                                            {
                                                new Cell
                                                {
                                                    Value = "Insgdrgt1"
                                                },
                                                new Cell
                                                {
                                                    Value = "Steetrtertertrtring1"
                                                }
                                            }
                                        },
                                        new Row
                                        {
                                            Cells = new List<Cell>
                                            {
                                                null,
                                                null,
                                                new Cell
                                                {
                                                    Value = "DaertertseteTime1"
                                                },
                                                new Cell
                                                {
                                                    Value = "Floawetertt1"
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            var data = new List<TestPoco>
            {
                new TestPoco
                {
                    Int1 = 10,
                    String1 = "Test 1",
                    DateTime1 = DateTime.Now,
                    Float1 = 10f/3f
                },
                new TestPoco
                {
                    Int1 = 20,
                    String1 = "Test 20",
                    DateTime1 = DateTime.Now.AddDays(-10),
                    Float1 = 10f/4f
                },
                new TestPoco
                {
                    Int1 = 30,
                    String1 = "Test 300",
                    DateTime1 = DateTime.Now.AddDays(200),
                    Float1 = 10f/5f
                }
            };

            var datasource = new ObjectDataSourceBrowser(data);

            report.Sheets.First().Content.DataSource = datasource;

            report.Sheets.First().Content.SubSection.Header.Rows.First().Cells.Skip(0).First().Bindings.Add(new PropertyBinding("Value", datasource, nameof(TestPoco.Int1)));
            report.Sheets.First().Content.SubSection.Header.Rows.First().Cells.Skip(1).First().Bindings.Add(new PropertyBinding("Value", datasource, nameof(TestPoco.String1)));
            report.Sheets.First().Content.SubSection.Header.Rows.Skip(1).First().Cells.Skip(2).First().Bindings.Add(new PropertyBinding("Value", datasource, nameof(TestPoco.DateTime1)));
            report.Sheets.First().Content.SubSection.Header.Rows.Skip(1).First().Cells.Skip(3).First().Bindings.Add(new PropertyBinding("Value", datasource, nameof(TestPoco.Float1)));

            return report;
        }

        class TestPoco
        {
            public int Int1 { get; set; }

            public string String1 { get; set; }

            public DateTime DateTime1 { get; set; }

            public float Float1 { get; set; }
        }
    }
}
