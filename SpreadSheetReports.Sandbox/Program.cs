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

            renderer = new SpreadsheetLightRenderer.SpreadsheetLightRenderer();
            stream = File.Create("Workbook2.xlsx");
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
                                                            Blue = 255,
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
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
                                                Value = "Testing"
                                            },
                                            new Cell
                                            {
                                                Value = "Testig33"
                                            },
                                            new Cell
                                            {
                                                Value = "NotReplaced"
                                            },
                                            new Cell
                                            {
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
            System.Xml.Serialization.XmlSerializer reader = new System.Xml.Serialization.XmlSerializer(typeof(ReportDefinition));
            reader.Serialize(File.Create("Def.xml"), testing);
            return testing;
        }
    }
}
