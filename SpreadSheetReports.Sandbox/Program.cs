﻿using SpreadSheetsReports;
using SpreadSheetsReports.NPOIRenderer;
using SpreadSheetsReports.Renderer;
using SpreadSheetsReports.ReportModel;
using SpreadSheetsReports.SpreadsheetLightRenderer;
using System;
using System.Collections;
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

            definition = TestReport3();
            renderer = new NPOIRenderer.NPOIRenderer();
            stream = File.Create("NestedNpoi.xlsx");
            renderer.Render(definition).CopyTo(stream);
            stream.Flush();
            stream.Close();

            definition = TestReport3();
            renderer = new SpreadsheetLightRenderer.SpreadsheetLightRenderer();
            stream = File.Create("NestedSL.xlsx");
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
                                                        Color = new DocumentModel.Color(0,255,0)
                                                    },
                                                    FontStyle = new DocumentModel.FontStyle
                                                    {
                                                        Color = new DocumentModel.Color(255, 127, 127, 255),
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
                                                    FillPatternColor = new DocumentModel.Color(
                                                        red:64,
                                                        blue:255,
                                                        green:0
                                                    ),
                                                    BackgroundColor = new DocumentModel.Color
                                                    (
                                                        red:255,
                                                        blue:64,
                                                        green:0
                                                    ),
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Type = DocumentModel.BorderType.Thick,
                                                        Color = new DocumentModel.Color
                                                        (
                                                            red:255
                                                        )
                                                    },
                                                    BorderStyleDiagonalUpLeftToBottomRight = new DocumentModel.BorderStyle
                                                    {
                                                        Type = DocumentModel.BorderType.DashDotDot,
                                                        Color = new DocumentModel.Color
                                                        (
                                                            green:127
                                                        )
                                                    },
                                                    FontStyle = new DocumentModel.FontStyle
                                                    {
                                                        Color = new DocumentModel.Color
                                                        (
                                                            alpha:255,
                                                            red:127,
                                                            blue:127,
                                                            green:255
                                                        ),
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
                                                (
                                                    red:255
                                                )
                                            },
                                            FontStyle = new DocumentModel.FontStyle
                                            {
                                                Color = new DocumentModel.Color
                                                (
                                                    alpha:255,
                                                    red:127,
                                                    blue:127,
                                                    green:255
                                                ),
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                    (
                                                        alpha:255,
                                                        red:255
                                                    )
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
                                                        (
                                                            red:255
                                                        )
                                                    },
                                                    FontStyle = new DocumentModel.FontStyle
                                                    {
                                                        Color = new DocumentModel.Color
                                                        (
                                                            alpha:255,
                                                            red:127,
                                                            blue:127,
                                                            green:255
                                                        ),
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
                                                    (
                                                        red:64,
                                                        blue:255,
                                                        green:0
                                                    ),
                                                    BackgroundColor = new DocumentModel.Color
                                                    (
                                                        red:255,
                                                        blue:64,
                                                        green:0
                                                    ),
                                                    BorderStyleBottom = new DocumentModel.BorderStyle
                                                    {
                                                        Type = DocumentModel.BorderType.Thick,
                                                        Color = new DocumentModel.Color
                                                        (
                                                            red:255
                                                        )
                                                    },
                                                    BorderStyleDiagonalUpLeftToBottomRight = new DocumentModel.BorderStyle
                                                    {
                                                        Type = DocumentModel.BorderType.DashDotDot,
                                                        Color = new DocumentModel.Color
                                                        (
                                                            green:127
                                                        )
                                                    },
                                                    FontStyle = new DocumentModel.FontStyle
                                                    {
                                                        Color = new DocumentModel.Color
                                                        (
                                                            alpha:255,
                                                            red:127,
                                                            blue:127,
                                                            green:255
                                                        ),
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
            var reportData = Enumerable.Range(1, 3).Select(i => new
            {
                Data = TestPocoData( i * 100 + 1, i * 100 + 100),
                Name = $"Page {i}"
            }).ToList();

            var datasource = new ObjectDataSourceBrowser(reportData);
            var contentSource = new ObjectDataSourceBrowser(datasource, "Data");

            var headerStyle = new DocumentModel.CellStyle
            {
                BorderStyleBottom = new DocumentModel.BorderStyle
                {
                    Type = DocumentModel.BorderType.Medium
                },
                FontStyle = new DocumentModel.FontStyle
                {
                    Color = new DocumentModel.Color
                    (
                        red:127
                    )
                }
            };

            var report = new ReportDefinition
            {
                Sheets = new List<Sheet>
                {
                    new Sheet
                    {
                        Bindings = new PropertyBindingCollection
                        {
                            new PropertyBinding
                            {
                                DataSource = datasource,
                                DataMember = "Name",
                                PropertyName = nameof(Sheet.Name)
                            }
                        },
                        Content = new ReportSection
                        {
                            DataSource = contentSource,
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
                                                    Style = new DocumentModel.CellStyle
                                                    {
                                                        FontStyle = new DocumentModel.FontStyle
                                                        {
                                                            IsItalic = true
                                                        }
                                                    },
                                                    Bindings = new PropertyBindingCollection
                                                    {
                                                        new PropertyBinding(nameof(Cell.Value), contentSource, nameof(TestPoco.Int1))
                                                    }
                                                },
                                                new Cell
                                                {
                                                    Style = new DocumentModel.CellStyle
                                                    {
                                                        FontStyle = new DocumentModel.FontStyle
                                                        {
                                                            IsBold = true
                                                        }
                                                    },
                                                    Bindings = new PropertyBindingCollection
                                                    {
                                                        new PropertyBinding(nameof(Cell.Value), contentSource, nameof(TestPoco.String1))
                                                    }
                                                },
                                        //    }
                                        //},
                                        //new Row
                                        //{
                                        //    Cells = new List<Cell>
                                        //    {
                                        //        null,
                                        //        null,
                                                new Cell
                                                {
                                                    Style = new DocumentModel.CellStyle
                                                    {
                                                        FontStyle = new DocumentModel.FontStyle
                                                        {
                                                            IsStrikeout = true
                                                        }
                                                    },
                                                    Bindings = new PropertyBindingCollection
                                                    {
                                                        new PropertyBinding(nameof(Cell.Value), contentSource, nameof(TestPoco.DateTime1))
                                                    }
                                                },
                                                new Cell
                                                {
                                                    Style = new DocumentModel.CellStyle
                                                    {
                                                        FontStyle = new DocumentModel.FontStyle
                                                        {
                                                            ScriptStyle = DocumentModel.FontScriptStyle.Superscript
                                                        }
                                                    },
                                                    Bindings = new PropertyBindingCollection
                                                    {
                                                        new PropertyBinding(nameof(Cell.Value), contentSource, nameof(TestPoco.Float1))
                                                    }
                                                }
                                            }
                                        },
                                        //null,
                                        //null
                                    }
                                }
                            }
                        }
                    }
                }
            };
           

            report.DataSource = datasource;

            return report;
        }

        private static ReportDefinition TestReport3()
        {
            var reportData = Enumerable.Range(1, 3).Select(i => new
            {
                Data = NestedPocoData(i * 10 + 1, i * 10 + 10, 1, 10),
                Name = $"Page {i}"
            }).ToList();

            var datasource = new ObjectDataSourceBrowser(reportData);
            var headerSource = new ObjectDataSourceBrowser(datasource, "Data");
            var contentSource = new ObjectDataSourceBrowser(headerSource, "Inner");

            var subHeaderStyle = new DocumentModel.CellStyle
            {
                BorderStyleBottom = new DocumentModel.BorderStyle
                {
                    Type = DocumentModel.BorderType.SlantedDashDot
                },
                FontStyle = new DocumentModel.FontStyle
                {
                    Color = new DocumentModel.Color
                    (
                        red:127
                    )
                }
            };

            var mainHeaderStyle = new DocumentModel.CellStyle
            {
                BorderStyleBottom = new DocumentModel.BorderStyle
                {
                    Type = DocumentModel.BorderType.Thick
                },
                HorizontalAlignment = DocumentModel.HorizontalAlignment.Right,
                FontStyle = new DocumentModel.FontStyle
                {
                    Color = new DocumentModel.Color
                    (
                        green:127
                    )
                }
            };

            var report = new ReportDefinition
            {
                Sheets = new List<Sheet>
                {
                    new Sheet
                    {
                        Bindings = new PropertyBindingCollection
                        {
                            new PropertyBinding
                            {
                                DataSource = datasource,
                                DataMember = "Name",
                                PropertyName = nameof(Sheet.Name)
                            }
                        },
                        Content = new ReportSection
                        {
                            DataSource = headerSource,
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
                                                Value ="Header1"
                                            },
                                            new Cell
                                            {
                                                Value ="Header2"
                                            }
                                        }
                                    }
                                }
                            },
                            SubSection = new ReportSection
                            {
                                DataSource = contentSource,
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
                                                    Style = mainHeaderStyle,
                                                    Bindings = new PropertyBindingCollection
                                                    {
                                                        new PropertyBinding(nameof(Cell.Value), headerSource, nameof(NestedPoco.Header1))
                                                    }
                                                },
                                                new Cell
                                                {
                                                    Style = mainHeaderStyle,
                                                    Bindings = new PropertyBindingCollection
                                                    {
                                                        new PropertyBinding(nameof(Cell.Value), headerSource, nameof(NestedPoco.Header2))
                                                    }
                                                }
                                            }
                                        },
                                        new Row
                                        {
                                            Cells = new List<Cell>
                                            {
                                                new Cell
                                                {
                                                    Value = "Int1",
                                                    Style = subHeaderStyle
                                                },
                                                new Cell
                                                {
                                                    Value = "String1",
                                                    Style = subHeaderStyle
                                                },
                                                new Cell
                                                {
                                                    Value = "DateTime1",
                                                    Style = subHeaderStyle
                                                },
                                                new Cell
                                                {
                                                    Value = "Float1",
                                                    Style = subHeaderStyle
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
                                                        Style = new DocumentModel.CellStyle
                                                        {
                                                            FontStyle = new DocumentModel.FontStyle
                                                            {
                                                                IsItalic = true
                                                            }
                                                        },
                                                        Bindings = new PropertyBindingCollection
                                                        {
                                                            new PropertyBinding(nameof(Cell.Value), contentSource, nameof(TestPoco.Int1))
                                                        }
                                                    },
                                                    new Cell
                                                    {
                                                        Style = new DocumentModel.CellStyle
                                                        {
                                                            FontStyle = new DocumentModel.FontStyle
                                                            {
                                                                IsBold = true
                                                            }
                                                        },
                                                        Bindings = new PropertyBindingCollection
                                                        {
                                                            new PropertyBinding(nameof(Cell.Value), contentSource, nameof(TestPoco.String1))
                                                        }
                                                    },
                                                    new Cell
                                                    {
                                                        Style = new DocumentModel.CellStyle
                                                        {
                                                            FontStyle = new DocumentModel.FontStyle
                                                            {
                                                                IsStrikeout = true
                                                            }
                                                        },
                                                        Bindings = new PropertyBindingCollection
                                                        {
                                                            new PropertyBinding(nameof(Cell.Value), contentSource, nameof(TestPoco.DateTime1))
                                                        }
                                                    },
                                                    new Cell
                                                    {
                                                        Style = new DocumentModel.CellStyle
                                                        {
                                                            FontStyle = new DocumentModel.FontStyle
                                                            {
                                                                ScriptStyle = DocumentModel.FontScriptStyle.Superscript
                                                            }
                                                        },
                                                        Bindings = new PropertyBindingCollection
                                                        {
                                                            new PropertyBinding(nameof(Cell.Value), contentSource, nameof(TestPoco.Float1))
                                                        }
                                                    }
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


            report.DataSource = datasource;

            return report;
        }

        private static List<NestedPoco> NestedPocoData(int headerStart, int headerEnd, int innerStart, int innerEnd)
        {
            var data = new List<NestedPoco>();
            for (int i = headerStart; i < headerEnd; i++)
            {
                data.Add(new NestedPoco
                {
                    Header1 = i,
                    Header2 = $"Header {i}",
                    Inner = TestPocoData(10 * i + innerStart, 10 * i + innerEnd)
                });
            }
            return data;
        }

        private static List<TestPoco> TestPocoData(int start, int end)
        {
            var data = new List<TestPoco>();
            for (int i = start; i < end + 1; i++)
            {
                data.Add(new TestPoco
                {
                    Int1 = i * end,
                    String1 = $"Test {i}",
                    DateTime1 = DateTime.Now.AddHours(i),
                    Float1 = (float)end / i
                });
            }
            return data;
        }

        class NestedPoco
        {
            public int Header1 { get; set; }
            public string Header2 { get; set; }
            public IEnumerable<TestPoco> Inner { get; set; }
        
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
