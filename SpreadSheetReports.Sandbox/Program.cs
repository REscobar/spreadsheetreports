using SpreadSheetsReports.NPOIRenderer;
using SpreadSheetsReports.ReportModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetReports.Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {

            var testing = new ReportDefinition
            {
                Sheets = new List<Sheet>
                {
                    new Sheet
                    {
                        Content = new ReportSection
                        {
                            Header = new HeaderSection
                            {
                                Rows = new RowCollection
                                {
                                    new Row
                                    {
                                        Cells = new List<Cell>
                                        {
                                            new Cell
                                            {
                                                Value = "Testing"
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            System.Xml.Serialization.XmlSerializer reader = new System.Xml.Serialization.XmlSerializer(typeof(ReportDefinition));
            var definition = reader.Deserialize(File.OpenRead("TestData.xml")) as ReportDefinition;

            reader.Serialize(File.Create("Def.xml"), testing);

            definition.DataSource = new { Some = "" };

            var renderer = new NPOIRenderer();
            var stream = File.Create("Workbook1.xlsx");
            renderer.Render(definition).CopyTo(stream);
            stream.Flush();
            stream.Close();

            
        }
    }
}
