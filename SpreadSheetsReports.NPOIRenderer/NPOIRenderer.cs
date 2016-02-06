using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadSheetsReports.DocumentModel;

namespace SpreadSheetsReports.NPOIRenderer
{
    class NPOIRenderer : IReportRenderer
    {
        public void Render(ReportDefinition report)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            int rowCounter = 0;
            foreach (var row in report.Content.Header.Rows)
            {
              
                IRow sheetrow = sheet.CreateRow(rowCounter);
                RenderRow(sheetrow, row);
               
                rowCounter++;
            }
        }

        private void RenderRow(IRow sheetrow, Row row)
        {
            int cellCounter = 0;
            foreach (var cell in row.Cells)
            {
                ICell sheetCell = sheetrow.CreateCell(cellCounter);
                RenderCell(sheetCell, cell);
                ApplyStyle(sheetCell, cell);

                cellCounter++;
            }
        }

        private void RenderCell(ICell sheetCell, Cell cell)
        {
            switch (Type.GetTypeCode( cell.Value.GetType()))
            {
                case TypeCode.Boolean:
                    sheetCell.SetCellValue(Convert.ToBoolean(cell.Value));
                    break;
                case TypeCode.SByte:
                case TypeCode.Byte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    sheetCell.SetCellValue(Convert.ToDouble(cell.Value));
                    break;
                case TypeCode.DateTime:
                    sheetCell.SetCellValue(Convert.ToDateTime(cell.Value));
                    break;
                case TypeCode.Empty:
                case TypeCode.Object:
                case TypeCode.String:
                case TypeCode.DBNull:
                case TypeCode.Char:
                default:
                    sheetCell.SetCellValue(cell.Value.ToString());
                    break;
            }

        }

        private void ApplyStyle(ICell sheetCell, Cell cell)
        {
            ApplyStyle(sheetCell.CellStyle, cell.Style);
        }

        private void ApplyStyle(ICellStyle cellStyle, CellStyle style)
        {
            ApplyHorizontalAlignment(cellStyle, style);
            ApplyVerticalAlignment(cellStyle, style);
        }

        private void ApplyVerticalAlignment(ICellStyle cellStyle, CellStyle style)
        {
            switch (style.VerticalAlignment)
            {
                case DocumentModel.VerticalAlignment.Top:
                    cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Top;
                    break;
                case DocumentModel.VerticalAlignment.Center:
                    cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                    break;
                case DocumentModel.VerticalAlignment.Justify:
                    cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Justify;
                    break;
                case DocumentModel.VerticalAlignment.Distributed:
                    cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Distributed;
                    break;
                default:
                    break;
            }
        }


        private void ApplyHorizontalAlignment(ICellStyle cellStyle, CellStyle style)
        {
            switch (style.HorizontalAlignment)
            {
                case DocumentModel.HorizontalAlignment.General:
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.General;
                    break;
                case DocumentModel.HorizontalAlignment.Left:
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                    break;
                case DocumentModel.HorizontalAlignment.Center:
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    break;
                case DocumentModel.HorizontalAlignment.Right:
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
                    break;
                case DocumentModel.HorizontalAlignment.Fill:
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Fill;
                    break;
                case DocumentModel.HorizontalAlignment.Justify:
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Justify;
                    break;
                case DocumentModel.HorizontalAlignment.CenterAcrossSelection:
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CenterSelection;
                    break;
                case DocumentModel.HorizontalAlignment.Distributed:
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Distributed;
                    break;
                default:
                    throw new InvalidOperationException("Unknown Horizontal Alignment value: " + style.HorizontalAlignment.ToString());
            }
        }
    }
}
