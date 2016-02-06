﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetLight;
using SpreadSheetsReports.DocumentModel;

namespace SpreadSheetsReports.SpreadsheetLightRenderer
{
    class SpreadsheetLightRenderer : IReportRenderer
    {
        public void Render(ReportDefinition report)
        {
            using (SLDocument document = new SLDocument())
            {
                int rowCounter = 0;
                foreach (var row in report.Content.Header.Rows)
                {
                    RenderRow(document, row,rowCounter);
                    rowCounter++;
                }
            }
        }

        private void RenderRow(SLDocument document, Row row, int rowCounter)
        {
            int cellCounter = 0;
            foreach (var cell in row.Cells)
            {
                RenderCell(document, cell, rowCounter, cellCounter);
                ApplyStyle(document, cell, rowCounter, cellCounter);
                cellCounter++;
            }
        }

        private void ApplyStyle(SLDocument document, Cell cell, int rowCounter, int cellCounter)
        {
            SLStyle style = document.CreateStyle();
            ApplyStyle(style, cell.Style);
        }

        private void ApplyStyle(SLStyle cellStyle, CellStyle style)
        {
            ApplyHorizontalAlignment(cellStyle, style);
        }

        private void ApplyHorizontalAlignment(SLStyle cellStyle, CellStyle style)
        {
            switch (style.HorizontalAlignment)
            {
                case HorizontalAlignment.General:
                    cellStyle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General);
                    break;
                case HorizontalAlignment.Left:

                    break;
                case HorizontalAlignment.Center:
                    break;
                case HorizontalAlignment.Right:
                    break;
                case HorizontalAlignment.Fill:
                    break;
                case HorizontalAlignment.Justify:
                    break;
                case HorizontalAlignment.CenterAcrossSelection:
                    break;
                case HorizontalAlignment.Distributed:
                    break;
                default:
                    break;
            }
        }

        private void RenderCell(SLDocument document, Cell cell, int rowCounter, int cellCounter)
        {
            switch (Type.GetTypeCode(cell.Value.GetType()))
            {
                case TypeCode.Boolean:
                    document.SetCellValue(rowCounter, cellCounter, Convert.ToBoolean(cell.Value));
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
                    document.SetCellValue(rowCounter, cellCounter, Convert.ToDouble(cell.Value));
                    break;
                case TypeCode.DateTime:
                    document.SetCellValue(rowCounter, cellCounter, Convert.ToDateTime(cell.Value));
                    break;
                case TypeCode.Empty:
                case TypeCode.Object:
                case TypeCode.String:
                case TypeCode.DBNull:
                case TypeCode.Char:
                default:
                    document.SetCellValue(rowCounter, cellCounter, cell.Value.ToString());
                    break;
            }
        }
    }
}