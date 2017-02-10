namespace SpreadSheetsReports.NPOIRenderer
{
    using System;
    using System.IO;
    using DocumentModel;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using Renderer;

    public class NPOIRenderer : BaseReportRenderer
    {
        protected override Stream RenderToStream(Document document)
        {
            IWorkbook workbook = new XSSFWorkbook();

            foreach (var sheet in document.Sheets)
            {
                ISheet documentSheet = sheet.Name != null ? workbook.CreateSheet(sheet.Name) : workbook.CreateSheet();
                int rowCounter = 0;
                foreach (var row in sheet.Rows)
                {
                    IRow sheetrow = documentSheet.CreateRow(rowCounter);
                    rowCounter++;

                    if (row == null)
                    {
                        continue;
                    }

                    this.RenderRow(sheetrow, row);
                }
            }

            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Stream stream = File.Create(path);

            workbook.Write(stream);

            stream = new MemoryStream(File.ReadAllBytes(path));

            return stream;
        }

        private void RenderRow(IRow sheetrow, Row row)
        {
            int cellCounter = 0;

            if (row.Height.HasValue)
            {
                sheetrow.HeightInPoints = row.Height.Value;
            }

            foreach (var cell in row.Cells)
            {
                ICell sheetCell = sheetrow.CreateCell(cellCounter);
                cellCounter++;

                if (cell == null)
                {
                    continue;
                }

                this.RenderCell(sheetCell, cell);
                this.ApplyStyle(sheetCell, cell);
            }
        }

        private void RenderCell(ICell sheetCell, Cell cell)
        {
            switch (Type.GetTypeCode(cell.Value.GetType()))
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
            this.ApplyStyle(sheetCell.CellStyle, cell.Style);
        }

        private void ApplyStyle(ICellStyle cellStyle, CellStyle style)
        {
            this.ApplyHorizontalAlignment(cellStyle, style.HorizontalAlignment);
            this.ApplyVerticalAlignment(cellStyle, style.VerticalAlignment);
            this.ApplyBorderStyleTop(cellStyle, style.BorderStyleTop);
            this.ApplyBorderStyleBottom(cellStyle, style.BorderStyleBottom);
            this.ApplyBorderStyleLeft(cellStyle, style.BorderStyleLeft);
            this.ApplyBorderStyleRight(cellStyle, style.BorderStyleRight);

            //this.ApplyBorderDiagonalStyle(cellStyle, style.BorderStyleDiagonalUpLeftToBottomRight, style.BorderStyleDiagonalUpRightToBottomLeft);
            //this.ApplyFont(cellStyle, style.Font);
        }

        private void ApplyBorderStyleRight(ICellStyle cellStyle, DocumentModel.BorderStyle borderStyleRight)
        {
            if (borderStyleRight == null)
            {
                return;
            }

            cellStyle.BorderRight = this.GetBorderStyle(borderStyleRight.Type);
            XSSFCellStyle style = cellStyle as XSSFCellStyle;
            var color = borderStyleRight.Color.Value;
            style.RightBorderXSSFColor.SetRgb(new[] { color.Red, color.Green, color.Blue });
        }

        private void ApplyBorderStyleLeft(ICellStyle cellStyle, DocumentModel.BorderStyle borderStyleLeft)
        {
            if (borderStyleLeft == null)
            {
                return;
            }

            cellStyle.BorderLeft = this.GetBorderStyle(borderStyleLeft.Type);
            XSSFCellStyle style = cellStyle as XSSFCellStyle;
            var color = borderStyleLeft.Color.Value;
            style.RightBorderXSSFColor.SetRgb(new[] { color.Red, color.Green, color.Blue });
        }

        private void ApplyBorderStyleBottom(ICellStyle cellStyle, DocumentModel.BorderStyle borderStyleBottom)
        {
            if (borderStyleBottom == null)
            {
                return;
            }

            cellStyle.BorderBottom = this.GetBorderStyle(borderStyleBottom.Type);
            XSSFCellStyle style = cellStyle as XSSFCellStyle;
            var color = borderStyleBottom.Color.Value;
            style.RightBorderXSSFColor.SetRgb(new[] { color.Red, color.Green, color.Blue });
        }

        private void ApplyBorderStyleTop(ICellStyle cellStyle, DocumentModel.BorderStyle borderStyleTop)
        {
            if (borderStyleTop == null)
            {
                return;
            }

            cellStyle.BorderTop = this.GetBorderStyle(borderStyleTop.Type);
            XSSFCellStyle style = cellStyle as XSSFCellStyle;
            var color = borderStyleTop.Color.Value;
            style.RightBorderXSSFColor.SetRgb(new[] { color.Red, color.Green, color.Blue });
        }

        private void ApplyBorderDiagonalStyle(ICellStyle cellStyle, DocumentModel.BorderStyle borderStyleDiagonalUpLeftToBottomRight, DocumentModel.BorderStyle borderStyleDiagonalUpRightToBottomLeft)
        {
            if (borderStyleDiagonalUpLeftToBottomRight == null || borderStyleDiagonalUpRightToBottomLeft == null)
            {
                return;
            }
            if(borderStyleDiagonalUpLeftToBottomRight.Type != BorderType.None && borderStyleDiagonalUpRightToBottomLeft.Type != BorderType.None)
            {
                cellStyle.BorderDiagonal = BorderDiagonal.Both;
            }
            else if(borderStyleDiagonalUpLeftToBottomRight.Type != BorderType.None)
            {
                cellStyle.BorderDiagonal = BorderDiagonal.Backward;
            }
            else if(borderStyleDiagonalUpRightToBottomLeft.Type != BorderType.None)
            {
                cellStyle.BorderDiagonal = BorderDiagonal.Forward;
            }
        }

        private NPOI.SS.UserModel.BorderStyle GetBorderStyle(DocumentModel.BorderType borderStyle)
        {
            switch (borderStyle)
            {
                case BorderType.None:
                    return NPOI.SS.UserModel.BorderStyle.None;
                case BorderType.Thin:
                    return NPOI.SS.UserModel.BorderStyle.Thin;
                case BorderType.Medium:
                    return NPOI.SS.UserModel.BorderStyle.Medium;
                case BorderType.Dashed:
                    return NPOI.SS.UserModel.BorderStyle.Dashed;
                case BorderType.Dotted:
                    return NPOI.SS.UserModel.BorderStyle.Dotted;
                case BorderType.Thick:
                    return NPOI.SS.UserModel.BorderStyle.Thick;
                case BorderType.Double:
                    return NPOI.SS.UserModel.BorderStyle.Double;
                case BorderType.Hair:
                    return NPOI.SS.UserModel.BorderStyle.Hair;
                case BorderType.MediumDashed:
                    return NPOI.SS.UserModel.BorderStyle.MediumDashed;
                case BorderType.DashDot:
                    return NPOI.SS.UserModel.BorderStyle.DashDot;
                case BorderType.MediumDashDot:
                    return NPOI.SS.UserModel.BorderStyle.MediumDashDot;
                case BorderType.DashDotDot:
                    return NPOI.SS.UserModel.BorderStyle.DashDotDot;
                case BorderType.MediumDashDotDot:
                    return NPOI.SS.UserModel.BorderStyle.MediumDashDotDot;
                case BorderType.SlantedDashDot:
                    return NPOI.SS.UserModel.BorderStyle.SlantedDashDot;
                default:
                    throw new ArgumentOutOfRangeException(nameof(borderStyle));
            }
        }

        private void ApplyFont(ICellStyle cellStyle, FontStyle font)
        {
            throw new NotImplementedException();
        }

        private void ApplyVerticalAlignment(ICellStyle cellStyle, DocumentModel.VerticalAlignment style)
        {
            switch (style)
            {
                case DocumentModel.VerticalAlignment.Top:
                    cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Top;
                    break;
                case DocumentModel.VerticalAlignment.Center:
                    cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                    break;
                case DocumentModel.VerticalAlignment.Bottom:
                    cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Bottom;
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

        private void ApplyHorizontalAlignment(ICellStyle cellStyle, DocumentModel.HorizontalAlignment style)
        {
            switch (style)
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
                    throw new InvalidOperationException("Unknown Horizontal Alignment value: " + style.ToString());
            }
        }
    }
}
