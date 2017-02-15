namespace SpreadSheetsReports.SpreadsheetLightRenderer
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using DocumentFormat.OpenXml.Spreadsheet;
    using Renderer;
    using SpreadsheetLight;
    using SpreadSheetsReports.DocumentModel;

    public class SpreadsheetLightRenderer : BaseReportRenderer
    {
        protected override Stream RenderToStream(Document document)
        {
            using (SLDocument slDocument = new SLDocument())
            {

                var sheets = document.Sheets.ToList();

                if (!string.IsNullOrWhiteSpace(sheets[0].Name))
                {
                    slDocument.RenameWorksheet(slDocument.GetCurrentWorksheetName(), sheets[0].Name);
                }

                int rowCounter = 0;
                foreach (var row in sheets[0].Rows)
                {
                    rowCounter++;
                    this.RenderRow(slDocument, row, rowCounter);
                }

                var sheetIndex = 1;
                foreach (var sheet in sheets.Skip(1))
                {
                    var res = slDocument.AddWorksheet($"Sheet{sheetIndex++}");

                    if (!string.IsNullOrWhiteSpace(sheet.Name))
                    {
                        slDocument.RenameWorksheet(slDocument.GetCurrentWorksheetName(), sheet.Name);
                    }

                    rowCounter = 0;
                    foreach (var row in sheet.Rows)
                    {
                        rowCounter++;
                        this.RenderRow(slDocument, row, rowCounter);
                    }
                }

                var stream = new MemoryStream();
                slDocument.SaveAs(stream);

                stream.Seek(0, SeekOrigin.Begin);

                return stream;
            }
        }

        private void RenderRow(SLDocument document, DocumentModel.Row row, int rowCounter)
        {
            int cellCounter = 0;

            if (row.Height.HasValue)
            {
                document.SetRowHeight(rowCounter, row.Height.Value);
            }

            foreach (var cell in row.Cells)
            {
                cellCounter++;
                if (cell != null)
                {
                    this.RenderCell(document, cell, rowCounter, cellCounter);
                    this.ApplyStyle(document, cell, rowCounter, cellCounter);
                }
            }
        }

        private void ApplyStyle(SLDocument document, DocumentModel.Cell cell, int rowCounter, int cellCounter)
        {
            SLStyle style = document.CreateStyle();
            this.ApplyStyle(style, cell.Style);
            document.SetCellStyle(rowCounter, cellCounter, style);

        }

        private void ApplyStyle(SLStyle cellStyle, DocumentModel.CellStyle style)
        {
            this.ApplyHorizontalAlignment(cellStyle, style.HorizontalAlignment);
            this.ApplyVerticalStyle(cellStyle, style.VerticalAlignment);
            this.ApplyBorderStyle(cellStyle.Border.TopBorder, style.BorderStyleTop);
            this.ApplyBorderStyle(cellStyle.Border.LeftBorder, style.BorderStyleLeft);
            this.ApplyBorderStyle(cellStyle.Border.RightBorder, style.BorderStyleRight);
            this.ApplyBorderStyle(cellStyle.Border.BottomBorder, style.BorderStyleBottom);

            this.ApplyBorderStyle(cellStyle.Border.DiagonalBorder, style.BorderStyleDiagonalUpLeftToBottomRight);
            this.ApplyBorderStyle(cellStyle.Border.DiagonalBorder, style.BorderStyleDiagonalUpRightToBottomLeft);
            this.ApplyBorderDiagonalStyle(cellStyle.Border, style.BorderStyleDiagonalUpLeftToBottomRight, style.BorderStyleDiagonalUpRightToBottomLeft);

            this.ApplyFont(cellStyle.Font, style.FontStyle);
            this.ApplyFill(cellStyle.Fill, style.FillPatternStyle, style.FillPatternColor, style.BackgroundColor);
            this.ApplyMisc(cellStyle, style);
        }

        private void ApplyMisc(SLStyle cellStyle, DocumentModel.CellStyle style)
        {
            cellStyle.Alignment.Indent = style.Indent;
            cellStyle.Alignment.WrapText = style.WrapText;
            cellStyle.Alignment.TextRotation = style.Rotation;
            cellStyle.Alignment.ShrinkToFit = style.ShrinkToFit;
        }

        private void ApplyFill(SLFill fill, FillPatternStyle? fillPatternStyle, DocumentModel.Color? fillPatternColor, DocumentModel.Color? backgroundColor)
        {
            if (fillPatternStyle.HasValue)
            {
                fill.SetPatternType(this.GetFillPattern(fillPatternStyle.Value));
            }

            if (fillPatternColor.HasValue)
            {
                var color = fillPatternColor.Value;
                fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(color.Alpha, color.Red, color.Green, color.Blue));
            }

            if (backgroundColor.HasValue)
            {
                var color = backgroundColor.Value;
                fill.SetPatternBackgroundColor(System.Drawing.Color.FromArgb(color.Alpha, color.Red, color.Green, color.Blue));

            }
        }

        private PatternValues GetFillPattern(FillPatternStyle fillPatterStyle)
        {
            switch (fillPatterStyle)
            {
                case FillPatternStyle.Solid:
                    return PatternValues.Solid;
                case FillPatternStyle.ThreeQuarters:
                    return PatternValues.DarkGray;
                case FillPatternStyle.OneHalf:
                    return PatternValues.MediumGray;
                case FillPatternStyle.OneQuarter:
                    return PatternValues.LightGray;
                case FillPatternStyle.OneEight:
                    return PatternValues.Gray125;
                case FillPatternStyle.OneSixteenth:
                    return PatternValues.Gray0625;
                case FillPatternStyle.HorizontalStripe:
                    return PatternValues.DarkHorizontal;
                case FillPatternStyle.VerticalStripe:
                    return PatternValues.DarkVertical;
                case FillPatternStyle.ReverseDiagonalStripe:
                    return PatternValues.DarkDown;
                case FillPatternStyle.DiagonalStripe:
                    return PatternValues.DarkUp;
                case FillPatternStyle.DiagonalCrosshatch:
                    return PatternValues.DarkGrid;
                case FillPatternStyle.ThickDiagonalCrosshatch:
                    return PatternValues.DarkTrellis;
                case FillPatternStyle.ThinHorizontalStripe:
                    return PatternValues.LightHorizontal;
                case FillPatternStyle.ThinVerticalStripe:
                    return PatternValues.LightVertical;
                case FillPatternStyle.ThinReverseDiagonalStripe:
                    return PatternValues.LightDown;
                case FillPatternStyle.ThinDiagonalStripe:
                    return PatternValues.LightUp;
                case FillPatternStyle.ThinHorizontalCrosshatch:
                    return PatternValues.LightGrid;
                case FillPatternStyle.ThinDiagonalCrosshatch:
                    return PatternValues.LightTrellis;
                default:
                    throw new ArgumentOutOfRangeException(nameof(fillPatterStyle));
            }
        }

        private void ApplyFont(SLFont font, FontStyle fontStyle)
        {
            if (fontStyle == null)
            {
                return;
            }

            font.FontName = fontStyle.FontName;
            font.Bold = fontStyle.IsBold;
            font.Italic = fontStyle.IsItalic;
            font.Strike = fontStyle.IsStrikeout;

            font.VerticalAlignment = this.GetFontSuperScript(fontStyle.ScriptStyle);
            font.Underline = this.GetFontUnderLine(fontStyle.Underline);

            if (fontStyle.Color.HasValue)
            {
                var color = fontStyle.Color.Value;
                font.FontColor = System.Drawing.Color.FromArgb(color.Alpha, color.Red, color.Green, color.Blue);
            }
        }

        private UnderlineValues GetFontUnderLine(UnderLineStyle underline)
        {
            switch (underline)
            {
                case UnderLineStyle.None:
                    return UnderlineValues.None;
                case UnderLineStyle.Single:
                    return UnderlineValues.Single;
                case UnderLineStyle.Double:
                    return UnderlineValues.Double;
                case UnderLineStyle.SingleAccounting:
                    return UnderlineValues.SingleAccounting;
                case UnderLineStyle.DoubleAccounting:
                    return UnderlineValues.DoubleAccounting;
                default:
                    throw new ArgumentOutOfRangeException(nameof(underline));
            }
        }

        private VerticalAlignmentRunValues GetFontSuperScript(FontScriptStyle scriptStyle)
        {
            switch (scriptStyle)
            {
                case FontScriptStyle.None:
                    return VerticalAlignmentRunValues.Baseline;
                case FontScriptStyle.Superscript:
                    return VerticalAlignmentRunValues.Superscript;
                case FontScriptStyle.Subscript:
                    return VerticalAlignmentRunValues.Subscript;
                default:
                    throw new ArgumentOutOfRangeException(nameof(scriptStyle));
            }
        }

        private void ApplyBorderDiagonalStyle(SLBorder border, BorderStyle borderStyleDiagonalUpLeftToBottomRight, BorderStyle borderStyleDiagonalUpRightToBottomLeft)
        {
            if (borderStyleDiagonalUpLeftToBottomRight == null && borderStyleDiagonalUpRightToBottomLeft == null)
            {
                return;
            }

            if (borderStyleDiagonalUpLeftToBottomRight != null && borderStyleDiagonalUpLeftToBottomRight.Type != BorderType.None)
            {
                border.DiagonalDown = true;
            }

            if (borderStyleDiagonalUpRightToBottomLeft != null && borderStyleDiagonalUpRightToBottomLeft.Type != BorderType.None)
            {
                border.DiagonalUp = true;
            }
        }

        private void ApplyBorderStyle(SLBorderProperties border, BorderStyle borderStyle)
        {
            if (borderStyle == null)
            {
                return;
            }

            border.BorderStyle = this.GetBorderStyle(borderStyle.Type);

            if (borderStyle.Color.HasValue)
            {
                var color = borderStyle.Color.Value;
                border.Color = System.Drawing.Color.FromArgb(color.Alpha, color.Red, color.Green, color.Blue);
            }
        }

        private BorderStyleValues GetBorderStyle(BorderType type)
        {
            switch (type)
            {
                case BorderType.None:
                    return BorderStyleValues.None;
                case BorderType.Thin:
                    return BorderStyleValues.Thin;
                case BorderType.Medium:
                    return BorderStyleValues.Medium;
                case BorderType.Dashed:
                    return BorderStyleValues.Dashed;
                case BorderType.Dotted:
                    return BorderStyleValues.Dotted;
                case BorderType.Thick:
                    return BorderStyleValues.Thick;
                case BorderType.Double:
                    return BorderStyleValues.Double;
                case BorderType.Hair:
                    return BorderStyleValues.Hair;
                case BorderType.MediumDashed:
                    return BorderStyleValues.MediumDashed;
                case BorderType.DashDot:
                    return BorderStyleValues.DashDot;
                case BorderType.MediumDashDot:
                    return BorderStyleValues.MediumDashDot;
                case BorderType.DashDotDot:
                    return BorderStyleValues.DashDotDot;
                case BorderType.MediumDashDotDot:
                    return BorderStyleValues.MediumDashDotDot;
                case BorderType.SlantedDashDot:
                    return BorderStyleValues.SlantDashDot;
                default:
                    throw new ArgumentOutOfRangeException(nameof(type));
            }
        }

        private void ApplyVerticalStyle(SLStyle cellStyle, VerticalAlignment verticalAlignment)
        {
            switch (verticalAlignment)
            {
                case VerticalAlignment.Bottom:
                    cellStyle.SetVerticalAlignment(VerticalAlignmentValues.Bottom);
                    break;
                case VerticalAlignment.Top:
                    cellStyle.SetVerticalAlignment(VerticalAlignmentValues.Top);
                    break;
                case VerticalAlignment.Center:
                    cellStyle.SetVerticalAlignment(VerticalAlignmentValues.Center);
                    break;
                case VerticalAlignment.Justify:
                    cellStyle.SetVerticalAlignment(VerticalAlignmentValues.Justify);
                    break;
                case VerticalAlignment.Distributed:
                    cellStyle.SetVerticalAlignment(VerticalAlignmentValues.Distributed);
                    break;
                case VerticalAlignment.JustifyDistributed:
                    throw new NotImplementedException();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(verticalAlignment));
            }
        }

        private void ApplyHorizontalAlignment(SLStyle cellStyle, HorizontalAlignment horizotalAligment)
        {
            switch (horizotalAligment)
            {
                case HorizontalAlignment.General:
                    cellStyle.SetHorizontalAlignment(HorizontalAlignmentValues.General);
                    break;
                case HorizontalAlignment.Left:
                    cellStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Left);
                    break;
                case HorizontalAlignment.Center:
                    cellStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
                    break;
                case HorizontalAlignment.Right:
                    cellStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Right);
                    break;
                case HorizontalAlignment.Fill:
                    cellStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Fill);
                    break;
                case HorizontalAlignment.Justify:
                    cellStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Justify);
                    break;
                case HorizontalAlignment.CenterAcrossSelection:
                    cellStyle.SetHorizontalAlignment(HorizontalAlignmentValues.CenterContinuous);
                    break;
                case HorizontalAlignment.Distributed:
                    cellStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Distributed);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(horizotalAligment));
            }
        }

        private void RenderCell(SLDocument document, DocumentModel.Cell cell, int rowCounter, int cellCounter)
        {
            if (cell.Type == ReportModel.CellType.Formula)
            {
                document.SetCellValue(rowCounter, cellCounter, $"={cell.Value.ToString()}");
                return;
            }

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
