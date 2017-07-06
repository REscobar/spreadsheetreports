namespace SpreadSheetsReports.DocumentModel
{
    public class FontStyle
    {
        public FontStyle()
        {
            this.Color = DocumentModel.Color.BlackColor;
        }

        public string FontName { get; set; }

        public bool IsItalic { get; set; }

        public bool IsBold { get; set; }

        public bool IsStrikeout { get; set; }

        public decimal Size { get; set; }

        public UnderLineStyle Underline { get; set; }

        public FontScriptStyle ScriptStyle { get; set; }

        public Color? Color { get; set; }
    }
}
