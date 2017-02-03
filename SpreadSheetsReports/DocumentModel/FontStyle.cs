namespace SpreadSheetsReports.DocumentModel
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    public class FontStyle
    {
        public string FontNmae { get; set; }

        public bool IsItalic { get; set; }

        public bool IsBold { get; set; }

        public decimal Size { get; set; }

        public UnderLineStyle Underline { get; set; }

        public FontScriptStyle ScriptStyle { get; set; }

        public Color? Color { get; set; }
    }
}
