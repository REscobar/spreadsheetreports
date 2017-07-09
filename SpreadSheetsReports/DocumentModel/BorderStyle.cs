// Copyright (c). All rights reserved.

namespace SpreadSheetsReports.DocumentModel
{
    /// <summary>
    /// Represents a border style
    /// </summary>
    public class BorderStyle
    {
        public BorderStyle()
        {
            this.Color = DocumentModel.Color.BlackColor;
            this.Type = BorderType.None;
        }

        /// <summary>
        /// Gets or sets the <see cref="BorderType"/>
        /// </summary>
        public BorderType Type { get; set; }

        /// <summary>
        /// Gets or sets the <see cref="DocumentModel.Color"/>
        /// </summary>
        public Color? Color { get; set; }
    }
}