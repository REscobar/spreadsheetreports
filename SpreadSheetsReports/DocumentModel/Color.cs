namespace SpreadSheetsReports.DocumentModel
{
    /// <summary>
    /// Represents a color in the document model.
    /// </summary>
    public struct Color
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Color"/> struct.
        /// </summary>
        /// <param name="red">The red channel value.</param>
        /// <param name="green">The green channel value.</param>
        /// <param name="blue">The blue channel value.</param>
        public Color(byte red, byte green, byte blue)
        {
            this.Alpha = 0;
            this.Red = red;
            this.Green = green;
            this.Blue = blue;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Color"/> struct.
        /// </summary>
        /// <param name="alpha">The alpha channel value.</param>
        /// <param name="red">The red channel value.</param>
        /// <param name="green">The green channel value.</param>
        /// <param name="blue">The blue channel value.</param>
        public Color(byte alpha,byte red, byte green, byte blue)
        {
            this.Alpha = alpha;
            this.Red = red;
            this.Green = green;
            this.Blue = blue;
        }

        /// <summary>
        /// Gets the alpha channel of the color.
        /// </summary>
        public byte Alpha { get; private set; }

        /// <summary>
        /// Gets the red channel of the color.
        /// </summary>
        public byte Red { get; private set; }

        /// <summary>
        /// Gets the green channel of the color.
        /// </summary>
        public byte Green { get; private set; }

        /// <summary>
        /// Gets the blue channel of the color.
        /// </summary>
        public byte Blue { get; private set; }
    }
}
