namespace SpreadSheetsReports.WpfUi.Cells
{
    using System.Collections.Generic;

    public class FontSizesList
    {
        public int[] GetFontSizes()
        {
            var intList = new List<int>();
            for (int i = 8; i <= 12; i++)
            {
                intList.Add(i);
            }

            for (int i = 14; i <= 28; i++)
            {
                intList.Add(i);
            }

            intList.Add(36);
            intList.Add(48);
            intList.Add(72);

            return intList.ToArray();
        }
    }
}
