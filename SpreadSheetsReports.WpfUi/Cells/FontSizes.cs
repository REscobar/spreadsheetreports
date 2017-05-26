using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetsReports.WpfUi.Cells
{
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
