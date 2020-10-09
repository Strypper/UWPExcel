using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TTExcelGenerator
{
    public class Row
    {
        public int NumRow;
        public List<Cell> Cells;

        private Row() { }

        public Row(int numRow, List<Cell> cells)
        {
            NumRow = numRow;
            Cells = cells;
        }
    }
}
