using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeExcelTab
{
    class CellEntity
    {
        Style style;

        String value;

        Cell cell;

        public Style Style
        {
            get
            {
                return style;
            }

            set
            {
                style = value;
            }
        }

        public string Value
        {
            get
            {
                return value;
            }

            set
            {
                this.value = value;
            }
        }

        public Cell Cell
        {
            get
            {
                return cell;
            }

            set
            {
                cell = value;
            }
        }
    }
}
