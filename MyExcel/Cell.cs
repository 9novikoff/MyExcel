using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LabCalculator;

namespace MyExcel
{
    public class Cell : DataGridViewTextBoxCell
    {
        public new double Value { get; set; } = 0;
        public string Name { get; set; } = "";
        public string Expression { get; set; } = "0";
        public int Row { get; set; } = 0;
        public int Column { get; set; } = 0;

        public List<string> thisDepend = new List<string>();
        public List<string> dependOnThis = new List<string>();

    }
}
