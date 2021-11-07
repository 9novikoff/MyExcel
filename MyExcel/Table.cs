using LabCalculator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;

namespace MyExcel
{
    class Table
    {
        public Dictionary<string, Cell> CellsDictionary = new Dictionary<string, Cell>();
        public DataGridView dgv { get; set; }
        private int _rows = 0;
        private int _columns = 0;

        public Table(int rows, int columns)
        {
            this._rows = rows;
            this._columns = columns;
            CreateTable();
        }
        public void InitializeDataGridView()
        {
            for (int i = 1; i <= _columns; i++)
            {
                DataGridViewColumn col = new DataGridViewColumn();
                col.HeaderText = GetColumnName(i);
                col.Name = col.HeaderText;
                Cell cell = new Cell();
                col.CellTemplate = cell;
                dgv.Columns.Add(col);
            }

            for (int i = 1; i <= _rows; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.HeaderCell.Value = i.ToString();
                dgv.Rows.Add(row);
            }
        }

        private void CreateTable()
        {
            for (int i = 1; i <= _rows; i++)
            {
                for (int j = 1; j <= _columns; j++)
                {
                    Cell currentCell = new Cell();
                    string cellName = GetColumnName(j) + i.ToString();
                    currentCell.Column = j;
                    currentCell.Row = i;
                    currentCell.Name = cellName;
                    CellsDictionary.Add(cellName, currentCell);
                }
            }
        }

        public string GetColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int temp = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + temp) + columnName;
                columnNumber = (columnNumber - temp) / 26;
            }

            return columnName;

        }

        public void SetCell(string cellName, string expression)
        {
            if (CheckingForRecursion(CellsDictionary[cellName], CellsDictionary[cellName]))
            {
                MessageBox.Show("Recursion is present! Some cells will be cleared!");
                RemoveDependentCells(cellName, cellName);
                return;
            }
            
            foreach(var name in CellsDictionary[cellName].thisDepend)
            {
                CellsDictionary[name].dependOnThis.Remove(cellName);
            }
            CellsDictionary[cellName].thisDepend.Clear();

            string refExpression = expression;

            Regex regex = new Regex(@"[A-Z]{1,2}[0-9]+");
            MatchCollection references = regex.Matches(expression);

            try
            {
                while (references.Count > 0)
                {
                    var reference = references[0].ToString();

                    CellsDictionary[cellName].thisDepend.Add(reference);
                    CellsDictionary[reference].dependOnThis.Add(cellName);

                    expression = expression.Replace(reference, CellsDictionary[reference].Value.ToString());
                    references = regex.Matches(expression);
                }

                double value = Calculator.Evaluate(expression);

                if (value == double.NegativeInfinity || value == double.PositiveInfinity)
                    throw new Exception();

                CellsDictionary[cellName].Expression = refExpression;
                CellsDictionary[cellName].Value = value;

                dgv[CellsDictionary[cellName].Column - 1, CellsDictionary[cellName].Row - 1].Value = value;

                RefreshCells(cellName);
            }
            catch (Exception)
            {
                MessageBox.Show("Invalid Expression. Cell will be cleared!");
                CellsDictionary[cellName].Value = 0;
                CellsDictionary[cellName].Expression = "0";
                CellsDictionary[cellName].thisDepend.Clear();
                dgv[CellsDictionary[cellName].Column - 1, CellsDictionary[cellName].Row - 1].Value = null;
            }
        }

        public void RefreshCells(string cellName)
        {
            foreach (var currentCellName in CellsDictionary[cellName].dependOnThis.ToList<string>())
            {
                var cell = CellsDictionary[currentCellName];
                SetCell(cell.Name, cell.Expression);
                RefreshCells(currentCellName);

            }
        }

        private bool CheckingForRecursion(Cell cell1, Cell cell2)
        {
            if (cell1.thisDepend.Contains(cell2.Name))
            {
                return true;
            }
            foreach (var cellName in cell1.thisDepend)
            {
                if (CheckingForRecursion(CellsDictionary[cellName], cell2))
                {
                    return true;
                }
            }
            return false;
        }

        private void RemoveDependentCells(string cellName, string cellName2)
        {
            List<string> res = new List<string>();
            res.Add(cellName);
            foreach (var name in CellsDictionary[cellName].dependOnThis)
            {
                if (name == cellName2)
                    continue;
                RemoveDependentCells(name, cellName2);
            }
            foreach(var name in res)
            {
                CellsDictionary[name].Value = 0;
                CellsDictionary[name].Expression = "0";
                CellsDictionary[name].dependOnThis.Clear();
                CellsDictionary[name].thisDepend.Clear();
                dgv[CellsDictionary[name].Column - 1, CellsDictionary[name].Row - 1].Value = null;
            }
        }

        public void AddRow()
        {
            try
            {
                DataGridViewRow row = new DataGridViewRow();
                row.HeaderCell.Value = (_rows + 1).ToString();
                dgv.Rows.Add(row);
                _rows++;

                for (int i = 1; i <= _columns; i++)
                {
                    string cellName = GetColumnName(i) + _rows.ToString();
                    Cell cell = new Cell();
                    cell.Name = cellName;
                    cell.Column = i;
                    cell.Row = _rows;
                    CellsDictionary.Add(cellName, cell);
                }
            }
            catch
            {
                MessageBox.Show("Unable to create more rows!");
                return;
            }
        }

        public void AddColumn()
        {
            try
            {
                DataGridViewColumn column = new DataGridViewColumn();
                column.Name = GetColumnName(_columns + 1);
                Cell templateCell = new Cell();
                column.CellTemplate = templateCell;
                dgv.Columns.Add(column);
                _columns++;
 
                for (int i = 1; i <= _rows; i++)
                {
                    string cellName = column.HeaderCell.Value + i.ToString();
                    Cell cell = new Cell();
                    cell.Name = cellName;
                    cell.Column = _columns;
                    cell.Row = i;
                    CellsDictionary.Add(cellName, cell);
                }
            }
            catch
            {
                MessageBox.Show("Unable to create more columns!");
                return;
            }
        }

        public void RemoveRow()
        {
            for (int i = 1; i <= _columns; i++)
            {
                string cellName = GetColumnName(i) + _rows.ToString();
                
                if(dgv[i-1, _rows - 1].Value != null || CellsDictionary[cellName].dependOnThis.Count() != 0)
                {
                    MessageBox.Show("Unable to remove last row!");
                    return;
                }

            }

            for (int i = 1; i <= _columns; i++)
            {
                string cellName = GetColumnName(i) + _rows.ToString();
                CellsDictionary.Remove(cellName);
            }

            dgv.Rows.RemoveAt(_rows - 1);
            _rows--;
        }

        public void RemoveColumn()
        {
            string columnName = GetColumnName(_columns);

            for (int i = 1; i <= _rows; i++)
            {
                string cellName = columnName + i.ToString();

                if (dgv[_columns - 1, i - 1].Value != null || CellsDictionary[cellName].dependOnThis.Count() != 0)
                {
                    MessageBox.Show("Unable to remove last column!");
                    return;
                }

            }

            for (int i = 1; i <= _rows; i++)
            {
                string cellName = columnName + i.ToString();
                CellsDictionary.Remove(cellName);
            }

            dgv.Columns.RemoveAt(_columns - 1);
            _columns--;
        }

        private void ClearTable()
        {
            CellsDictionary.Clear();
            dgv.Rows.Clear();
            dgv.Columns.Clear();

        }

        private void SetFullTable()
        {
            for(int i = 1; i <= _columns; i++)
            {
                for(int j = 1; j <= _rows; j++)
                {
                    string cellName = GetColumnName(i) + j.ToString();

                    var value = CellsDictionary[cellName].Value;

                    if (value == 0 && CellsDictionary[cellName].Expression == "0")
                        dgv[i - 1, j - 1].Value = null;
                    else
                        dgv[i - 1, j - 1].Value = value;
                }
            }
        }

        public void SaveTable(string filePath)
        {
            using(StreamWriter sw = new StreamWriter(filePath))
            {
                sw.WriteLine(_rows);
                sw.WriteLine(_columns);

                foreach(var pair in CellsDictionary)
                {
                    sw.WriteLine(pair.Key);
                    sw.WriteLine(pair.Value.Expression);
                    sw.WriteLine(pair.Value.Value);
                    sw.WriteLine(pair.Value.Row);
                    sw.WriteLine(pair.Value.Column);
                    
                    foreach(var item in pair.Value.dependOnThis)
                    {
                        sw.Write(item + " ");
                    }

                    sw.WriteLine();

                    foreach (var item in pair.Value.thisDepend)
                    {
                        sw.Write(item + " ");
                    }
                    
                    sw.WriteLine();

                }
            }
        }

        public void OpenTable(string filePath)
        {
            ClearTable();
            using (StreamReader sr = new StreamReader(filePath))
            {
                _rows = Convert.ToInt32(sr.ReadLine());
                _columns = Convert.ToInt32(sr.ReadLine());

                for (int i = 1; i <= _rows; i++)
                {
                    for (int j = 1; j <= _columns; j++)
                    {
                        string cellName = sr.ReadLine();
                        string expression = sr.ReadLine();
                        double value = Convert.ToDouble(sr.ReadLine());
                        int row = Convert.ToInt32(sr.ReadLine());
                        int column = Convert.ToInt32(sr.ReadLine());
                        List<string> dOT = sr.ReadLine().Split(" ").ToList<string>();
                        List<string> tD = sr.ReadLine().Split(" ").ToList<string>();

                        var itemToRemove = dOT.SingleOrDefault(r => r == "");
                        if (itemToRemove != null)
                            dOT.Remove(itemToRemove);

                        itemToRemove = tD.SingleOrDefault(r => r == "");
                        if (itemToRemove != null)
                            tD.Remove(itemToRemove);

                        Cell cell = new Cell();

                        cell.Name = cellName;
                        cell.Expression = expression;
                        cell.Value = value;
                        cell.Column = column;
                        cell.Row = row;
                        cell.dependOnThis = dOT;
                        cell.thisDepend = tD;

                        CellsDictionary[cellName] = cell;

                    }
                }

                InitializeDataGridView();
                SetFullTable();
                

            }
        }


    }
}
