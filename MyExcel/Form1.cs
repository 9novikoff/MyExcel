using LabCalculator;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyExcel
{
    public partial class MyExcelForm : Form
    { 
        const int Rows = 20;
        const int Columns = 20; // 655 max
        private Table _table = new Table(Rows, Columns); 

        public MyExcelForm()
        {
            InitializeComponent();
            _table.dgv = dgv;
            _table.InitializeDataGridView();

            this.ActiveControl = dgv;
        }
        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void aBoutToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void MyExcelForm_Load(object sender, EventArgs e)
        {

        }

        private void dgv_ChooseCell(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (e.RowIndex >= 0)
            {
                string cellName = senderGrid.Columns[e.ColumnIndex].Name + senderGrid.Rows[e.RowIndex].HeaderCell.Value;

                senderGrid.CellValueChanged -= dgv_CellValueChanged;

                senderGrid.CurrentCell.Value = _table.CellsDictionary[cellName].Expression;

                senderGrid.CellValueChanged += dgv_CellValueChanged;

                toolStripTextBox1.Text = cellName;
            }
        }

        private void dgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (e.RowIndex >= 0)
            {
                int column = e.ColumnIndex;
                int row = e.RowIndex;

                if(senderGrid[column, row].Value == null)
                {
                    return;
                }

                string expression = senderGrid[column, row].Value.ToString();

                string cellName = _table.GetColumnName(column + 1) + (row + 1).ToString();

                senderGrid.CellValueChanged -= dgv_CellValueChanged;

                _table.SetCell(cellName, expression);

                senderGrid.CellValueChanged += dgv_CellValueChanged;
            }
        }

        private void dgv_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            senderGrid.CellValueChanged -= dgv_CellValueChanged;

            Cell cell = _table.CellsDictionary[_table.GetColumnName(e.ColumnIndex + 1) + (e.RowIndex + 1).ToString()];
            double value = cell.Value;

            if(value == 0 && cell.Expression == "0" )
            {
                senderGrid[e.ColumnIndex, e.RowIndex].Value = null;
            }
            else
            {
                senderGrid[e.ColumnIndex, e.RowIndex].Value = value;
            }

            senderGrid.CellValueChanged += dgv_CellValueChanged;
        }

        private void AddRowButton_Click(object sender, EventArgs e)
        {
            _table.AddRow();
        }

        private void AddColumnButton_Click(object sender, EventArgs e)
        {
            _table.AddColumn();
        }

        private void RemoveRowButton_Click(object sender, EventArgs e)
        {
            _table.RemoveRow();
        }

        private void RemoveColumnButton_Click(object sender, EventArgs e)
        {
            _table.RemoveColumn();
        }

        private void SaveMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog fileSave = new SaveFileDialog();

            if (fileSave.ShowDialog() == DialogResult.OK)
            {
                string filePath = fileSave.FileName;

                _table.SaveTable(filePath);
            }
        }

        private void OpenMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileOpen = new OpenFileDialog();

            if (fileOpen.ShowDialog() == DialogResult.OK)
            {
                string filePath = fileOpen.FileName;


                dgv.CellValueChanged -= dgv_CellValueChanged;
                dgv.CellEnter -= dgv_ChooseCell;
                dgv.CellLeave -= dgv_CellLeave;


                _table.OpenTable(filePath);


                dgv.CellValueChanged += dgv_CellValueChanged;
                dgv.CellLeave += dgv_CellLeave;
                dgv.CellEnter += dgv_ChooseCell;



            }
        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
