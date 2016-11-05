//Name: Xueliang Sun ID: 11387859
using CptS322;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Spreadsheet_XSun
{
    public partial class Form1 : Form
    {
        Spreadsheet sheet;
        
        public Form1()
        {
            InitializeComponent();
        }
        
        // set the Form in OnLoad
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            dataGridView1.Columns.Clear();

            // add 26 columns
            for (int i = 0; i < 26; i++)
            {
                DataGridViewColumn col = new DataGridViewColumn();
                char a = (char)('A' + i);
                col.HeaderCell.Value = a.ToString();
                col.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Add(col);
            }

            // add 50 rows
            for (int i = 0; i < 50; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.HeaderCell.Value = (i+1).ToString();
                try
                {
                    dataGridView1.Rows.Add(row);
                }
                catch (InvalidCastException ea) when (ea != null)
                {
                    throw new System.InvalidOperationException("Can't add a row!");
                }
            }

            // initialize the spreadsheet class and add event handler
            int i_row = 50, i_col = 26;
            sheet = new Spreadsheet(i_row, i_col);
            sheet.CellPropertyChanged += new PropertyChangedEventHandler(SheetChange);
        }

        // update gridview's cell value
        public void SheetChange(object sender, PropertyChangedEventArgs e)
        {
            Cell cell = sender as Cell;
            if (e.PropertyName == "CellBackColorChange")
            {
                Color tmp = (cell.BackColor == 0) ? Color.Empty : Color.FromArgb((int)cell.BackColor);
                dataGridView1.Rows[cell.RowIndex].Cells[cell.ColIndex].Style.BackColor = tmp;
            }
            else
            {
                dataGridView1.Rows[cell.RowIndex].Cells[cell.ColIndex].Value = (cell.Value == null) ? string.Empty : cell.Value;
            }
        }

        // set the onclick event
        private void button1_Click(object sender, EventArgs e)
        {
            Random rnd = new Random();

            // get 50 hello world
            for(int i = 0; i < 50; i++)
            {
                int i_row = rnd.Next(sheet.RowCount);
                int i_col = rnd.Next(sheet.ColumnCount);
                sheet.cell_arr[i_row][i_col].Text = "Hello World";
            }

            // change column B
            for(int i = 0; i < 50; i++)
            {
                int j = i + 1;
                sheet.cell_arr[i][1].Text = "This is cell B" + j;
            }

            // change column A
            for (int i = 0; i < 50; i++)
            {
                int j = i + 1;
                sheet.cell_arr[i][0].Text = "=B" + j;
            }
        }
        // we change the gridcell value to cell's original text while editing
        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            DataGridViewCell tmp = (sender as DataGridView).CurrentCell;
            tmp.Value = sheet.cell_arr[tmp.RowIndex][tmp.ColumnIndex].Text;
            Cell cell_tmp = sheet.cell_arr[tmp.RowIndex][tmp.ColumnIndex];
            // get the RestoreCell command and build the UndoRedoCollection, then send the commands to undo stack
            IUndoRedoCmd cmd_tmp = new RestoreCell(cell_tmp, cell_tmp.Text);
            UndoRedoCollection coll_tmp = new UndoRedoCollection(cmd_tmp);
            coll_tmp.SetPropertyChange("Text Change");
            sheet.AddUndo(coll_tmp);
        }

        // we change the gridcell value to cell's value after editing
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell tmp = (sender as DataGridView).CurrentCell;
            //sheet.isupdate = false; // after editing, the spreadsheet should modify that cell and should not be in status of updating other refering cells
            if (tmp != null && tmp.Value != null)
            {
                if (sheet.cell_arr[tmp.RowIndex][tmp.ColumnIndex].Text == tmp.Value.ToString())
                    tmp.Value = sheet.cell_arr[tmp.RowIndex][tmp.ColumnIndex].Value;
                else
                    sheet.cell_arr[tmp.RowIndex][tmp.ColumnIndex].Text = tmp.Value.ToString();
            }
        }
        
        // when click background color, we change selected cells' colors
        private void chooseBackgroudColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //get selected cells
            DataGridViewSelectedCellCollection cell_coll = dataGridView1.SelectedCells;
            if (cell_coll.Count == 0) return;

            //use tag to check whether this colordialog exists or not
            ColorDialog cd = chooseBackgroudColorToolStripMenuItem.Tag as ColorDialog;
            if(cd == null)
            {
                // if not exist, we need to instantiate a new color dialog
                cd = new ColorDialog();
                cd.Color = cell_coll[0].Style.BackColor;
                chooseBackgroudColorToolStripMenuItem.Tag = cd;
            }
            // then we traverse the cells and set background color
            List<IUndoRedoCmd> list = new List<IUndoRedoCmd>();
            if (cd.ShowDialog() == DialogResult.OK)
            {
                // use list to add single command
                for (int i = 0; i < cell_coll.Count; i++)
                {
                    Cell tmp = sheet.cell_arr[cell_coll[i].RowIndex][cell_coll[i].ColumnIndex];
                    list.Add(new RestoreCell(tmp, tmp.BackColor));
                    sheet.cell_arr[cell_coll[i].RowIndex][cell_coll[i].ColumnIndex].BackColor = (uint) cd.Color.ToArgb();
                }
                // transform list to the undoredocollection and add to the undo stack of spreadsheet
                UndoRedoCollection coll_tmp = new UndoRedoCollection(list);
                coll_tmp.SetPropertyChange("BackGround Color Change");
                sheet.AddUndo(coll_tmp);
            } 
        }
        
        // when click undo, implement undo
        private void UndoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sheet.Undo();
        }

        // when click redo, implement redo
        private void RedoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sheet.Redo();
        }

        // when drop down the toolstrip, we need to set the texts of two dropdown iterms
        private void EditToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
        {
            // if null, set directly to Undo
            if(sheet.UndoTopMsg() == null)
            {
                UndoColorToolStripMenuItem.Text = "Undo";
                UndoColorToolStripMenuItem.ForeColor = Color.LightGray;
            }
            else// else, set to the top command's text
            {
                UndoColorToolStripMenuItem.Text = "Undo changing Cell " + sheet.UndoTopMsg();
                UndoColorToolStripMenuItem.ForeColor = Color.Black;
            }
            // if null, set directly to Redo
            if (sheet.RedoTopMsg() == null)
            {
                RedoToolStripMenuItem.Text = "Redo";
                RedoToolStripMenuItem.ForeColor = Color.LightGray;
            }
            else// else, set to the top command's text
            {
                RedoToolStripMenuItem.Text = "Redo changing Cell " + sheet.RedoTopMsg();
                RedoToolStripMenuItem.ForeColor = Color.Black;
            }
        }

        // click to load the spreadsheet
        private void LoadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // use filter to restrict to .xml files
            openFileDialog1.Filter = "XML files (*.xml)|*.xml";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sheet.Load(openFileDialog1.FileName); // call load function inside sheet
            }
        }

        // click to save the spreadsheet
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // use filter to restrict to .xml files
            saveFileDialog1.Filter = "XML files (*.xml)|*.xml";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sheet.Save(saveFileDialog1.FileName); // call save function inside sheet
            }
        }
    }
}
