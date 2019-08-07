using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelBaseOnEPPlus
{
    public partial class Form1 : Form
    {
        private SheetsOperate so;
        public Form1()
        {
            InitializeComponent();
            so = new SheetsOperate(tabControl1);
            so.rightClickMenu = this.contextMenuStrip2;
            
        }


        private void TabControl1_MouseDown(object sender, MouseEventArgs e)
        {
            //右键弹出菜单
            if (e.Button == MouseButtons.Right)
            {
                contextMenuStrip1.Show(Control.MousePosition);

            }
        }

        private void 删除工作表ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.RemoveSelectedSheet();
        }

        private void 重命名ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InputTextForm inputForm = new InputTextForm("重命名工作表名称", tabControl1.SelectedTab.Text);
            inputForm.ShowDialog();
            if (inputForm.DialogResult == DialogResult.OK)
            {
                string inputText = inputForm.GetInputText();
                so.RenameSelectedSheet(inputText);
            }
            //so.RenameSelectedSheet();
        }

        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel表格文件 (*.xlsx)|*.xlsx";
            var dialogResult = openFileDialog1.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                so.NewWorkBook(new System.IO.FileInfo(openFileDialog1.FileName));
            }
            
        }

        private void 新建ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.NewWorkBook();
        }


        private void DataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            Console.WriteLine("DragEnter");
        }

        private void DataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            Console.WriteLine("DragDrop");
        }

        private void DataGridView1_DragLeave(object sender, EventArgs e)
        {
            Console.WriteLine("DragLeave");
        }

        private void 新建工作表ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.NewSheet();
        }

        private void 添加列ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.AddColmunToSelectedSheet();
        }

        private void 添加行ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.AddRowToSelectedSheet();
        }

        private void 另存为ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";

            var dialogResult = saveFileDialog1.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                so.SaveAs(new System.IO.FileInfo(saveFileDialog1.FileName));
            }
        }

        private void DataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {

                foreach (DataGridViewTextBoxCell cell in dataGridView1.SelectedCells)
                {
                    cell.Style.BackColor = Color.Blue;
                }
            }
        }

        private void 红色ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.SetSelectedRowsColor(Color.Red);
        }

        private void 黄色ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.SetSelectedRowsColor(Color.Yellow);
        }

        private void 蓝色ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.SetSelectedRowsColor(Color.Blue);
        }

        private void 绿色ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.SetSelectedRowsColor(Color.Green);
        }

        private void 青色ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.SetSelectedRowsColor(Color.Cyan);
        }

        private void 粉色ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.SetSelectedRowsColor(Color.Pink);
        }

        private void 紫色ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            so.SetSelectedRowsColor(Color.Purple);
        }
    }
}
