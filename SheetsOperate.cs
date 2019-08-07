using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExcelBaseOnEPPlus
{
    class SheetsOperate
    {
        private TabControl tabControl;
        public int sheetNameNumber = 0;
        public ContextMenuStrip rightClickMenu;
        ExcelPackage ep;

        public SheetsOperate()
        {
        }

        public SheetsOperate(TabControl tc)
        {
            tabControl = tc;

        }

        /// <summary>
        /// 新建一个工作表
        /// </summary>
        /// <returns></returns>
        public TabControl NewSheet()
        {
            return NewSheet(tabControl);
        }

        /// <summary>
        /// 新建一个工作表
        /// </summary>
        /// <param name="tc">TabControl to be add</param>
        /// <returns></returns>
        public TabControl NewSheet(TabControl tc)
        {
            //更新工作表名字sheet后缀数字
            sheetNameNumber++;

            ////获取tc的工作表数量
            //var tpCount = tc.TabPages.Count;

            //构建工作表tp
            var tp = new TabPage("Sheet" + sheetNameNumber);

            //构建新表格dg
            DataGridView dg = new DataGridView();
            dg.AllowUserToAddRows = false;
            dg.ContextMenuStrip = this.rightClickMenu;
            dg.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dg.Dock = DockStyle.Fill;
            dg.RowHeadersWidth = 40;
            dg.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            dg.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dg_MouseDown);
            //构建表格列名称ABCD...
            for (int i = 0; i < 27; i++)
            {
                string colmunName = GenerateAlphabetColmunName(i);
                //string colmunName = Convert.ToString((char)(i + 65));
                dg.Columns.Add(colmunName, colmunName);
            }
            //dg.Location = new System.Drawing.Point(3, 3);
            dg.Name = "dataGridView" + sheetNameNumber;
            dg.RowTemplate.Height = 23;
            //dg.Size = new System.Drawing.Size(762, 378);
            dg.TabIndex = 0;

            //将dg添加进tp
            tp.Controls.Add(dg);

            //将tp添加进tc
            tc.TabPages.Add(tp);

            return tc;
        }

        private void dg_MouseDown(object sender, MouseEventArgs e)
        {
            var dg = sender as DataGridView;
            if (e.Button == MouseButtons.Right)
            {
                var gridview = GetSelectedGrid();
                gridview.ContextMenuStrip.Show(Control.MousePosition);
            }
        }

        public void SetSelectedRowsColor(Color color)
        {
            var dg = GetSelectedGrid();

            foreach (DataGridViewTextBoxCell cell in dg.SelectedCells)
            {
                cell.Style.BackColor = color;
            }
        }

        private string GenerateAlphabetColmunName(int bi)
        {
            string columnName = "";
            int alphaNumber = (bi + 26) / 26;
            int alpha = (bi + 26) % 26;
            for (int i = 0; i < alphaNumber; i++)
            {
                columnName += Convert.ToString((char)(alpha + 65));
            }

            return columnName;
        }


        /// <summary>
        /// 新建一个工作表
        /// </summary>
        /// <param name="tp"></param>
        /// <returns></returns>
        public TabControl NewSheet(TabPage tp)
        {
            tabControl.TabPages.Add(tp);
            return tabControl;
        }

        public TabControl NewSheet(DataGridView dg)
        {
            //更新工作表名字sheet后缀数字
            sheetNameNumber++;

            ////获取tc的工作表数量
            //var tpCount = tc.TabPages.Count;

            //构建工作表tp
            var tp = new TabPage(dg.Name);

            //将dg添加进tp
            tp.Controls.Add(dg);

            //将tp添加进tc
            tabControl.TabPages.Add(tp);

            return tabControl;

        }

        public TabControl NewSheet(TabControl tc, DataGridView dg)
        {
            //更新工作表名字sheet后缀数字
            sheetNameNumber++;

            ////获取tc的工作表数量
            //var tpCount = tc.TabPages.Count;

            //构建工作表tp
            var tp = new TabPage("Sheet" + sheetNameNumber);

            //将dg添加进tp
            tp.Controls.Add(dg);

            //将tp添加进tc
            tc.TabPages.Add(tp);

            return tc;
        }


        /// <summary>
        /// 删除选中工作表
        /// </summary>
        public void RemoveSelectedSheet()
        {
            tabControl.TabPages.RemoveAt(tabControl.SelectedIndex);
        }


        /// <summary>
        /// 删除选中工作表
        /// </summary>
        /// <param name="tc"></param>
        public void RemoveSelectedSheet(TabControl tc)
        {
            tc.TabPages.RemoveAt(tc.SelectedIndex);
        }


        /// <summary>
        /// 删除指定索引的工作表
        /// </summary>
        /// <param name="index"></param>
        public void RemoveSheetAt(int index)
        {
            tabControl.TabPages.RemoveAt(index);
        }


        /// <summary>
        /// 删除指定索引的工作表
        /// </summary>
        /// <param name="tc"></param>
        /// <param name="index"></param>
        public void RemoveSheetAt(TabControl tc, int index)
        {
            tc.TabPages.RemoveAt(index);
        }

        /// <summary>
        /// 重命名选中的工作表
        /// </summary>
        /// <param name="name"></param>
        public void RenameSelectedSheet(string name)
        {
            TabPage selectedTabPage = tabControl.SelectedTab;
            selectedTabPage.Text = name;
            selectedTabPage.Name = name;
        }

        /// <summary>
        /// 重命名选中的工作表
        /// </summary>
        /// <param name="name"></param>
        public void RenameSelectedSheet(TabControl tc, string name)
        {
            TabPage selectedTabPage = tc.SelectedTab;
            tc.Text = name;
            tc.Name = name;
        }

        /// <summary>
        /// 重命名指定索引的工作表
        /// </summary>
        /// <param name="index"></param>
        /// <param name="name"></param>
        public void RenameSheetAt(int index, string name)
        {
            tabControl.TabPages[index].Text = name;
            tabControl.TabPages[index].Name = name;
        }

        /// <summary>
        /// 重命名指定索引的工作表
        /// </summary>
        /// <param name="index"></param>
        /// <param name="name"></param>
        public void RenameSheetAt(TabControl tc, int index, string name)
        {
            tc.TabPages[index].Text = name;
            tc.TabPages[index].Name = name;
        }

        public void NewWorkBook()
        {
            tabControl.TabPages.Clear();
            NewSheet();
        }

        public void NewWorkBook(FileInfo fi)
        {
            using (var ep = new ExcelPackage(fi))
            {
                InitPackageToUI(ep);
            }
        }

        private void InitPackageToUI(ExcelPackage ep)
        {
            //先清空工作表
            ClearSheets();

            //将ep中所有工作表添加到tabControl
            foreach (var ws in ep.Workbook.Worksheets)
            {
                //获取工作表最右下角元素的列索引
                int colmunEnd = ws.Dimension.End.Column;
                //获取工作表最右下角元素的行索引
                int rowEnd = ws.Dimension.End.Row;

                //构建新表格dg
                DataGridView dg = new DataGridView();
                dg.AllowUserToAddRows = false;
                dg.ContextMenuStrip = this.rightClickMenu;
                dg.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                dg.Dock = DockStyle.Fill;
                dg.RowHeadersWidth = 40;
                dg.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
                dg.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dg_MouseDown);
                //构建表格列名称，数量等同于ws工作表
                for (int i = 0; i < colmunEnd; i++)
                {
                    string colmunName = GenerateAlphabetColmunName(i);
                    //string colmunName = Convert.ToString((char)(i + 65));
                    dg.Columns.Add(colmunName, colmunName);
                }
                //dg.Location = new System.Drawing.Point(3, 3);
                dg.Name = ws.Name;
                dg.RowTemplate.Height = 23;
                //dg.Size = new System.Drawing.Size(762, 378);
                dg.TabIndex = 0;

                //遍历ws工作表的行
                for (int row = 0; row < rowEnd; row++)
                {
                    //生成行的头名
                    var gridRow = new DataGridViewRow {HeaderCell = { Value = (row + 1).ToString()}};
                    
                    //遍历ws工作表的列,生成一条gridrow
                    for (var col = 0; col < colmunEnd; col++)
                    {
                        
                        var cell = ws.Cells[row + 1, col + 1];
                        using (var uiCell = new DataGridViewTextBoxCell())
                        {
                            uiCell.Value = cell.Value;
                            gridRow.Cells.Add(uiCell);
                        }
                        //NewSheet();
                    }

                    //将一条gridrow添加进dg
                    dg.Rows.Add(gridRow);

                    //tabControl.TabPages.Add(sheet);
                }

                //生成一个基于dg的工作表
                NewSheet(dg);
            }
            
        }

        private void ClearSheets()
        {
            tabControl.TabPages.Clear();
        }

        public void AddColmunToSelectedSheet()
        {
            var gridview = GetSelectedGrid();
            int alphabet = gridview.ColumnCount;
            string colmunName = GenerateAlphabetColmunName(alphabet);
            gridview.Columns.Add(colmunName, colmunName);
        }

        public void AddRowToSelectedSheet()
        {
            var gridview = GetSelectedGrid();
            var rowCount = gridview.Rows.Count;
            var gridRow = new DataGridViewRow { HeaderCell = { Value = (rowCount + 1).ToString() } };
            gridview.Rows.Add(gridRow);
        }

        private DataGridView GetSelectedGrid()
        {
            var page = tabControl.SelectedTab;
            DataGridView dg = new DataGridView();
            //获取到tabpage里的gridview
            foreach (var cc in page.Controls)
            {
                dg = cc as DataGridView;
            }

            //var dataGrid = tabControl.SelectedTab.Controls[0] as DataGridView;
            return dg;
        }

        private DataGridView GetGridFromTabPage(TabPage tp)
        {
            DataGridView dg = new DataGridView();
            //获取到tabpage里的gridview
            foreach (var cc in tp.Controls)
            {
                dg = cc as DataGridView;
            }

            //var dataGrid = tabControl.SelectedTab.Controls[0] as DataGridView;
            return dg;
        }

        public void SaveAs(FileInfo fi)
        {
            ep = new ExcelPackage(new MemoryStream());
            foreach (TabPage tp in tabControl.TabPages)
            {
                var ws = ep.Workbook.Worksheets.Add(tp.Text);
                var gridview = GetGridFromTabPage(tp);
                for (int row = 0; row < gridview.Rows.Count; row++)
                {
                    for (int col = 0; col < gridview.Columns.Count; col++)
                    {
                        ws.Cells[row + 1, col + 1].Value = gridview.Rows[row].Cells[col].Value;
                        //ws.Cells[row + 1, col + 1, row + 1, col + 1].Style.Fill.PatternColor.SetColor(Color.White);
                        if (gridview.Rows[row].Cells[col].Style.BackColor.ToArgb() != Color.FromArgb(0, 0, 0, 0).ToArgb())
                        {
                            ws.Cells[row + 1, col + 1, row + 1, col + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells[row + 1, col + 1, row + 1, col + 1].Style.Fill.BackgroundColor.SetColor(gridview.Rows[row].Cells[col].Style.BackColor);
                        }
                    }
                }

            }
            try
            {
                ep.SaveAs(fi);
                ep.Dispose();
            }
            catch (InvalidOperationException ioe)
            {
                MessageBox.Show(ioe.Message);
                ep.Dispose();
            }
        }
    }
}
