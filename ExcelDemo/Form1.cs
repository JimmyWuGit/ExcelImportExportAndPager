using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace ExcelDemo
{
    public partial class Form1 : Form
    {
        DataTable dtTotal;
        public Form1()
        {
            InitializeComponent();
            dtTotal = new DataTable();
            //将分页控件初始化
            paginationControl1.SetPage(1);
        }

        /// <summary>
        /// 导入Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false; //不能多选
            ofd.Title = "请选择文件";
            ofd.Filter = "所有文件(*.*)|*.*|Excel 文件(*.xls)|*.xls|Excel 文件(*.xlsx)|*.xlsx";
            string filePath = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;
            }
            else
            {
                return;
            }
             
            if (!File.Exists(filePath))
            {
                MessageBox.Show("导入的文件不存在！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                dtTotal = ExcelHelper.ImportToTable(filePath, 6);
                if (dtTotal != null)
                {
                    //this.dgvData.DataSource = dtTotal;
                    //调用委托，进行分页操作
                    paginationControl1.Method = LoadData;
                    LoadData();

                    MessageBox.Show("导入成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("导入失败！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show("请先关闭要导入的Excel文件：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //throw;
                return;
            }
            
        }

        /// <summary>
        /// 根据委托定义一个方法，实现分页功能
        /// </summary>
        /// <param name="currentPage"></param>
        private void LoadData(int currentPage =1)
        {
            paginationControl1.Total = dtTotal.Rows.Count;
            int firstRow = paginationControl1.FirstRow;
            int lastRow = paginationControl1.LastRow;

            //从总的数据集dtTotal中，获取分页数据
            DataTable dt = new DataTable();
            //加载列的名称
            for(int i = 0; i < dtTotal.Columns.Count; i++)
            {
                dt.Columns.Add(dtTotal.Columns[i].ColumnName);
            }
            for(int i = firstRow; i < lastRow; i++)
            {
                dt.ImportRow(dtTotal.Rows[i]);
            }
            this.dgvData.DataSource = dt;

            paginationControl1.SetPage(currentPage);
        }

        /// <summary>
        /// 导出到Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgvData.DataSource;
            if (dt == null || dt.Rows.Count <= 0)
            {
                MessageBox.Show("没有数据导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }                

            string filePath = DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss");
            //打开文件对话框
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "保存文件";
            saveFileDialog.Filter = "Excel 文件(*.xls)|*.xls|Excel 文件(*.xlsx)|*.xlsx|所有文件(*.*)|*.*";
            saveFileDialog.FileName = filePath + "上料记录.xls";
            if(saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveFileDialog.FileName;
                ExcelHelper.ExportToExcel(dt, fileName);

                ExcelHelper.ExportSuccessTips(fileName);
            }            
        }

        /// <summary>
        /// 清空数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClear_Click(object sender, EventArgs e)
        {
            if (dtTotal != null && dtTotal.Rows.Count > 0)
            {
                if (MessageBox.Show("确定清空数据吗？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    dtTotal.Clear(); //清空数据
                    dtTotal.Columns.Clear(); //列名也清空掉
                    dgvData.DataSource = dtTotal;

                    //将分页控件初始化
                    paginationControl1.Total = dtTotal.Rows.Count;
                    paginationControl1.CurrentPage = 1;
                    paginationControl1.SetPage(1);
                }
            }
            else
            {
                MessageBox.Show("暂无数据！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }                        
        }
    }
}
