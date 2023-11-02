using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace ExcelDemo
{
    public partial class PaginationControl : UserControl
    {
        //定义一个委托：根据当前页码加载数据
        public delegate void LoadDataDelegate(int currentPage);  

        public PaginationControl()
        {
            InitializeComponent();
        }

        private bool _isChangeByUser = true;  //是否由用户改变每页记录数或页码，注意何时使用该变量，以及何时赋值
        private int _currentPage = 1; //当前页码
        /// <summary>
        /// 当前页
        /// </summary>
        public int CurrentPage
        {
            get { return _currentPage; }
            set { _currentPage = value; }
        }
        private int _totalPage = 0;   //总页数
        private int _total = 0; //总记录数
        [Description("总记录数"), Category("自定义")]
        public int Total
        {
            get { return _total; }
            set { _total = value; }
        }
        
        private int _pageSize = 10; //每页显示数量，默认10条
        /// <summary>
        /// 每页记录数
        /// </summary>
        [Description("每页记录数"),Category("自定义"),DefaultValue(10)]
        public int PageSize
        {
            get { return _pageSize; }
            set 
            { 
                _pageSize = value;
                _isChangeByUser = false;
                cboPageSize.SelectedText = _pageSize.ToString();
                _isChangeByUser= true;
            }
        }

        private event LoadDataDelegate _method;
        [Description("加载数据的方法（事件）名称"),Category("自定义")]
        public LoadDataDelegate Method
        {
            set { _method = value; }
        }

        private int _firstRow;
        /// <summary>
        /// 每页首行索引
        /// </summary>
        public int FirstRow
        {
            get
            {
                if (_currentPage == 1)
                    _firstRow = 0;
                else
                    _firstRow = (_currentPage - 1) * _pageSize;
                return _firstRow;
            }
        }

        private int _lastRow;
        /// <summary>
        /// 每页行数最大的索引值
        /// </summary>
        public int LastRow
        {
            get
            {
                if (_currentPage == _totalPage)
                    _lastRow = _total;
                else if(_pageSize >= _total)
                    _lastRow = _total;
                else
                    _lastRow = _pageSize * _currentPage;
                return _lastRow;
            }
        }
                
        /// <summary>
        /// 设置分页组件内部各控件状态
        /// </summary>
        /// <param name="currentPage">当前页</param>
        /// <param name="total">总记录数</param>
        public void SetPage(int currentPage) //,int total)
        {
            _totalPage = (int)Math.Ceiling((double)_total / _pageSize);
            if (_totalPage <= 1)
            {
                btnFirst.Enabled = false;
                btnPrevious.Enabled = false;
                btnNext.Enabled = false;
                btnLast.Enabled = false;
                txtCurrentPage.Enabled = false;
            }else if (currentPage == 1)
            {
                btnFirst.Enabled = false;
                btnPrevious.Enabled = false;
                btnNext.Enabled = true;
                btnLast.Enabled = true;
                txtCurrentPage.Enabled = true;
            }else if (currentPage == _totalPage)
            {
                btnFirst.Enabled = true;
                btnPrevious.Enabled = true;
                btnNext.Enabled = false;
                btnLast.Enabled = false;
                txtCurrentPage.Enabled = true;
            }
            else
            {
                btnFirst.Enabled = true;
                btnPrevious.Enabled = true;
                btnNext.Enabled = true;
                btnLast.Enabled = true;
                txtCurrentPage.Enabled = true;
            }
            _currentPage = currentPage;
            _isChangeByUser = false;
            txtCurrentPage.Text = currentPage.ToString();
            _isChangeByUser = true;
        }

        #region 各点击事件的处理

        /// <summary>
        /// 首页
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFirst_Click(object sender, EventArgs e)
        {
            if (_method == null) return;
            _currentPage = 1;            
            _method(_currentPage);            
        }

        /// <summary>
        /// 上一页
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPrevious_Click(object sender, EventArgs e)
        {
            if (_method == null) return;
            _method(--_currentPage);
            //_currentPage--;
        }
        /// <summary>
        /// 下一页
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (_method == null) return;
            _method(++_currentPage);
            //_currentPage++;
        }

        /// <summary>
        /// 末页
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnLast_Click(object sender, EventArgs e)
        {
            if (_method == null) return;
            _currentPage = _totalPage;            
            _method(_currentPage);
        }

        /// <summary>
        /// 页码改变：用户手动输入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtCurrentPage_TextChanged(object sender, EventArgs e)
        {            
            if (!_isChangeByUser) return;
            if (_method == null) return;
            if (string.IsNullOrEmpty(txtCurrentPage.Text))
            {
                _currentPage = 1;
                _method(_currentPage);
                return;
            }
            //如果输入的不是整数类型，则为False
            if(int.TryParse(txtCurrentPage.Text, out _currentPage))
            {
                if (_currentPage < 1 || _currentPage == 0)
                {
                    _currentPage = 1;
                    _method(_currentPage);
                }
                else if(_currentPage > _totalPage)
                {
                    _currentPage = _totalPage;
                    _method(_totalPage);
                }
                else
                {
                    _method(_currentPage);
                }
            }
            else
            {
                MessageBox.Show("请输入有效数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        /// <summary>
        /// 每页记录数改变：用户选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboPageSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_method == null) return;
            if (_isChangeByUser)
            {
                _pageSize = Convert.ToInt32(cboPageSize.SelectedItem);
                _currentPage = 1;
                _method(_currentPage);
            }
        }

        #endregion
    }
}
