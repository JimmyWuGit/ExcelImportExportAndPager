using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace ExcelDemo
{
    public static class ExcelHelper
    {
        /// <summary>
        /// 从Excel导入：通用型(默认第一行为列名，第二行为数据起始行)
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static DataTable ImportToTable(string fileName)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(fileName).ToLower();
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (fileExt == ".xlsx")
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileExt == ".xls")
                {
                    workbook = new HSSFWorkbook(fs);
                }
                else
                {
                    workbook = null;
                    return null;
                }

                ISheet sheet = workbook.GetSheetAt(0);//Sheet总数量：workbook.NumberOfSheets

                //表头  
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else

                        dt.Columns.Add(new DataColumn(obj.ToString()));
                }
                //数据  
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    IRow row = sheet.GetRow(i);
                    for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        dr[j] = GetValueType(row.GetCell(j));
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }

                return dt;
            }

        }

        /// <summary>
        /// 从Excel导入：自定义数据的起始行
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="rowIndex">数据行(包括列名)</param>
        /// <returns></returns>
        public static DataTable ImportToTable(string fileName,int rowIndex)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(fileName).ToLower();
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (fileExt == ".xlsx")
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileExt == ".xls")
                {
                    workbook = new HSSFWorkbook(fs);
                }
                else
                {
                    workbook = null;
                    return null;
                }

                ISheet sheet = workbook.GetSheetAt(0);//Sheet总数量：workbook.NumberOfSheets

                //表头  
                IRow header = sheet.GetRow(rowIndex-1);
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else
                    {
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    }
                }
                //数据  
                for (int i = rowIndex; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    IRow row = sheet.GetRow(i);
                    for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        dr[j] = GetValueType(row.GetCell(j));
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }

                return dt;
            }

        }

        /// <summary>
        /// 从Excel导入：含合并单元格的导入，并自定义数据的起始行
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="rowIndex">数据行(包括列名)</param>
        /// <param name="isMergedCell">Excel是否包含合并的单元格</param>
        /// <returns></returns>
        public static DataTable ImportToTable(string fileName, int rowIndex,bool isMergedCell)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(fileName).ToLower();
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (fileExt == ".xlsx")
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileExt == ".xls")
                {
                    workbook = new HSSFWorkbook(fs);
                }
                else
                {
                    workbook = null;
                    return null;
                }

                ISheet sheet = workbook.GetSheetAt(0);//Sheet总数量：workbook.NumberOfSheets

                //表头  
                IRow header = sheet.GetRow(rowIndex - 1);
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else
                    {
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    }
                }
                //数据  
                for (int i = rowIndex; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    IRow row = sheet.GetRow(i);

                    for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null) continue;

                        if(cell.IsMergedCell && i > rowIndex) //检测列的单元格是否合并
                        {
                            //获取合并单元格的维度
                            Dimension dimension = GetDimension(sheet, i, j);                            
                            if (GetValueType(dimension.DataCell) == null) continue;

                            var cellValue = GetValueType(cell) == null ? "" : GetValueType(cell).ToString(); //单元格的值
                            if (string.IsNullOrEmpty(cellValue))
                            {
                                //空值由上一行的单元格的值获取
                                //dr[j] = dt.Rows[i - 2][j];
                                dr[j] = dimension.DataCell;
                            }
                            else
                            {
                                dr[j] = cellValue;
                                /*if (string.IsNullOrWhiteSpace(dr[j].ToString()) && j > 0)
                                {
                                    dr[j] = dr[j - 1];
                                }*/
                            }
                        }
                        else
                        {
                            dr[j] = GetValueType(cell); //不是合并的，则直接获取
                            /*if (string.IsNullOrWhiteSpace(dr[j].ToString()) && j > 0)
                            {
                                dr[j] = dr[j - 1];
                            }*/
                        }
                        
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }

                return dt;
            }

        }

        /// <summary>
        /// 从Excel导入：含合并单元格的导入，能自定义数据的起始行和添加新的列
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="rowIndex">数据行(包括列名)</param>
        /// <param name="isMergedCell">Excel是否包含合并的单元格</param>
        /// <param name="newColumns">新增的列名数组</param>
        /// <returns></returns>
        public static DataTable ImportToTable(string fileName, int rowIndex, bool isMergedCell,params string[] newColumns)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(fileName).ToLower();
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (fileExt == ".xlsx")
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileExt == ".xls")
                {
                    workbook = new HSSFWorkbook(fs);
                }
                else
                {
                    workbook = null;
                    return null;
                }

                ISheet sheet = workbook.GetSheetAt(0);//Sheet总数量：workbook.NumberOfSheets

                //获取工单号：第四行第五列
                string orderNo = GetValueType(sheet.GetRow(3).GetCell(4)).ToString();

                //表头  
                IRow header = sheet.GetRow(rowIndex - 1);
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else
                    {
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    }                    
                }
                if (newColumns.Length >= 1)
                {
                    //当遍历完原有的表头列名后，开始添加新增的列名
                    for (int k = 0; k < newColumns.Length; k++)
                    {
                        dt.Columns.Add(new DataColumn(newColumns[k]));
                    }
                }

                //数据  
                for (int i = rowIndex; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    IRow row = sheet.GetRow(i);

                    for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null) continue;

                        if (cell.IsMergedCell && i > rowIndex) //检测列的单元格是否合并
                        {
                            //获取合并单元格的维度
                            Dimension dimension = GetDimension(sheet, i, j);
                            //如果是Excel的最后一列，那么开始向新增的DataTable中插入值，且忽略合并单元格的起始行
                            if (j == row.LastCellNum - 1 && dt.Columns[j + 1].ColumnName.ToString() == "是否替代料" && dimension.MergedFirstRowIndex != i)
                            {
                                dr[j + 1] = "是";
                                dr[j + 2] = orderNo ?? "工单号为空";
                            }

                            //如果 合并单元格 整个为空，则直接跳过此次循环
                            if (GetValueType(dimension.DataCell) == null) continue;

                            var cellValue = GetValueType(cell) == null ? "" : GetValueType(cell).ToString(); //单元格的值
                            if (string.IsNullOrEmpty(cellValue))
                            {
                                //空值由合并单元格的值获取
                                dr[j] = dimension.DataCell;
                            }
                            else
                            {
                                dr[j] = cellValue;
                                /*if (string.IsNullOrWhiteSpace(dr[j].ToString()) && j > 0)
                                {
                                    dr[j] = dr[j - 1];
                                }*/
                            }                            
                        }
                        else
                        {
                            dr[j] = GetValueType(cell); //不是合并的，则直接获取
                            /*if (string.IsNullOrWhiteSpace(dr[j].ToString()) && j > 0)
                            {
                                dr[j] = dr[j - 1];
                            }*/
                        }

                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }

                return dt;
            }

        }

        /// <summary>
        /// 从DataTable导出到Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="fileName"></param>
        public static void ExportToExcel(DataTable dt,string fileName)
        {
            IWorkbook workbook = null;
            string fileExt = Path.GetExtension(fileName).ToLower();
            //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
            if(fileExt == ".xlsx")
            {
                workbook = new XSSFWorkbook();
            }else if(fileExt == ".xls")
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                return;
            }
            ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(dt.TableName);

            //表头
            IRow row = sheet.CreateRow(0);
            for(int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }
            //数据
            for(int i = 0;i < dt.Rows.Count; i++)
            {
                IRow rowNew = sheet.CreateRow(i + 1);
                for(int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = rowNew.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //保存为Excel文件
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        /// <summary>
        /// 数据类型的转换
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return null;
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    return cell.NumericCellValue;
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }

        /// <summary>
        /// 验证导入的Excel是否有数据
        /// </summary>
        /// <param name="excelFileStream"></param>
        /// <returns></returns>
        public static bool HasData(Stream excelFileStream)
        {
            using (excelFileStream)
            {
                IWorkbook workbook = new HSSFWorkbook(excelFileStream);
                if (workbook.NumberOfSheets > 0)
                {
                    ISheet sheet = workbook.GetSheetAt(0);
                    return sheet.PhysicalNumberOfRows > 0;
                }
            }
            return false;
        }

        /// <summary>
        /// 导出成功提示
        /// </summary>
        /// <param name="filePathAndName">文件路径</param>
        public static void ExportSuccessTips(string filePathAndName)
        {
            if (string.IsNullOrEmpty(filePathAndName))
            {
                MessageBox.Show("导出失败！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            if (MessageBox.Show("保存成功，是否打开文件？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(filePathAndName);
            }
        }

        /// <summary>
        /// 导入成功提示
        /// </summary>
        /// <param name="filePathAndName">文件路径</param>
        public static void ImportSuccessTips(string filePathAndName)
        {
            if (string.IsNullOrEmpty(filePathAndName)) return;

            MessageBox.Show("导入成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 输出 合并单元格 的维度
        /// </summary>
        /// <param name="rowIndex">行索引，从0开始</param>
        /// <param name="ColumnIndex">列索引，从0开始</param>
        /// <returns>单元格维度</returns>
        public static Dimension GetDimension(ISheet sheet,int rowIndex,int columnIndex)
        {
            Dimension dimension = new Dimension();

            //遍历一个Sheet中的所有合并单元格
            for(int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress rangeAddress = sheet.MergedRegions[i];
                //根据传入的行索引和列索引找到其所在的合并单元格
                if((rowIndex>=rangeAddress.FirstRow && rowIndex<=rangeAddress.LastRow) && (columnIndex>=rangeAddress.FirstColumn && columnIndex <= rangeAddress.LastColumn))
                {
                    dimension.DataCell = sheet.GetRow(rangeAddress.FirstRow).GetCell(rangeAddress.FirstColumn);
                    dimension.RowSpan = rangeAddress.LastRow - rangeAddress.FirstRow + 1;
                    dimension.ColumnSpan = rangeAddress.LastColumn - rangeAddress.FirstColumn + 1;
                    dimension.MergedFirstRowIndex = rangeAddress.FirstRow;
                    dimension.MergedLastRowIndex = rangeAddress.LastRow;
                    dimension.MergedFirstColumnIndex = rangeAddress.FirstColumn;
                    dimension.MergedLastColumnIndex = rangeAddress.LastColumn;

                    break;
                }
            }

            return dimension;
        }
    }

    /// <summary>
    /// 表示单元格的维度，通常用于表达合并单元格的维度
    /// </summary>
    public struct Dimension
    {
        /// <summary>
        /// 含有数据的单元格(通常表示合并单元格的第一个跨度行第一个跨度列)，该字段可能为null
        /// </summary>
        public ICell DataCell;

        /// <summary>
        /// 行跨度(跨越了多少行)
        /// </summary>
        public int RowSpan;

        /// <summary>
        /// 列跨度(跨越了多少列)
        /// </summary>
        public int ColumnSpan;

        /// <summary>
        /// 合并单元格的起始行索引
        /// </summary>
        public int MergedFirstRowIndex;

        /// <summary>
        /// 合并单元格的结束行索引
        /// </summary>
        public int MergedLastRowIndex;

        /// <summary>
        /// 合并单元格的起始列索引
        /// </summary>
        public int MergedFirstColumnIndex;

        /// <summary>
        /// 合并单元格的结束列索引
        /// </summary>
        public int MergedLastColumnIndex;
    }
}
