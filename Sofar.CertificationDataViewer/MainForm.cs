using Sofar.DataBaseHelper;
using Sofar.PointTableHelper.ExcelLib;
using Sunny.UI;
using System.Data;
using ClosedXML.Excel;
using System.Text;
using System.Globalization;
using System.Data.Common;
using System.Data.SQLite;
using System.Collections.Specialized;

namespace Sofar.CertificationDataViewer;

public partial class Form1 : UIForm
{

    /// <summary>
    /// 连接DB字符串
    /// </summary>
    private static string strConn = System.Configuration.ConfigurationManager.AppSettings["sqlite_strConn"].ToString();
    //private static string strConn = "Data Source = D:\\VS2022_CodeClone\\B5_认证数据工具\\Sofar.CertificationDataViewer\\db\\Certificate.db;Version=3;"; 
    SqlLiteHelper sqlHelper = new SqlLiteHelper(strConn);
    EmailHelper emailHelper = new EmailHelper();

    public Form1()
    {
        InitializeComponent();
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        //每周二8点半定时启动程序,更新数据库
        UpdateSqliteTbble();

        //刷新界面表格数据
        RefreshDataGridView();

        //刷新更新表
        EPPlusHelpr.ExportToXLSX_BasedOnOriginalTable(uiDataGridView1, Directory.GetCurrentDirectory() + "\\" + "产品认证统计表_原始.xlsx", Directory.GetCurrentDirectory() + "\\" + "产品认证统计表_更新.xlsx", false);

        // 获取当前时间  
        DateTime now = DateTime.Now;
        // 设置发送邮件时间范围-周二当天早上8点至9点允许发送
        DateTime startTime = new DateTime(now.Year, now.Month, now.Day, 8, 0, 0); // 当天8点0分0秒  
        DateTime endTime = new DateTime(now.Year, now.Month, now.Day, 9, 0, 0); // 当天9点0分0秒  

        // 判断当前时间是否在8点到9点之间  
        if (now >= startTime && now <= endTime && now.DayOfWeek == DayOfWeek.Tuesday)
        {
            // 发送预警邮件  
            SendWarningEmail();
        }
    }


    /// <summary>
    /// 发送预警邮件
    /// </summary>
    public void SendWarningEmail()
    {
        //获取产品线和产品系列的配置信息
        var productLineToProducts = new Dictionary<string, List<string>>();
        var productConfigSection = (NameValueCollection)System.Configuration.ConfigurationManager.GetSection("ProductConfigSection");
        foreach (string key in productConfigSection)
        {
            var products = productConfigSection[key].Split(',').Select(p => p.Trim()).ToList();
            productLineToProducts[key] = products;
        }

        //获取产品线和收件人的配置信息
        var productLineToEmails = new Dictionary<string, List<string>>();
        var recipientConfigSection = (NameValueCollection)System.Configuration.ConfigurationManager.GetSection("RecipientConfigSection");
        foreach (string key in recipientConfigSection)
        {
            var emails = recipientConfigSection[key].Split(',').Select(p => p.Trim()).ToList();
            productLineToEmails[key] = emails;
        }

        //筛选产品系列对应产品线且证书预警180天以内的记录
        var productLineToRows = new Dictionary<string, List<DataGridViewRow>>();
        foreach (var productLine in productLineToProducts.Keys)
        {
            productLineToRows[productLine] = new List<DataGridViewRow>();
        }

        foreach (DataGridViewRow row in uiDataGridView1.Rows)
        {
            if (!row.IsNewRow && row.Cells["证书预警"].Value != null && row.Cells["产品系列"].Value != null)
            {
                if (int.TryParse(row.Cells["证书预警"].Value.ToString(), out int warningValue))
                {
                    string product = row.Cells["产品系列"].Value.ToString();
                    foreach (var productLine in productLineToProducts)
                    {
                        if (productLine.Value.Contains(product))
                        {
                            productLineToRows[productLine.Key].Add(row);
                            break;
                        }
                    }
                }
            }
        }

        // 为满足条件的产品线生成相应的邮件内容,并发送给产品线下的所有收件人
        foreach (var productLine in productLineToRows.Keys)
        {
            if (productLineToRows[productLine].Count > 0)
            {
                var contentBuilder = new StringBuilder();

                // 邮件正文
                contentBuilder.AppendLine("Dear All,<br/>");
                contentBuilder.AppendLine($"早上好! 以下是{productLine}今日预警信息(所有产品认证证书信息详情见附件表格):<br/>");

                // 添加HTML表格的名称，并居中
                contentBuilder.AppendLine($"<h3 style='text-align: center;'>{productLine}证书预警信息表</h3>");

                // 添加HTML表格的开始标签，并居中
                contentBuilder.AppendLine("<table border='1' style='border-collapse: collapse; width: 100%; margin: 0 auto;'>");

                // 添加表头
                contentBuilder.AppendLine("<tr>");
                contentBuilder.AppendLine(string.Format("<th>{0}</th><th>{1}</th><th>{2}</th><th>{3}</th><th>{4}</th><th>{5}</th><th>{6}</th><th>{7}</th><th>{8}</th><th>{9}</th><th>{10}</th><th>{11}</th><th>{12}</th>",
                    "序号", "证书编号", "发证机构", "认证或检测标准", "产品系列", "发证日期", "证书有效期", "标准有效期", "可销售区域", "证书有效预警", "标准预警", "证书预警", "说明"));
                contentBuilder.AppendLine("</tr>");

                // 添加数据行
                foreach (var row in productLineToRows[productLine])
                {
                    contentBuilder.AppendLine("<tr>");
                    contentBuilder.AppendLine(string.Format("<td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td><td style='color: red;'>{11}</td><td>{12}</td>",
                        row.Cells["ID"].Value?.ToString() ?? "",
                        row.Cells["证书编号"].Value?.ToString() ?? "",
                        row.Cells["发证机构"].Value?.ToString() ?? "",
                        row.Cells["认证或检测标准"].Value?.ToString() ?? "",
                        row.Cells["产品系列"].Value?.ToString() ?? "",
                        row.Cells["发证日期"].Value?.ToString() ?? "",
                        row.Cells["证书有效期"].Value?.ToString() ?? "",
                        row.Cells["标准有效期"].Value?.ToString() ?? "",
                        row.Cells["可销售区域"].Value?.ToString() ?? "",
                        row.Cells["证书有效预警"].Value?.ToString() ?? "",
                        row.Cells["标准预警"].Value?.ToString() ?? "",
                        row.Cells["证书预警"].Value?.ToString() ?? "",
                        row.Cells["说明"].Value?.ToString() ?? ""));
                    contentBuilder.AppendLine("</tr>");
                }

                // 添加HTML表格的结束标签
                contentBuilder.AppendLine("</table>");
                // 添加表注
                contentBuilder.AppendLine("<p><strong><span style='color: red;'>注：</span>该表格列出了180天内即将到期的所有认证证书预警信息。</strong></p>");
                contentBuilder.AppendLine("如有任何问题，请随时与我联系。<br/>谢谢!<br/>");

                // 发送邮件
                if (productLineToEmails.ContainsKey(productLine))
                {
                    string recipients = string.Join(",", productLineToEmails[productLine]);
                    emailHelper.sendMail(contentBuilder.ToString(), recipients);
                }
            }
        }
    }

    /// <summary>
    /// 导入excel文件按钮
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void buttonSelectFile_Click(object sender, EventArgs e)
    {
        textBoxExcelPath.Text = GetExcelFilePath();
        if (string.IsNullOrEmpty(textBoxExcelPath.Text))
        {
            MessageBox.Show("No file selected.");
        }
    }

    /// <summary>
    /// 获取excel文件路径
    /// </summary>
    /// <returns></returns>
    public static string GetExcelFilePath()
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "Import File| *.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Select an Excel File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }
        }

        return null;
    }

    /// <summary>
    /// Excel文件导入数据库按钮
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void buttonImport_Click(object sender, EventArgs e)
    {

        if (string.IsNullOrEmpty(textBoxExcelPath.Text))
        {
            MessageBox.Show("Please provide both Excel file path .");
            return;
        }

        if (ImportTable(textBoxExcelPath.Text)) MessageBox.Show("导入数据库成功");

    }

    /// <summary>
    /// 检查导入的excel文件，更新表格内容
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    public virtual bool ImportTable(string filePath)
    {

        List<DataTable> dtList = EPPlusHelpr.LoadFromFile(filePath);

        foreach (var dt in dtList)
        {

            foreach (DataRow row in dt.Rows)
            {
                string CertificateValidityPeriod = row["证书有效期"].ToString();
                string StandardValidityPeriod = row["标准有效期"].ToString();

                row["证书有效预警"] = CheckStatus_CertificateValidityPeriod(CertificateValidityPeriod);
                row["标准预警"] = CheckStatus_StandardValidityPeriod(StandardValidityPeriod);
                row["证书预警"] = CheckStatus_CertificateValidityPeriod2(CertificateValidityPeriod);
            }

            this.ImportDt(dt);
        }
        return true;
    }

    /// <summary>
    /// 通过检查证书有效期获取证书有效预警
    /// </summary>
    /// <param name="CertificateValidityPeriod"></param>
    /// <returns></returns>
    public static string CheckStatus_CertificateValidityPeriod(string CertificateValidityPeriod)
    {
        if (CertificateValidityPeriod == "" || CertificateValidityPeriod == "-" || CertificateValidityPeriod == "长期有效")
        {
            return CertificateValidityPeriod;
        }

        // 将字符串转换为日期  
        DateTime dateInCertificateValidityPeriod;
        if (!DateTime.TryParseExact(CertificateValidityPeriod, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateInCertificateValidityPeriod))
        {

            return CertificateValidityPeriod + "为无效的日期格式，请检查该行的证书有效期格式";
        }


        // 检查今天日期是否大于证书有效期 
        if (DateTime.Today > dateInCertificateValidityPeriod)
        {
            return "过期";
        }

        // 检查证书有效期与今天日期的差是否小于等于180天  
        TimeSpan difference = dateInCertificateValidityPeriod - DateTime.Today;
        if (difference.Days <= 180)
        {
            // 如果小于等于180天，则返回剩余的天数  
            return $"{difference.Days}";
        }
        return "正常";
    }

    /// <summary>
    /// 通过检查标准有效期获取标准预警
    /// </summary>
    /// <param name="StandardValidityPeriod"></param>
    /// <returns></returns>
    public static string CheckStatus_StandardValidityPeriod(string StandardValidityPeriod)
    {
        if (StandardValidityPeriod == "" || StandardValidityPeriod == "-" || StandardValidityPeriod == "过期" || StandardValidityPeriod == "长期有效")
        {
            return StandardValidityPeriod;
        }

        // 将字符串转换为日期  
        DateTime dateInStandardValidityPeriod;
        if (!DateTime.TryParseExact(StandardValidityPeriod, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateInStandardValidityPeriod))
        {

            return StandardValidityPeriod + "为无效的日期格式，请检查该行的标准有效期格式";
        }

        // 检查今天日期是否大于标准有效期  
        if (DateTime.Today > dateInStandardValidityPeriod)
        {
            return "过期";
        }

        // 检查标准有效期与今天日期的差是否小于等于180天  
        TimeSpan difference = dateInStandardValidityPeriod - DateTime.Today;
        if (difference.Days <= 180)
        {
            // 如果小于等于180天，则返回剩余的天数  
            return $"{difference.Days}";
        }
        return "正常";
    }

    /// <summary>
    /// 通过检查证书有效期获取证书预警
    /// </summary>
    /// <param name="CertificateValidityPeriod"></param>
    /// <returns></returns>
    public static string CheckStatus_CertificateValidityPeriod2(string CertificateValidityPeriod)
    {
        if (CertificateValidityPeriod == "" || CertificateValidityPeriod == "-" || CertificateValidityPeriod == "过期" || CertificateValidityPeriod == "长期有效")
        {
            return CertificateValidityPeriod;
        }

        // 将字符串转换为日期  
        DateTime dateInCertificateValidityPeriod;
        if (!DateTime.TryParseExact(CertificateValidityPeriod, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateInCertificateValidityPeriod))
        {

            return CertificateValidityPeriod + "为无效的日期格式，请检查该行的证书有效期格式";
        }

        // 检查今天日期是否大于证书有效期  
        if (DateTime.Today > dateInCertificateValidityPeriod)
        {
            return "过期";
        }

        // 检查证书有效期与今天日期的差是否小于等于180天  
        TimeSpan difference = dateInCertificateValidityPeriod - DateTime.Today;
        if (difference.Days <= 180)
        {
            // 如果小于等于180天，则返回剩余的天数  
            return $"{difference.Days}";
        }
        return "正常";
    }


    /// <summary>
    /// 导入数据库
    /// </summary>
    /// <param name="table"></param>
    /// <param name="dt"></param>
    public void ImportDt(DataTable dt)
    {
        string tableName = "Test";//表名
        string tableProperty = string.Join(",", dt.Columns.Cast<DataColumn>().Select(col => col.ColumnName));
        string tableProp = string.Join(",", dt.Columns.Cast<DataColumn>().Select(col => $"@{col.ColumnName}"));
        try
        {
            if (string.IsNullOrEmpty(tableName) || string.IsNullOrEmpty(tableProperty) || string.IsNullOrEmpty(tableProp))
                return;

            string commandText = $"INSERT INTO {tableName} ({tableProperty}) VALUES({tableProp})";
            sqlHelper.ExecuteNonQuery($"delete from {tableName}");//操作前清空表(仅数据)，避免重复
            sqlHelper.ExecuteMutliQuery(commandText, dt);
        }
        catch (Exception ex)
        {
            MessageBox.Show("导入数据库错误，info: " + ex.Message);
        }
    }

    /// <summary>
    /// 更新数据库
    /// </summary>
    public void UpdateSqliteTbble()
    {
        DataTable dt = new DataTable();
        string sql = $"select * from Test";
        try
        {
            dt = sqlHelper.ExecuteTable(sql);

            foreach (DataRow row in dt.Rows)
            {
                string CertificateValidityPeriod = row["证书有效期"].ToString();
                string StandardValidityPeriod = row["标准有效期"].ToString();

                row["证书有效预警"] = CheckStatus_CertificateValidityPeriod(CertificateValidityPeriod);
                row["标准预警"] = CheckStatus_StandardValidityPeriod(StandardValidityPeriod);
                row["证书预警"] = CheckStatus_CertificateValidityPeriod2(CertificateValidityPeriod);
            }

            this.ImportDt(dt);
        }
        catch (Exception ex)
        {
            MessageBox.Show("发生错误: " + ex.Message);
        }
    }
    /// <summary>
    /// 查询数据库指定表内容
    /// </summary>
    /// <returns></returns>
    public DataTable DbPointTable()
    {
        DataTable dt = new DataTable();
        string sql = $"SELECT * FROM Test ORDER BY ID ASC";//按ID升序排列

        try
        {
            dt = sqlHelper.ExecuteTable(sql);
            //数据库表绑定至界面表格
            uiDataGridView1.DataSource = dt;

            if (dt.Rows.Count <= 0)
                return default;
        }
        catch (Exception ex)
        {
            MessageBox.Show("发生错误: " + ex.Message);
        }
        return dt;
    }

    /// <summary>
    /// 手动查询数据库刷新界面表格按钮
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void btnRead_Click_1(object sender, EventArgs e)
    {
        RefreshDataGridView();
    }

    /// <summary>
    /// 导出为CSV文件
    /// </summary>
    /// <param name="dataGridView"></param>
    /// <param name="filePath"></param>
    public void ExportToCSV(DataGridView dataGridView, string filePath, bool prompt)
    {

        // 检查文件路径是否为空  
        if (string.IsNullOrEmpty(filePath))
        {
            if (prompt) MessageBox.Show("请指定一个文件路径。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // 检查 DataGridView 是否有数据  
        if (uiDataGridView1.Rows.Count == 0)
        {
            if (prompt) MessageBox.Show("DataGridView 中没有数据。", "注意", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            // 使用 StreamWriter 创建或打开文件  
            using (StreamWriter sw = new StreamWriter(filePath, false))
            {
                // 写入列标题  
                for (int i = 0; i < uiDataGridView1.Columns.Count; i++)
                {
                    if (i > 0) sw.Write(",");
                    sw.Write(uiDataGridView1.Columns[i].HeaderText);
                }
                sw.WriteLine();

                // 写入行数据  
                foreach (DataGridViewRow row in uiDataGridView1.Rows)
                {
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        if (i > 0) sw.Write(",");

                        // 检查是否为空值，避免写入空字符串  
                        if (row.Cells[i].Value != null)
                        {
                            // 如果包含逗号、换行符等特殊字符，则添加引号并转义内部引号  
                            string value = row.Cells[i].Value.ToString().Replace(",", "\",\"")
                                .Replace(Environment.NewLine, "\\n")
                                .Replace("\"", "\"\"");
                            sw.Write("\"" + value + "\"");
                        }
                    }
                    sw.WriteLine();
                }
            }

            if (prompt) MessageBox.Show("导出成功。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            if (prompt) MessageBox.Show("导出失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

    }

    /// <summary>
    /// 导出为XLSX文件
    /// </summary>
    /// <param name="dataGridView"></param>
    /// <param name="filePath"></param>
    public void ExportToXLSX(DataGridView dataGridView, string filePath, bool prompt)
    {
        // 检查文件路径是否为空  
        if (string.IsNullOrEmpty(filePath))
        {
            if (prompt) MessageBox.Show("请指定一个文件路径。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // 检查 DataGridView 是否有数据  
        if (dataGridView.Rows.Count == 0)
        {
            if (prompt) MessageBox.Show("DataGridView 中没有数据。", "注意", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            // 创建一个新的 Excel 工作簿  
            var workbook = new XLWorkbook();
            // 添加一个工作表  
            var worksheet = workbook.Worksheets.Add("证书有效性");

            // 写入列标题  
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                worksheet.Cell(1, i + 1).Value = dataGridView.Columns[i].HeaderText;
            }

            // 写入行数据  
            for (int row = 0; row < dataGridView.Rows.Count; row++)
            {
                for (int col = 0; col < dataGridView.Columns.Count; col++)
                {
                    // 检查是否为空值  
                    if (dataGridView.Rows[row].Cells[col].Value != null)
                    {
                        worksheet.Cell(row + 2, col + 1).Value = dataGridView.Rows[row].Cells[col].Value.ToString();
                    }
                }
            }

            // 保存工作簿  
            workbook.SaveAs(filePath);

            if (prompt) MessageBox.Show("导出成功。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            if (prompt) MessageBox.Show("导出失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// 导出xlsx文件按钮
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void btnExport_Click(object sender, EventArgs e)
    {

        //SaveFileDialog saveFileDialog = new SaveFileDialog();
        //saveFileDialog.Filter = "XLSX files (*.xlsx)|*.xlsx";
        //saveFileDialog.Title = "导出 xlsx 文件";

        //if (saveFileDialog.ShowDialog() == DialogResult.OK)
        //{
        //    string filePath = saveFileDialog.FileName;
        //    ExportToXLSX(uiDataGridView1, filePath, true);
        //}

        //SaveFileDialog saveFileDialog = new SaveFileDialog();
        //saveFileDialog.Filter = "CSV files (*.csv)|*.csv";
        //saveFileDialog.Title = "导出 CSV 文件";

        //if (saveFileDialog.ShowDialog() == DialogResult.OK)
        //{
        //    string filePath = saveFileDialog.FileName;
        //    ExportToCSV(uiDataGridView1, filePath, true);
        //}

        string OriginalTablePath = Directory.GetCurrentDirectory() + "\\" + "产品认证统计表_原始.xlsx"; // 原始表路径
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "XLSX files (*.xlsx)|*.xlsx";

        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            string filePath = saveFileDialog.FileName;
            EPPlusHelpr.ExportToXLSX_BasedOnOriginalTable(uiDataGridView1, OriginalTablePath, filePath, true);
        }

    }

    /// <summary>
    /// 新增表格数据，同步数据库
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void toolStripMenuItem1_Click(object sender, EventArgs e)
    {
        int insertId;
        if (uiDataGridView1.SelectedRows.Count > 0)
        {
            // 获取当前选中行的索引
            int selectedRowIndex = uiDataGridView1.SelectedRows[0].Index;
            DataGridViewRow selectedRow = uiDataGridView1.Rows[selectedRowIndex];
            int selectedRowId = Convert.ToInt32(selectedRow.Cells["ID"].Value);
            insertId = selectedRowId + 1;//选中ID的下一个插入
        }
        else
        {
            // 获取数据库中最大ID，加1
            insertId = GetMaxId() + 1;
        }

        // 弹出输入数据的对话框，并传入初始数据
        InsertForm insertForm = new InsertForm();
        insertForm.GetID(insertId.ToString());//默认将目标ID赋值文本框
        DialogResult result = insertForm.ShowDialog();

        if (result == DialogResult.OK)
        {
            // 获取用户输入的数据
            string[] newData = insertForm.GetInsertedData();

            // 用户输入的ID
            if (int.TryParse(newData[0], out insertId) && insertId >= 1)
            {
                UpdateSubsequentIds(insertId);//更新新增的ID前，修改受影响行ID                              
            }
            else
            {
                MessageBox.Show("输入的ID无效，请输入一个大于或者等于1的整数！");
                return;
            }

            // 将新数据封装成字典
            Dictionary<string, string> insertData = new Dictionary<string, string>
            {
                { "ID",             newData[0] },
                { "证书编号",       newData[1] },
                { "发证机构",       newData[2] },
                { "认证或检测标准", newData[3] },
                { "产品系列",       newData[4] },
                { "发证日期",       newData[5] },
                { "证书有效期",     newData[6] },
                { "标准有效期",     newData[7] },
                { "可销售区域",     newData[8] },
                { "证书有效预警",   CheckStatus_CertificateValidityPeriod(newData[6])},
                { "标准预警",       CheckStatus_StandardValidityPeriod(newData[7])},
                { "证书预警",       CheckStatus_CertificateValidityPeriod2(newData[6]) },
                { "说明",           newData[12] },
            };
            // 插入数据到数据库
            sqlHelper.ExecuteInsert("Test", insertData);
            //刷新界面表格
            RefreshDataGridView();
        }
    }

    /// <summary>
    /// 插入前先将大于等于目标ID的所有ID+1
    /// </summary>
    /// <param name="insertId"></param>
    private void UpdateSubsequentIds(int insertId)
    {
        string query = @"UPDATE Test SET ID = ID + 1 WHERE ID >= @insertId";
        var parameters = new List<DbParameter>
        {
              new SQLiteParameter("@insertId", insertId)
        };
        sqlHelper.ExecuteNonQuery(query, parameters.ToArray());
    }

    /// <summary>
    /// 从数据库中获取最大ID
    /// </summary>
    /// <returns></returns>
    private int GetMaxId()
    {
        string query = "SELECT MAX(ID) FROM Test";
        object result = sqlHelper.ExecuteScalar(query);
        return result == DBNull.Value ? 0 : Convert.ToInt32(result);
    }

    /// <summary>
    /// 删除表格数据，同步数据库(支持删除多个连续、不连续ID，删除后使ID保持连续自增)
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void toolStripMenuItem2_Click(object sender, EventArgs e)
    {
        // 确保有行被选中
        if (uiDataGridView1.SelectedRows.Count > 0)
        {
            // 记录要删除的ID
            List<string> idsToDelete = new List<string>();

            foreach (DataGridViewRow selectedRow in uiDataGridView1.SelectedRows)
            {
                string id = selectedRow.Cells["ID"].Value.ToString();
                idsToDelete.Add(id);
            }

            if (idsToDelete.Count > 0)
            {
                // 创建一个确认删除弹窗消息
                string message = $"请确认要删除如下ID的所有信息吗？\n{string.Join("\n", idsToDelete)}";
                DialogResult result = MessageBox.Show(message, "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    // 执行删除操作
                    foreach (string id in idsToDelete)
                    {
                        sqlHelper.ExecuteDelete("Test", $"ID='{id}'");
                    }
                    // 更新数据库ID(升序且连续)
                    UpdateIDsSequential();

                    //刷新界面表格
                    RefreshDataGridView();
                    MessageBox.Show("删除成功！");
                }
            }
        }
        else
        {
            MessageBox.Show("请点击最左侧空白列，选中需要删除的行！");
        }
    }

    /// <summary>
    /// 更新数据库ID(升序且连续)
    /// </summary>
    private void UpdateIDsSequential()
    {
        // 从数据库中获取所有数据

        var dataTable = sqlHelper.ExecuteTable("SELECT * FROM Test ORDER BY ID");

        // 更新ID使其连续
        int newId = 1;
        foreach (DataRow row in dataTable.Rows)
        {
            string oldId = row["ID"].ToString();
            string updateQuery = $"UPDATE Test SET ID={newId} WHERE ID='{oldId}'";
            sqlHelper.ExecuteNonQuery(updateQuery);
            newId++;
        }
    }

    /// <summary>
    /// 修改表格数据，同步数据库
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void toolStripMenuItem3_Click(object sender, EventArgs e)
    {

        // 获取当前选中行的数据
        if (uiDataGridView1.SelectedRows.Count > 0)
        {
            DataGridViewRow selectedRow = uiDataGridView1.SelectedRows[0];
            string[] originalData = new string[]
            {
            selectedRow.Cells["ID"].Value.ToString(),
            selectedRow.Cells["证书编号"].Value.ToString(),
            selectedRow.Cells["发证机构"].Value.ToString(),
            selectedRow.Cells["认证或检测标准"].Value.ToString(),
            selectedRow.Cells["产品系列"].Value.ToString(),
            selectedRow.Cells["发证日期"].Value.ToString(),
            selectedRow.Cells["证书有效期"].Value.ToString(),
            selectedRow.Cells["标准有效期"].Value.ToString(),
            selectedRow.Cells["可销售区域"].Value.ToString(),
            selectedRow.Cells["证书有效预警"].Value.ToString(),
            selectedRow.Cells["标准预警"].Value.ToString(),
            selectedRow.Cells["证书预警"].Value.ToString(),
            selectedRow.Cells["说明"].Value.ToString()
            };

            // 弹出修改数据的对话框         
            InsertForm insertForm = new InsertForm();
            insertForm.GetSelectedData(originalData);
            DialogResult result = insertForm.ShowDialog();

            if (result == DialogResult.OK)
            {
                // 获取用户修改后的数据
                string[] editedData = insertForm.GetInsertedData();
                Dictionary<string, string> updateData = new Dictionary<string, string>
                {
                    { "ID",             editedData[0] },
                    { "证书编号",       editedData[1] },
                    { "发证机构",       editedData[2] },
                    { "认证或检测标准", editedData[3] },
                    { "产品系列",       editedData[4] },
                    { "发证日期",       editedData[5] },
                    { "证书有效期",     editedData[6] },
                    { "标准有效期",     editedData[7] },
                    { "可销售区域",     editedData[8] },
                    { "证书有效预警",   CheckStatus_CertificateValidityPeriod(editedData[6])},
                    { "标准预警",       CheckStatus_StandardValidityPeriod(editedData[7])},
                    { "证书预警",       CheckStatus_CertificateValidityPeriod2(editedData[6]) },
                    { "说明",           editedData[12] },
                };

                // 更新数据到数据库
                string id = selectedRow.Cells["ID"].Value.ToString();
                sqlHelper.ExecuteUpdate("Test", $"ID={id}", updateData);

                //刷新界面表格
                RefreshDataGridView();

            }
        }
        else
        {
            MessageBox.Show("请点击最左侧空白列，选中需要修改的行！");
        }
    }


    /// <summary>
    /// 刷新界面表格
    /// </summary>
    private void RefreshDataGridView()
    {
        //查询数据库表
        DbPointTable();
        uiDataGridView1.ContextMenuStrip = uiContextMenuStrip1;

        foreach (DataGridViewRow row in uiDataGridView1.Rows)
        {
            ProcessDateCell_Display("证书有效期", row, uiDataGridView1);
            ProcessDateCell_Display("标准有效期", row, uiDataGridView1);
            ProcessDateCell_Display("证书有效预警", row, uiDataGridView1);
            ProcessDateCell_Display("标准预警", row, uiDataGridView1);
            ProcessDateCell_Display("证书预警", row, uiDataGridView1);
        }
    }

    /// <summary>
    /// 指定列内容为过期，内容标注红色
    /// </summary>
    /// <param name="columnName"></param>
    /// <param name="row"></param>
    /// <param name="uiDataGridView"></param>
    private void ProcessDateCell_Display(string columnName, DataGridViewRow row, DataGridView uiDataGridView)
    {
        if (row.Cells[columnName].Value != null)
        {
            if (row.Cells[columnName].Value.ToString() == "过期")
            {
                row.Cells[columnName].Style.ForeColor = Color.Red;
            }

            if (columnName == "证书预警")
            {
                if (int.TryParse(row.Cells[columnName].Value.ToString(), out int warningValue))
                {
                    row.Cells[columnName].Style.ForeColor = Color.Yellow;
                    row.Cells[columnName].Style.BackColor = Color.Red;
                }
            }
        }
    }
}




