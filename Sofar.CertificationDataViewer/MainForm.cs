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
    /// ����DB�ַ���
    /// </summary>
    private static string strConn = System.Configuration.ConfigurationManager.AppSettings["sqlite_strConn"].ToString();
    //private static string strConn = "Data Source = D:\\VS2022_CodeClone\\B5_��֤���ݹ���\\Sofar.CertificationDataViewer\\db\\Certificate.db;Version=3;"; 
    SqlLiteHelper sqlHelper = new SqlLiteHelper(strConn);
    EmailHelper emailHelper = new EmailHelper();

    public Form1()
    {
        InitializeComponent();
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        //ÿ�ܶ�8��붨ʱ��������,�������ݿ�
        UpdateSqliteTbble();

        //ˢ�½���������
        RefreshDataGridView();

        //ˢ�¸��±�
        EPPlusHelpr.ExportToXLSX_BasedOnOriginalTable(uiDataGridView1, Directory.GetCurrentDirectory() + "\\" + "��Ʒ��֤ͳ�Ʊ�_ԭʼ.xlsx", Directory.GetCurrentDirectory() + "\\" + "��Ʒ��֤ͳ�Ʊ�_����.xlsx", false);

        // ��ȡ��ǰʱ��  
        DateTime now = DateTime.Now;
        // ���÷����ʼ�ʱ�䷶Χ-�ܶ���������8����9��������
        DateTime startTime = new DateTime(now.Year, now.Month, now.Day, 8, 0, 0); // ����8��0��0��  
        DateTime endTime = new DateTime(now.Year, now.Month, now.Day, 9, 0, 0); // ����9��0��0��  

        // �жϵ�ǰʱ���Ƿ���8�㵽9��֮��  
        if (now >= startTime && now <= endTime && now.DayOfWeek == DayOfWeek.Tuesday)
        {
            // ����Ԥ���ʼ�  
            SendWarningEmail();
        }
    }


    /// <summary>
    /// ����Ԥ���ʼ�
    /// </summary>
    public void SendWarningEmail()
    {
        //��ȡ��Ʒ�ߺͲ�Ʒϵ�е�������Ϣ
        var productLineToProducts = new Dictionary<string, List<string>>();
        var productConfigSection = (NameValueCollection)System.Configuration.ConfigurationManager.GetSection("ProductConfigSection");
        foreach (string key in productConfigSection)
        {
            var products = productConfigSection[key].Split(',').Select(p => p.Trim()).ToList();
            productLineToProducts[key] = products;
        }

        //��ȡ��Ʒ�ߺ��ռ��˵�������Ϣ
        var productLineToEmails = new Dictionary<string, List<string>>();
        var recipientConfigSection = (NameValueCollection)System.Configuration.ConfigurationManager.GetSection("RecipientConfigSection");
        foreach (string key in recipientConfigSection)
        {
            var emails = recipientConfigSection[key].Split(',').Select(p => p.Trim()).ToList();
            productLineToEmails[key] = emails;
        }

        //ɸѡ��Ʒϵ�ж�Ӧ��Ʒ����֤��Ԥ��180�����ڵļ�¼
        var productLineToRows = new Dictionary<string, List<DataGridViewRow>>();
        foreach (var productLine in productLineToProducts.Keys)
        {
            productLineToRows[productLine] = new List<DataGridViewRow>();
        }

        foreach (DataGridViewRow row in uiDataGridView1.Rows)
        {
            if (!row.IsNewRow && row.Cells["֤��Ԥ��"].Value != null && row.Cells["��Ʒϵ��"].Value != null)
            {
                if (int.TryParse(row.Cells["֤��Ԥ��"].Value.ToString(), out int warningValue))
                {
                    string product = row.Cells["��Ʒϵ��"].Value.ToString();
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

        // Ϊ���������Ĳ�Ʒ��������Ӧ���ʼ�����,�����͸���Ʒ���µ������ռ���
        foreach (var productLine in productLineToRows.Keys)
        {
            if (productLineToRows[productLine].Count > 0)
            {
                var contentBuilder = new StringBuilder();

                // �ʼ�����
                contentBuilder.AppendLine("Dear All,<br/>");
                contentBuilder.AppendLine($"���Ϻ�! ������{productLine}����Ԥ����Ϣ(���в�Ʒ��֤֤����Ϣ������������):<br/>");

                // ���HTML�������ƣ�������
                contentBuilder.AppendLine($"<h3 style='text-align: center;'>{productLine}֤��Ԥ����Ϣ��</h3>");

                // ���HTML���Ŀ�ʼ��ǩ��������
                contentBuilder.AppendLine("<table border='1' style='border-collapse: collapse; width: 100%; margin: 0 auto;'>");

                // ��ӱ�ͷ
                contentBuilder.AppendLine("<tr>");
                contentBuilder.AppendLine(string.Format("<th>{0}</th><th>{1}</th><th>{2}</th><th>{3}</th><th>{4}</th><th>{5}</th><th>{6}</th><th>{7}</th><th>{8}</th><th>{9}</th><th>{10}</th><th>{11}</th><th>{12}</th>",
                    "���", "֤����", "��֤����", "��֤�����׼", "��Ʒϵ��", "��֤����", "֤����Ч��", "��׼��Ч��", "����������", "֤����ЧԤ��", "��׼Ԥ��", "֤��Ԥ��", "˵��"));
                contentBuilder.AppendLine("</tr>");

                // ���������
                foreach (var row in productLineToRows[productLine])
                {
                    contentBuilder.AppendLine("<tr>");
                    contentBuilder.AppendLine(string.Format("<td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td><td style='color: red;'>{11}</td><td>{12}</td>",
                        row.Cells["ID"].Value?.ToString() ?? "",
                        row.Cells["֤����"].Value?.ToString() ?? "",
                        row.Cells["��֤����"].Value?.ToString() ?? "",
                        row.Cells["��֤�����׼"].Value?.ToString() ?? "",
                        row.Cells["��Ʒϵ��"].Value?.ToString() ?? "",
                        row.Cells["��֤����"].Value?.ToString() ?? "",
                        row.Cells["֤����Ч��"].Value?.ToString() ?? "",
                        row.Cells["��׼��Ч��"].Value?.ToString() ?? "",
                        row.Cells["����������"].Value?.ToString() ?? "",
                        row.Cells["֤����ЧԤ��"].Value?.ToString() ?? "",
                        row.Cells["��׼Ԥ��"].Value?.ToString() ?? "",
                        row.Cells["֤��Ԥ��"].Value?.ToString() ?? "",
                        row.Cells["˵��"].Value?.ToString() ?? ""));
                    contentBuilder.AppendLine("</tr>");
                }

                // ���HTML���Ľ�����ǩ
                contentBuilder.AppendLine("</table>");
                // ��ӱ�ע
                contentBuilder.AppendLine("<p><strong><span style='color: red;'>ע��</span>�ñ���г���180���ڼ������ڵ�������֤֤��Ԥ����Ϣ��</strong></p>");
                contentBuilder.AppendLine("�����κ����⣬����ʱ������ϵ��<br/>лл!<br/>");

                // �����ʼ�
                if (productLineToEmails.ContainsKey(productLine))
                {
                    string recipients = string.Join(",", productLineToEmails[productLine]);
                    emailHelper.sendMail(contentBuilder.ToString(), recipients);
                }
            }
        }
    }

    /// <summary>
    /// ����excel�ļ���ť
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
    /// ��ȡexcel�ļ�·��
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
    /// Excel�ļ��������ݿⰴť
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

        if (ImportTable(textBoxExcelPath.Text)) MessageBox.Show("�������ݿ�ɹ�");

    }

    /// <summary>
    /// ��鵼���excel�ļ������±������
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
                string CertificateValidityPeriod = row["֤����Ч��"].ToString();
                string StandardValidityPeriod = row["��׼��Ч��"].ToString();

                row["֤����ЧԤ��"] = CheckStatus_CertificateValidityPeriod(CertificateValidityPeriod);
                row["��׼Ԥ��"] = CheckStatus_StandardValidityPeriod(StandardValidityPeriod);
                row["֤��Ԥ��"] = CheckStatus_CertificateValidityPeriod2(CertificateValidityPeriod);
            }

            this.ImportDt(dt);
        }
        return true;
    }

    /// <summary>
    /// ͨ�����֤����Ч�ڻ�ȡ֤����ЧԤ��
    /// </summary>
    /// <param name="CertificateValidityPeriod"></param>
    /// <returns></returns>
    public static string CheckStatus_CertificateValidityPeriod(string CertificateValidityPeriod)
    {
        if (CertificateValidityPeriod == "" || CertificateValidityPeriod == "-" || CertificateValidityPeriod == "������Ч")
        {
            return CertificateValidityPeriod;
        }

        // ���ַ���ת��Ϊ����  
        DateTime dateInCertificateValidityPeriod;
        if (!DateTime.TryParseExact(CertificateValidityPeriod, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateInCertificateValidityPeriod))
        {

            return CertificateValidityPeriod + "Ϊ��Ч�����ڸ�ʽ��������е�֤����Ч�ڸ�ʽ";
        }


        // �����������Ƿ����֤����Ч�� 
        if (DateTime.Today > dateInCertificateValidityPeriod)
        {
            return "����";
        }

        // ���֤����Ч����������ڵĲ��Ƿ�С�ڵ���180��  
        TimeSpan difference = dateInCertificateValidityPeriod - DateTime.Today;
        if (difference.Days <= 180)
        {
            // ���С�ڵ���180�죬�򷵻�ʣ�������  
            return $"{difference.Days}";
        }
        return "����";
    }

    /// <summary>
    /// ͨ������׼��Ч�ڻ�ȡ��׼Ԥ��
    /// </summary>
    /// <param name="StandardValidityPeriod"></param>
    /// <returns></returns>
    public static string CheckStatus_StandardValidityPeriod(string StandardValidityPeriod)
    {
        if (StandardValidityPeriod == "" || StandardValidityPeriod == "-" || StandardValidityPeriod == "����" || StandardValidityPeriod == "������Ч")
        {
            return StandardValidityPeriod;
        }

        // ���ַ���ת��Ϊ����  
        DateTime dateInStandardValidityPeriod;
        if (!DateTime.TryParseExact(StandardValidityPeriod, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateInStandardValidityPeriod))
        {

            return StandardValidityPeriod + "Ϊ��Ч�����ڸ�ʽ��������еı�׼��Ч�ڸ�ʽ";
        }

        // �����������Ƿ���ڱ�׼��Ч��  
        if (DateTime.Today > dateInStandardValidityPeriod)
        {
            return "����";
        }

        // ����׼��Ч����������ڵĲ��Ƿ�С�ڵ���180��  
        TimeSpan difference = dateInStandardValidityPeriod - DateTime.Today;
        if (difference.Days <= 180)
        {
            // ���С�ڵ���180�죬�򷵻�ʣ�������  
            return $"{difference.Days}";
        }
        return "����";
    }

    /// <summary>
    /// ͨ�����֤����Ч�ڻ�ȡ֤��Ԥ��
    /// </summary>
    /// <param name="CertificateValidityPeriod"></param>
    /// <returns></returns>
    public static string CheckStatus_CertificateValidityPeriod2(string CertificateValidityPeriod)
    {
        if (CertificateValidityPeriod == "" || CertificateValidityPeriod == "-" || CertificateValidityPeriod == "����" || CertificateValidityPeriod == "������Ч")
        {
            return CertificateValidityPeriod;
        }

        // ���ַ���ת��Ϊ����  
        DateTime dateInCertificateValidityPeriod;
        if (!DateTime.TryParseExact(CertificateValidityPeriod, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateInCertificateValidityPeriod))
        {

            return CertificateValidityPeriod + "Ϊ��Ч�����ڸ�ʽ��������е�֤����Ч�ڸ�ʽ";
        }

        // �����������Ƿ����֤����Ч��  
        if (DateTime.Today > dateInCertificateValidityPeriod)
        {
            return "����";
        }

        // ���֤����Ч����������ڵĲ��Ƿ�С�ڵ���180��  
        TimeSpan difference = dateInCertificateValidityPeriod - DateTime.Today;
        if (difference.Days <= 180)
        {
            // ���С�ڵ���180�죬�򷵻�ʣ�������  
            return $"{difference.Days}";
        }
        return "����";
    }


    /// <summary>
    /// �������ݿ�
    /// </summary>
    /// <param name="table"></param>
    /// <param name="dt"></param>
    public void ImportDt(DataTable dt)
    {
        string tableName = "Test";//����
        string tableProperty = string.Join(",", dt.Columns.Cast<DataColumn>().Select(col => col.ColumnName));
        string tableProp = string.Join(",", dt.Columns.Cast<DataColumn>().Select(col => $"@{col.ColumnName}"));
        try
        {
            if (string.IsNullOrEmpty(tableName) || string.IsNullOrEmpty(tableProperty) || string.IsNullOrEmpty(tableProp))
                return;

            string commandText = $"INSERT INTO {tableName} ({tableProperty}) VALUES({tableProp})";
            sqlHelper.ExecuteNonQuery($"delete from {tableName}");//����ǰ��ձ�(������)�������ظ�
            sqlHelper.ExecuteMutliQuery(commandText, dt);
        }
        catch (Exception ex)
        {
            MessageBox.Show("�������ݿ����info: " + ex.Message);
        }
    }

    /// <summary>
    /// �������ݿ�
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
                string CertificateValidityPeriod = row["֤����Ч��"].ToString();
                string StandardValidityPeriod = row["��׼��Ч��"].ToString();

                row["֤����ЧԤ��"] = CheckStatus_CertificateValidityPeriod(CertificateValidityPeriod);
                row["��׼Ԥ��"] = CheckStatus_StandardValidityPeriod(StandardValidityPeriod);
                row["֤��Ԥ��"] = CheckStatus_CertificateValidityPeriod2(CertificateValidityPeriod);
            }

            this.ImportDt(dt);
        }
        catch (Exception ex)
        {
            MessageBox.Show("��������: " + ex.Message);
        }
    }
    /// <summary>
    /// ��ѯ���ݿ�ָ��������
    /// </summary>
    /// <returns></returns>
    public DataTable DbPointTable()
    {
        DataTable dt = new DataTable();
        string sql = $"SELECT * FROM Test ORDER BY ID ASC";//��ID��������

        try
        {
            dt = sqlHelper.ExecuteTable(sql);
            //���ݿ�����������
            uiDataGridView1.DataSource = dt;

            if (dt.Rows.Count <= 0)
                return default;
        }
        catch (Exception ex)
        {
            MessageBox.Show("��������: " + ex.Message);
        }
        return dt;
    }

    /// <summary>
    /// �ֶ���ѯ���ݿ�ˢ�½�����ť
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void btnRead_Click_1(object sender, EventArgs e)
    {
        RefreshDataGridView();
    }

    /// <summary>
    /// ����ΪCSV�ļ�
    /// </summary>
    /// <param name="dataGridView"></param>
    /// <param name="filePath"></param>
    public void ExportToCSV(DataGridView dataGridView, string filePath, bool prompt)
    {

        // ����ļ�·���Ƿ�Ϊ��  
        if (string.IsNullOrEmpty(filePath))
        {
            if (prompt) MessageBox.Show("��ָ��һ���ļ�·����", "����", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // ��� DataGridView �Ƿ�������  
        if (uiDataGridView1.Rows.Count == 0)
        {
            if (prompt) MessageBox.Show("DataGridView ��û�����ݡ�", "ע��", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            // ʹ�� StreamWriter ��������ļ�  
            using (StreamWriter sw = new StreamWriter(filePath, false))
            {
                // д���б���  
                for (int i = 0; i < uiDataGridView1.Columns.Count; i++)
                {
                    if (i > 0) sw.Write(",");
                    sw.Write(uiDataGridView1.Columns[i].HeaderText);
                }
                sw.WriteLine();

                // д��������  
                foreach (DataGridViewRow row in uiDataGridView1.Rows)
                {
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        if (i > 0) sw.Write(",");

                        // ����Ƿ�Ϊ��ֵ������д����ַ���  
                        if (row.Cells[i].Value != null)
                        {
                            // ����������š����з��������ַ�����������Ų�ת���ڲ�����  
                            string value = row.Cells[i].Value.ToString().Replace(",", "\",\"")
                                .Replace(Environment.NewLine, "\\n")
                                .Replace("\"", "\"\"");
                            sw.Write("\"" + value + "\"");
                        }
                    }
                    sw.WriteLine();
                }
            }

            if (prompt) MessageBox.Show("�����ɹ���", "���", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            if (prompt) MessageBox.Show("����ʧ��: " + ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

    }

    /// <summary>
    /// ����ΪXLSX�ļ�
    /// </summary>
    /// <param name="dataGridView"></param>
    /// <param name="filePath"></param>
    public void ExportToXLSX(DataGridView dataGridView, string filePath, bool prompt)
    {
        // ����ļ�·���Ƿ�Ϊ��  
        if (string.IsNullOrEmpty(filePath))
        {
            if (prompt) MessageBox.Show("��ָ��һ���ļ�·����", "����", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // ��� DataGridView �Ƿ�������  
        if (dataGridView.Rows.Count == 0)
        {
            if (prompt) MessageBox.Show("DataGridView ��û�����ݡ�", "ע��", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            // ����һ���µ� Excel ������  
            var workbook = new XLWorkbook();
            // ���һ��������  
            var worksheet = workbook.Worksheets.Add("֤����Ч��");

            // д���б���  
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                worksheet.Cell(1, i + 1).Value = dataGridView.Columns[i].HeaderText;
            }

            // д��������  
            for (int row = 0; row < dataGridView.Rows.Count; row++)
            {
                for (int col = 0; col < dataGridView.Columns.Count; col++)
                {
                    // ����Ƿ�Ϊ��ֵ  
                    if (dataGridView.Rows[row].Cells[col].Value != null)
                    {
                        worksheet.Cell(row + 2, col + 1).Value = dataGridView.Rows[row].Cells[col].Value.ToString();
                    }
                }
            }

            // ���湤����  
            workbook.SaveAs(filePath);

            if (prompt) MessageBox.Show("�����ɹ���", "���", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            if (prompt) MessageBox.Show("����ʧ��: " + ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// ����xlsx�ļ���ť
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void btnExport_Click(object sender, EventArgs e)
    {

        //SaveFileDialog saveFileDialog = new SaveFileDialog();
        //saveFileDialog.Filter = "XLSX files (*.xlsx)|*.xlsx";
        //saveFileDialog.Title = "���� xlsx �ļ�";

        //if (saveFileDialog.ShowDialog() == DialogResult.OK)
        //{
        //    string filePath = saveFileDialog.FileName;
        //    ExportToXLSX(uiDataGridView1, filePath, true);
        //}

        //SaveFileDialog saveFileDialog = new SaveFileDialog();
        //saveFileDialog.Filter = "CSV files (*.csv)|*.csv";
        //saveFileDialog.Title = "���� CSV �ļ�";

        //if (saveFileDialog.ShowDialog() == DialogResult.OK)
        //{
        //    string filePath = saveFileDialog.FileName;
        //    ExportToCSV(uiDataGridView1, filePath, true);
        //}

        string OriginalTablePath = Directory.GetCurrentDirectory() + "\\" + "��Ʒ��֤ͳ�Ʊ�_ԭʼ.xlsx"; // ԭʼ��·��
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "XLSX files (*.xlsx)|*.xlsx";

        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            string filePath = saveFileDialog.FileName;
            EPPlusHelpr.ExportToXLSX_BasedOnOriginalTable(uiDataGridView1, OriginalTablePath, filePath, true);
        }

    }

    /// <summary>
    /// ����������ݣ�ͬ�����ݿ�
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void toolStripMenuItem1_Click(object sender, EventArgs e)
    {
        int insertId;
        if (uiDataGridView1.SelectedRows.Count > 0)
        {
            // ��ȡ��ǰѡ���е�����
            int selectedRowIndex = uiDataGridView1.SelectedRows[0].Index;
            DataGridViewRow selectedRow = uiDataGridView1.Rows[selectedRowIndex];
            int selectedRowId = Convert.ToInt32(selectedRow.Cells["ID"].Value);
            insertId = selectedRowId + 1;//ѡ��ID����һ������
        }
        else
        {
            // ��ȡ���ݿ������ID����1
            insertId = GetMaxId() + 1;
        }

        // �����������ݵĶԻ��򣬲������ʼ����
        InsertForm insertForm = new InsertForm();
        insertForm.GetID(insertId.ToString());//Ĭ�Ͻ�Ŀ��ID��ֵ�ı���
        DialogResult result = insertForm.ShowDialog();

        if (result == DialogResult.OK)
        {
            // ��ȡ�û����������
            string[] newData = insertForm.GetInsertedData();

            // �û������ID
            if (int.TryParse(newData[0], out insertId) && insertId >= 1)
            {
                UpdateSubsequentIds(insertId);//����������IDǰ���޸���Ӱ����ID                              
            }
            else
            {
                MessageBox.Show("�����ID��Ч��������һ�����ڻ��ߵ���1��������");
                return;
            }

            // �������ݷ�װ���ֵ�
            Dictionary<string, string> insertData = new Dictionary<string, string>
            {
                { "ID",             newData[0] },
                { "֤����",       newData[1] },
                { "��֤����",       newData[2] },
                { "��֤�����׼", newData[3] },
                { "��Ʒϵ��",       newData[4] },
                { "��֤����",       newData[5] },
                { "֤����Ч��",     newData[6] },
                { "��׼��Ч��",     newData[7] },
                { "����������",     newData[8] },
                { "֤����ЧԤ��",   CheckStatus_CertificateValidityPeriod(newData[6])},
                { "��׼Ԥ��",       CheckStatus_StandardValidityPeriod(newData[7])},
                { "֤��Ԥ��",       CheckStatus_CertificateValidityPeriod2(newData[6]) },
                { "˵��",           newData[12] },
            };
            // �������ݵ����ݿ�
            sqlHelper.ExecuteInsert("Test", insertData);
            //ˢ�½�����
            RefreshDataGridView();
        }
    }

    /// <summary>
    /// ����ǰ�Ƚ����ڵ���Ŀ��ID������ID+1
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
    /// �����ݿ��л�ȡ���ID
    /// </summary>
    /// <returns></returns>
    private int GetMaxId()
    {
        string query = "SELECT MAX(ID) FROM Test";
        object result = sqlHelper.ExecuteScalar(query);
        return result == DBNull.Value ? 0 : Convert.ToInt32(result);
    }

    /// <summary>
    /// ɾ��������ݣ�ͬ�����ݿ�(֧��ɾ�����������������ID��ɾ����ʹID������������)
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void toolStripMenuItem2_Click(object sender, EventArgs e)
    {
        // ȷ�����б�ѡ��
        if (uiDataGridView1.SelectedRows.Count > 0)
        {
            // ��¼Ҫɾ����ID
            List<string> idsToDelete = new List<string>();

            foreach (DataGridViewRow selectedRow in uiDataGridView1.SelectedRows)
            {
                string id = selectedRow.Cells["ID"].Value.ToString();
                idsToDelete.Add(id);
            }

            if (idsToDelete.Count > 0)
            {
                // ����һ��ȷ��ɾ��������Ϣ
                string message = $"��ȷ��Ҫɾ������ID��������Ϣ��\n{string.Join("\n", idsToDelete)}";
                DialogResult result = MessageBox.Show(message, "ȷ��ɾ��", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    // ִ��ɾ������
                    foreach (string id in idsToDelete)
                    {
                        sqlHelper.ExecuteDelete("Test", $"ID='{id}'");
                    }
                    // �������ݿ�ID(����������)
                    UpdateIDsSequential();

                    //ˢ�½�����
                    RefreshDataGridView();
                    MessageBox.Show("ɾ���ɹ���");
                }
            }
        }
        else
        {
            MessageBox.Show("���������հ��У�ѡ����Ҫɾ�����У�");
        }
    }

    /// <summary>
    /// �������ݿ�ID(����������)
    /// </summary>
    private void UpdateIDsSequential()
    {
        // �����ݿ��л�ȡ��������

        var dataTable = sqlHelper.ExecuteTable("SELECT * FROM Test ORDER BY ID");

        // ����IDʹ������
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
    /// �޸ı�����ݣ�ͬ�����ݿ�
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void toolStripMenuItem3_Click(object sender, EventArgs e)
    {

        // ��ȡ��ǰѡ���е�����
        if (uiDataGridView1.SelectedRows.Count > 0)
        {
            DataGridViewRow selectedRow = uiDataGridView1.SelectedRows[0];
            string[] originalData = new string[]
            {
            selectedRow.Cells["ID"].Value.ToString(),
            selectedRow.Cells["֤����"].Value.ToString(),
            selectedRow.Cells["��֤����"].Value.ToString(),
            selectedRow.Cells["��֤�����׼"].Value.ToString(),
            selectedRow.Cells["��Ʒϵ��"].Value.ToString(),
            selectedRow.Cells["��֤����"].Value.ToString(),
            selectedRow.Cells["֤����Ч��"].Value.ToString(),
            selectedRow.Cells["��׼��Ч��"].Value.ToString(),
            selectedRow.Cells["����������"].Value.ToString(),
            selectedRow.Cells["֤����ЧԤ��"].Value.ToString(),
            selectedRow.Cells["��׼Ԥ��"].Value.ToString(),
            selectedRow.Cells["֤��Ԥ��"].Value.ToString(),
            selectedRow.Cells["˵��"].Value.ToString()
            };

            // �����޸����ݵĶԻ���         
            InsertForm insertForm = new InsertForm();
            insertForm.GetSelectedData(originalData);
            DialogResult result = insertForm.ShowDialog();

            if (result == DialogResult.OK)
            {
                // ��ȡ�û��޸ĺ������
                string[] editedData = insertForm.GetInsertedData();
                Dictionary<string, string> updateData = new Dictionary<string, string>
                {
                    { "ID",             editedData[0] },
                    { "֤����",       editedData[1] },
                    { "��֤����",       editedData[2] },
                    { "��֤�����׼", editedData[3] },
                    { "��Ʒϵ��",       editedData[4] },
                    { "��֤����",       editedData[5] },
                    { "֤����Ч��",     editedData[6] },
                    { "��׼��Ч��",     editedData[7] },
                    { "����������",     editedData[8] },
                    { "֤����ЧԤ��",   CheckStatus_CertificateValidityPeriod(editedData[6])},
                    { "��׼Ԥ��",       CheckStatus_StandardValidityPeriod(editedData[7])},
                    { "֤��Ԥ��",       CheckStatus_CertificateValidityPeriod2(editedData[6]) },
                    { "˵��",           editedData[12] },
                };

                // �������ݵ����ݿ�
                string id = selectedRow.Cells["ID"].Value.ToString();
                sqlHelper.ExecuteUpdate("Test", $"ID={id}", updateData);

                //ˢ�½�����
                RefreshDataGridView();

            }
        }
        else
        {
            MessageBox.Show("���������հ��У�ѡ����Ҫ�޸ĵ��У�");
        }
    }


    /// <summary>
    /// ˢ�½�����
    /// </summary>
    private void RefreshDataGridView()
    {
        //��ѯ���ݿ��
        DbPointTable();
        uiDataGridView1.ContextMenuStrip = uiContextMenuStrip1;

        foreach (DataGridViewRow row in uiDataGridView1.Rows)
        {
            ProcessDateCell_Display("֤����Ч��", row, uiDataGridView1);
            ProcessDateCell_Display("��׼��Ч��", row, uiDataGridView1);
            ProcessDateCell_Display("֤����ЧԤ��", row, uiDataGridView1);
            ProcessDateCell_Display("��׼Ԥ��", row, uiDataGridView1);
            ProcessDateCell_Display("֤��Ԥ��", row, uiDataGridView1);
        }
    }

    /// <summary>
    /// ָ��������Ϊ���ڣ����ݱ�ע��ɫ
    /// </summary>
    /// <param name="columnName"></param>
    /// <param name="row"></param>
    /// <param name="uiDataGridView"></param>
    private void ProcessDateCell_Display(string columnName, DataGridViewRow row, DataGridView uiDataGridView)
    {
        if (row.Cells[columnName].Value != null)
        {
            if (row.Cells[columnName].Value.ToString() == "����")
            {
                row.Cells[columnName].Style.ForeColor = Color.Red;
            }

            if (columnName == "֤��Ԥ��")
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




