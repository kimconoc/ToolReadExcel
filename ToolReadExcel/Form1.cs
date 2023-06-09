﻿using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Transactions;
using System.Windows.Forms;
using static System.Formats.Asn1.AsnWriter;

namespace ToolReadExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.MediumAquamarine;
            dataGridView.EnableHeadersVisualStyles = false;
            dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView.AllowUserToResizeRows = false;
            dataGridView.AllowUserToResizeColumns = false;
        }
        private int _count;
        private int _rowCount;
        private int _colCount;
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView.Rows.Clear();
            _count = 0;
            _rowCount = 0;
            btnSubmit.Enabled= false;
            if (string.IsNullOrEmpty(pathExcel.Text) || string.IsNullOrEmpty(strConnect.Text))
            {
                MessageBox.Show("Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (TransactionScope scope = new TransactionScope(TransactionScopeOption.Required,new System.TimeSpan(0, 5, 0)))
            {
                try
                {
                    string connectionString = strConnect.Text;
                    string filePath = pathExcel.Text;
                    // Tạo đối tượng SqlConnection
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        // Mở kết nối
                        connection.Open();
                        // ReadFileExcel
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                        {
                            // Lấy Sheet đầu tiên từ file Excel
                            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                            // Lấy số lượng hàng và cột trong Sheet
                            _rowCount = worksheet.Dimension.Rows;
                            _colCount = worksheet.Dimension.Columns;

                            // Đọc dữ liệu từ từng ô trong Sheet
                            for (int row = 2; row <= _rowCount; row++)
                            {
                                // Lấy giá trị từ ô tại vị trí (row, col)
                                var cellValueCMDCD = worksheet.Cells[row, 1].Value?.ToString();
                                var cellValueMdkikaku = worksheet.Cells[row, 10].Value?.ToString();
                                // Đoạn mã SQL UPDATE
                                string sqlUpdate = string.Format("update LSYOHIN_MST set UP_DT = GETDATE(), MDKIKAKU = '{0}' where CMDCD = '{1}'", cellValueMdkikaku, cellValueCMDCD);
                                // Tạo đối tượng SqlCommand với câu lệnh SQL và kết nối
                                using (SqlCommand command = new SqlCommand(sqlUpdate, connection))
                                {
                                    // Thực thi câu lệnh UPDATE
                                    command.ExecuteNonQuery();
                                }
                                ShowDataGridView(cellValueCMDCD, cellValueMdkikaku);
                            }
                        }
                    }
                    if (_count == _rowCount - 1)
                    {
                        scope.Complete();
                        MessageBox.Show("Update Success!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Update Error!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    // Giải phóng giao dịch 
                    scope.Dispose();
                    btnSubmit.Enabled = true;
                }
            }
        }
        private void ShowDataGridView(string? cellValueCMDCD,string? cellValueMdkikaku)
        {
            dataGridView.Invoke(new MethodInvoker(delegate ()
            {
                var row = new DataGridView();
                dataGridView.Rows.Add(row);
                dataGridView.Rows[_count].Cells[0].Value = _count + 1;
                dataGridView.Rows[_count].Cells[1].Value = cellValueCMDCD;
                dataGridView.Rows[_count].Cells[2].Value = cellValueMdkikaku;
                dataGridView.Rows[_count].Cells[3].Value = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt");
                _count++;
            }));
            //Điều chỉnh thanh Scroll theo vị trí mong muốn
            dataGridView.FirstDisplayedScrollingRowIndex = dataGridView.RowCount - 1;
            Task.Delay(100).Wait();
        }
    }
}