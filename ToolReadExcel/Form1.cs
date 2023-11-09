using OfficeOpenXml;
using System.Collections.Generic;
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
            btnExport.Enabled = false;
        }
        private int _count;
        private int _rowCount;
        private int _colCount;
        List<string> listCategoryNo = new List<string>() 
        {
            "100","101","102","103"
        };
        List<string> listSetCode = new List<string>()
        {
            "2302224","2302226","2302293","2302228","2302238","2302240","2302289","2302291","2302285","2302287"
        };
        List<PATTERN_MST> listPATTERN_MST = new List<PATTERN_MST>();
        private void btnSubmit_Click(object sender, EventArgs e)
        {
            dataGridView.Rows.Clear();
            _count = 0;
            _rowCount = 0;
            btnSubmit.Enabled= false;
            btnExport.Enabled = false;
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

                            List<string> listAGCD1= new List<string>();
                            // Đọc dữ liệu từ từng ô trong Sheet
                            for (int row = 2; row <= _rowCount; row++)
                            {
                                // Lấy giá trị từ ô tại vị trí (row, col)
                                var cellValueAGCD1 = worksheet.Cells[row, 1].Value?.ToString();
                                listAGCD1.Add(cellValueAGCD1);  
                            }

                            listPATTERN_MST = new List<PATTERN_MST>();
                            foreach (var category in listCategoryNo)
                            {
                                foreach (var setCode in listSetCode)
                                {
                                    foreach (var agcd1 in listAGCD1)
                                    {
                                        string query = $"select * from PATTERN_MST where CATEGCD = '{category}' and SETCMDCD = '{setCode}' and CUSTCD1 = '{agcd1}'";
                                        // Thực hiện công việc với câu truy vấn ở đây
                                        using (SqlCommand command = new SqlCommand(query, connection))
                                        {
                                            using (SqlDataReader reader = command.ExecuteReader())
                                            {
                                                while (reader.Read())
                                                {
                                                    PATTERN_MST pattern_mst = new PATTERN_MST();
                                                    pattern_mst.CATEGCD = reader["CATEGCD"].ToString();
                                                    pattern_mst.SETCMDCD = reader["SETCMDCD"].ToString();
                                                    pattern_mst.CUSTCD1 = reader["CUSTCD1"].ToString();
                                                    pattern_mst.AGCD = reader["AGCD"].ToString();
                                                    pattern_mst.KORMKS = reader["KORMKS"].ToString();
                                                    listPATTERN_MST.Add(pattern_mst);
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            ShowDataGridView(listPATTERN_MST);
                        }
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
                    btnExport.Enabled = true;
                }
            }
        }
        private void ShowDataGridView(List<PATTERN_MST> listPATTERN_MST)
        {
            foreach(var item in listPATTERN_MST)
            {
                dataGridView.Invoke(new MethodInvoker(delegate ()
                {
                    var row = new DataGridView();
                    dataGridView.Rows.Add(row);
                    dataGridView.Rows[_count].Cells[0].Value = _count + 1;
                    dataGridView.Rows[_count].Cells[1].Value = item.CATEGCD;
                    dataGridView.Rows[_count].Cells[2].Value = item.SETCMDCD;
                    dataGridView.Rows[_count].Cells[3].Value = item.CUSTCD1; ;
                    dataGridView.Rows[_count].Cells[4].Value = item.CUSTCD1; ;
                    dataGridView.Rows[_count].Cells[5].Value = item.KORMKS; ;
                    _count++;
                }));
                //Điều chỉnh thanh Scroll theo vị trí mong muốn
                dataGridView.FirstDisplayedScrollingRowIndex = dataGridView.RowCount - 1;
                //Task.Delay(100).Wait();
            }    
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if(listPATTERN_MST == null || listPATTERN_MST.Count == 0)
            {
                return;
            }    
            try
            {
                string desktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Sample.xlsx");
                var file = new FileInfo(desktopPath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sample Sheet");
                    //Vị trí bắt đầu
                    worksheet.Cells["A1"].LoadFromCollection(listPATTERN_MST, true);
                    // Thiết lập căn giữa cho phần header
                    using (var range = worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns])
                    {
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }
                    // Thiết lập căn lề trái cho phần giá trị
                    using (var range = worksheet.Cells[2, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns])
                    {
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    }
                    package.Save();
                    MessageBox.Show("Export Excel Success", "Success", MessageBoxButtons.OK, MessageBoxIcon.None);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            btnExport.Enabled = false;
        }
    }

    public class PATTERN_MST
    {
        public string CATEGCD { get; set; }
        public string SETCMDCD { get; set; }
        public string CUSTCD1 { get; set; }
        public string AGCD { get; set; }
        public string KORMKS { get; set; }
    }    
}