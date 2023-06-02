using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;

namespace ToolReadExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataSet ds;
        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(pathExcel.Text) || string.IsNullOrEmpty(strConnect.Text))
            {
                MessageBox.Show("Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                string filePath = pathExcel.Text;
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    // Lấy Sheet đầu tiên từ file Excel
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Lấy số lượng hàng và cột trong Sheet
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    // Đọc dữ liệu từ từng ô trong Sheet
                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Lấy giá trị từ ô tại vị trí (row, col)
                        var cellValueCMDCD = worksheet.Cells[row, 1].Value;
                        var cellValue = worksheet.Cells[row, 10].Value;
                        // Xử lý giá trị từ ô tại vị trí (row, col)
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
          

            //string connectionString = strConnect.Text;

            //// Đoạn mã SQL UPDATE
            //string sqlUpdate = "UPDATE TableName SET Column1 = 'NewValue1', Column2 = 'NewValue2' WHERE Condition";

            //// Tạo đối tượng SqlConnection
            //using (SqlConnection connection = new SqlConnection(connectionString))
            //{
            //    // Mở kết nối
            //    connection.Open();

            //    // Tạo đối tượng SqlCommand với câu lệnh SQL và kết nối
            //    using (SqlCommand command = new SqlCommand(sqlUpdate, connection))
            //    {
            //        // Thực thi câu lệnh UPDATE
            //        command.ExecuteNonQuery();
            //    }
            //}
        }
    }
}