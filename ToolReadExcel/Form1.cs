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
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.MediumAquamarine;
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.AllowUserToResizeColumns = false;
            //
            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.MediumAquamarine;
            dataGridView2.EnableHeadersVisualStyles = false;
            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView2.AllowUserToResizeRows = false;
            dataGridView2.AllowUserToResizeColumns = false;
        }
        private string sqlQuery = @"
            SELECT          
            CONVERT(VARCHAR(19), A.CRTDT, 20) as CRTDT,         
            A.CMDCD,         
            A.LOTNO,         
            CONVERT(VARCHAR(10), A.LIFTM, 111),         
            CONVERT(VARCHAR(10), A.MKDT, 111),         
            A.MKCNT,         
            A.QUANT,         
            ISNULL(A.PREQUA, 0) PREQUA,         
            A.STOTP,         
            A.STOCD,         
            C.CMDNMK,         
            B.PROCNM,         
            CONVERT(VARCHAR(19), A.UP_DT, 20),        
            A.APPDT,         
            A.PROKND,         
            A.DIRPGNO,         
            CASE ISNULL(D.EXIST, 0) WHEN 0 THEN 0 ELSE 1 END HAIKI_EXIST,         
            CASE ISNULL(E.EXIST, 0) WHEN 0 THEN 0 ELSE 1 END HAIKI_TARGET,         
            ISNULL(F.SETNO, '') AS SETNO,         
            ISNULL(F.SQUAT, 0) AS SETQUA,         
            ISNULL(F.JQUA, 0) AS SETJQUA,         
            ISNULL(G.SETNAME, '') AS OISETNAME,         
            ISNULL(G.SETNO, '') AS OISETNO  
            FROM         
            (ZAIKO_PRO A   
            INNER JOIN        
            V_PROC_MST B     
            ON  A.STOCD = B.PROCCD)    
            INNER JOIN        
            HSYOHIN_MST C    
            ON  A.CMDCD = C.CMDCD    
            LEFT OUTER JOIN 
            (   
                SELECT RELCRTDT, COUNT(*) EXIST FROM HAIKI_PLAN WHERE RINNO IS NULL GROUP BY RELCRTDT    
            ) D 
            ON D.RELCRTDT = A.CRTDT    
            LEFT OUTER JOIN 
            (   
                SELECT CRTDT, COUNT(*) EXIST FROM HAIKI_PLAN WHERE RINNO IS NULL GROUP BY CRTDT    
            ) E 
            ON E.CRTDT = A.CRTDT   
            LEFT OUTER JOIN 
            ( 
                SELECT FF.SETNO, FF.CRTDT, FF.SQUAT, FF.JQUA, FF.SHPDT FROM 
                (                        
                    SELECT FA.SETNO, FA.CRTDT, FA.QUANT AS SQUAT, FB.QUANT AS JQUA, FC.SHPDT  
                    FROM ZAIKOSETSUB_MST FA                       
                    LEFT OUTER JOIN 
                    (
                        SELECT FBA.ORDNO, FBB.CRTDT, FBB.SETNO, FBB.QUANT 
                        FROM JUCHU_INF FBA , JUCHU_DET FBB 
                        WHERE FBA.ORDNO = FBB.ORDNO AND FBA.DELSTA = '0'
                    ) FB 
                    ON FB.CRTDT = FA.CRTDT AND FB.SETNO = FA.SETNO                       
                    LEFT OUTER JOIN SYUKKA_INF FC 
                    ON FB.ORDNO = FC.ORDNO AND FC.DELSTA = '0' 
                ) FF 
                WHERE FF.SHPDT IS NULL 
            ) F 
            ON F.CRTDT = A.CRTDT  
            LEFT OUTER JOIN NUSETSUB_MST G 
            ON G.CRTDT = A.CRTDT    
            WHERE (A.CMDCD IN ('2080030','2080031','2080032','2080033', '2080034','2080035','2080036','2080037','2080038','2080039','2080042','2080043','2080044','2080048','2080049','2080050','2080053','2080056','2080064','2081044'))
            AND A.STOCD IN ('0012')  
            ORDER BY A.CMDCD, A.MKDT, A.LOTNO, A.STOCD, A.CRTDT
        ";

        string sqlQuery2 = @"
            SELECT 
                CONVERT(VARCHAR(10), A.CRTDT, 111) + ' ' + CONVERT(VARCHAR(8), A.CRTDT, 8) CRTDT, 
                A.CMDCD, C.CMDNMK, A.DIRPGNO, A.LOTNO, 
                CONVERT(VARCHAR(10), A.LIFTM, 111) LIFTM, CONVERT(VARCHAR(10), A.MKDT, 111) MKDT, 
                A.MKCNT, ISNULL(A.QUANT, 0) QUANT, ISNULL(A.PREQUA, 0) PREQUA, A.STOTP, 
                F.DETAIL STOTPNM, A.STOCD, B.PROCNM, A.PROKND, G.DETAIL PROKNDNM, 
                CONVERT(VARCHAR(10), A.APPDT, 111) APPDT, D.MKPRC, E.COSTCD, E.EMPNO, A.JANFLG, A.JANCD 
            FROM 
                ZAIKO_PRO A 
                INNER JOIN PROC_MST B ON B.PROCCD = A.STOCD 
                INNER JOIN HSYOHIN_MST C ON C.CMDCD = A.CMDCD 
                INNER JOIN LSYOHIN_MST D ON D.CMDCD = A.CMDCD 
                INNER JOIN SYAIN_MST E ON E.EMPNO = '99999' 
                LEFT OUTER JOIN GENERIC_CD_MST F ON F.TBLNAME = 'ZAIKO_PRO' AND F.CATEGORY = 'STOTP' AND F.DETAIL_CD = A.STOTP 
                LEFT OUTER JOIN GENERIC_CD_MST G ON G.TBLNAME = 'ZAIKO_PRO' AND G.CATEGORY = 'PROKND' AND G.DETAIL_CD = A.PROKND 
            WHERE 
                A.CRTDT = CONVERT (DATETIME, 'yyyy-MM-dd HH:mm:ss') 
        ";

        private List<string> listCRTDT = new List<string>();
        private List<ZAIKO_PRO> listZAIKO_PRO = new List<ZAIKO_PRO>();

        // Sử dụng chuỗi sqlQuery trong câu lệnh truy vấn đến cơ sở dữ liệu.

        private void btnSelectForwardedData_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            strConnect.Text = "Data Source=V002345\\MSSQLSERVER01;Initial Catalog=ncpc;User ID=sa;Password=ad1234567@;Connect Timeout=30;Pooling=False;";
            if (string.IsNullOrEmpty(strConnect.Text))
            {
                MessageBox.Show("Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (TransactionScope scope = new TransactionScope(TransactionScopeOption.Required,new System.TimeSpan(0, 15, 0)))
            {
                try
                {
                    listCRTDT = new List<string>();
                    listZAIKO_PRO = new List<ZAIKO_PRO>();
                    string connectionString = strConnect.Text;
                    string filePath = pathExcel.Text;
                    // Tạo đối tượng SqlConnection
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        // Mở kết nối
                        connection.Open();
                        // Thực hiện công việc với câu truy vấn lấy list CRTDT
                        using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    listCRTDT.Add(reader["CRTDT"].ToString());
                                }
                            }
                        }

                        foreach(string crtdt in listCRTDT)
                        {
                            string strSqlQuery = sqlQuery2.Replace("yyyy-MM-dd HH:mm:ss", crtdt);
                            // Thực hiện công việc với câu truy vấn lấy list ZAIKO_PRO
                            using (SqlCommand command = new SqlCommand(strSqlQuery, connection))
                            {
                                using (SqlDataReader reader = command.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        ZAIKO_PRO zAIKO_PRO = new ZAIKO_PRO();
                                        zAIKO_PRO.INS_PS = "ZAIKO300";
                                        zAIKO_PRO.UP_PS = "ZAIKO300";
                                        zAIKO_PRO.INS_HST = "ZAIKO300";
                                        zAIKO_PRO.UP_HST = "ZAIKO300";
                                        zAIKO_PRO.CRTDT = reader["CRTDT"].ToString();
                                        zAIKO_PRO.CMDCD = reader["CMDCD"].ToString();
                                        zAIKO_PRO.LOTNO = reader["LOTNO"].ToString();
                                        zAIKO_PRO.LIFTM = reader["LIFTM"].ToString();
                                        zAIKO_PRO.DIRPGNO = reader["DIRPGNO"].ToString();
                                        zAIKO_PRO.MKDT = reader["MKDT"].ToString();
                                        zAIKO_PRO.MKCNT = reader["MKCNT"].ToString();
                                        zAIKO_PRO.STOCD = reader["STOCD"].ToString();
                                        zAIKO_PRO.MODQUA = 0; // -quanty
                                        zAIKO_PRO.MODKND = "2"; // 0
                                        zAIKO_PRO.MODRSN = "廃棄　EK20009";
                                        zAIKO_PRO.STOTP = "C";
                                        zAIKO_PRO.APPDT = reader["APPDT"].ToString();
                                        zAIKO_PRO.PROKND = "1"; // reader["PROKND"].ToString();
                                        listZAIKO_PRO.Add(zAIKO_PRO);
                                    }
                                }
                            }
                        }
                        ShowDataGridView1(listZAIKO_PRO);
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
                }
            }
        }
        private void ShowDataGridView1(List<ZAIKO_PRO> listZAIKO_PRO)
        {
            if (listZAIKO_PRO == null || listZAIKO_PRO.Count == 0)
                return;

            int count = 0;
            foreach (var item in listZAIKO_PRO)
            {
                dataGridView1.Invoke(new MethodInvoker(delegate ()
                {
                    var row = new DataGridView();
                    dataGridView1.Rows.Add(row);
                    dataGridView1.Rows[count].Cells[0].Value = count + 1;
                    dataGridView1.Rows[count].Cells[1].Value = item.CRTDT;
                    dataGridView1.Rows[count].Cells[2].Value = item.CMDCD;
                    dataGridView1.Rows[count].Cells[3].Value = item.LOTNO; 
                    dataGridView1.Rows[count].Cells[4].Value = item.LIFTM; 
                    dataGridView1.Rows[count].Cells[5].Value = item.DIRPGNO; 
                    dataGridView1.Rows[count].Cells[6].Value = item.MKDT;
                    dataGridView1.Rows[count].Cells[7].Value = item.MKCNT;
                    dataGridView1.Rows[count].Cells[8].Value = item.STOCD;
                    dataGridView1.Rows[count].Cells[9].Value = item.MODQUA; 
                    dataGridView1.Rows[count].Cells[10].Value = item.MODKND; 
                    dataGridView1.Rows[count].Cells[11].Value = item.MODRSN;
                    dataGridView1.Rows[count].Cells[12].Value = item.STOTP;
                    dataGridView1.Rows[count].Cells[13].Value = item.APPDT;
                    dataGridView1.Rows[count].Cells[14].Value = item.PROKND;
                    count++;
                }));
                //Điều chỉnh thanh Scroll theo vị trí mong muốn
                dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
                //Task.Delay(100).Wait();
            }
        }

        public class ZAIKO_PRO
        {
            public string INS_PS { get; set; }
            public string UP_PS { get; set; }
            public string INS_HST { get; set; }
            public string UP_HST { get; set; }
            public string CRTDT { get; set; }
            public string CMDCD { get; set; }
            public string LOTNO { get; set; }
            public string LIFTM { get; set; }
            public string DIRPGNO { get; set; }
            public string MKDT { get; set; }
            public string MKCNT { get; set; }
            public string STOCD { get; set; }
            public int MODQUA { get; set; }
            public string MODKND { get; set; }
            public string MODRSN { get; set; }
            public string STOTP { get; set; }
            public string APPDT { get; set; }
            public string PROKND { get; set; }
        }
    }
}