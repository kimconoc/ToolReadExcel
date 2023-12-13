using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Threading;
using System.Transactions;
using System.Windows.Forms;
using static System.Formats.Asn1.AsnWriter;
using static ToolReadExcel.Form1;

namespace ToolReadExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region  Phần chung

        private string sqlQueryGetSpreaData = @"
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

        string sqlQuerySETLOCKTIMEOUT = @"
            SET LOCK_TIMEOUT 0
            SELECT CONVERT(VARCHAR(19), UP_DT , 20) 
            FROM ZAIKO_PRO WITH (UPDLOCK) 
            WHERE CRTDT = CONVERT(DATETIME, 'yyyy-MM-dd HH:mm:ss')
        ";

        string sqlQueryGetZAIKOPROByCRTDT = @"
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

        private List<CRTDTExcel> listCRTDTExcel = new List<CRTDTExcel>();
        private List<ZAIKO_PROEQual> listZAIKO_PROEQual = new List<ZAIKO_PROEQual>();
        private List<ZAIKO_PROELess> listZAIKO_PROELess = new List<ZAIKO_PROELess>();

        #endregion

        private void btnExecuteForwardedData_Click(object sender, EventArgs e)
        {
            strConnect.Text = "Data Source=V002345\\MSSQLSERVER01;Initial Catalog=ncpc;User ID=sa;Password=ad1234567@;Connect Timeout=30;Pooling=False;";
            pathExcel.Text = "D:\\VTI_Hoya\\KANRI\\KANRI-96\\QUANNTExecute.xlsx";
            if (string.IsNullOrEmpty(strConnect.Text))
            {
                MessageBox.Show("Empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (TransactionScope scope = new TransactionScope(TransactionScopeOption.Required,new System.TimeSpan(0, 15, 0)))
            {
                try
                {
                    listCRTDTExcel = new List<CRTDTExcel>();
                    listZAIKO_PROEQual = new List<ZAIKO_PROEQual>();
                    listZAIKO_PROELess = new List<ZAIKO_PROELess>();
                    string connectionString = strConnect.Text;
                    string filePath = pathExcel.Text;
                    // Tạo đối tượng SqlConnection
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        // Mở kết nối
                        connection.Open();
                        // Thực hiện công việc với câu truy vấn lấy list CRTDT
                        using (SqlCommand command = new SqlCommand(sqlQueryGetSpreaData, connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    CRTDTExcel cRTDTExcel = new CRTDTExcel();
                                    cRTDTExcel.CRTDT = reader["CRTDT"].ToString();
                                    listCRTDTExcel.Add(cRTDTExcel);
                                }
                            }
                        }
                        // Thực hiện công việc với đọc excel lấy list QUANNTExecute
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                            int rowCount = worksheet.Dimension.Rows;
                            int colCount = worksheet.Dimension.Columns;

                            for (int row = 2; row <= rowCount; row++)
                            {
                                int cellQUANNTExecute = int.Parse(worksheet.Cells[row, 5].Value?.ToString());
                                listCRTDTExcel[row - 2].QUANNTExecute = cellQUANNTExecute;
                            }
                        }

                        foreach (var crtdt in listCRTDTExcel)
                        {
                            string strSqlQueryGetZAIKOPROByCRTDT = sqlQueryGetZAIKOPROByCRTDT.Replace("yyyy-MM-dd HH:mm:ss", crtdt.CRTDT);
                            // Thực hiện công việc với câu truy vấn lấy list ZAIKO_PRO
                            using (SqlCommand command = new SqlCommand(strSqlQueryGetZAIKOPROByCRTDT, connection))
                            {
                                using (SqlDataReader reader = command.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        if(int.Parse(reader["QUANT"].ToString()) == crtdt.QUANNTExecute)
                                        {
                                            ZAIKO_PROEQual zAIKO_PROEQual = new ZAIKO_PROEQual();
                                            zAIKO_PROEQual.INS_PS = "ZAIKO300";
                                            zAIKO_PROEQual.UP_PS = "ZAIKO300";
                                            zAIKO_PROEQual.INS_HST = "ZAIKO300";
                                            zAIKO_PROEQual.UP_HST = "ZAIKO300";

                                            if (reader.GetValue(reader.GetOrdinal("CRTDT")) != DBNull.Value)
                                            {
                                                zAIKO_PROEQual.CRTDT = reader["CRTDT"].ToString();
                                            }
                                            if (reader.GetValue(reader.GetOrdinal("CMDCD")) != DBNull.Value)
                                            {
                                                zAIKO_PROEQual.CMDCD = reader["CMDCD"].ToString();
                                            }
                                            if (reader.GetValue(reader.GetOrdinal("LOTNO")) != DBNull.Value)
                                            {
                                                zAIKO_PROEQual.LOTNO = reader["LOTNO"].ToString();
                                            }

                                            if (reader.GetValue(reader.GetOrdinal("LIFTM")) != DBNull.Value)
                                            {
                                                zAIKO_PROEQual.LIFTM = reader["LIFTM"].ToString();
                                            }
                                            if (reader.GetValue(reader.GetOrdinal("DIRPGNO")) != DBNull.Value)
                                            {
                                                zAIKO_PROEQual.DIRPGNO = reader["DIRPGNO"].ToString();
                                            }
                                            if (reader.GetValue(reader.GetOrdinal("MKDT")) != DBNull.Value)
                                            {
                                                zAIKO_PROEQual.MKDT = reader["MKDT"].ToString();
                                            }
                                            
                                            string mkcntValue = reader["MKCNT"].ToString();
                                            int mkcnt;
                                            if (int.TryParse(mkcntValue, out mkcnt))
                                            {
                                                zAIKO_PROEQual.MKCNT = mkcnt;
                                            }
                                            else
                                            {
                                                zAIKO_PROEQual.MKCNT = null;
                                            }

                                            if (reader.GetValue(reader.GetOrdinal("STOCD")) != DBNull.Value)
                                            {
                                                zAIKO_PROEQual.STOCD = reader["STOCD"].ToString();
                                            }
                                            if (reader.GetValue(reader.GetOrdinal("MKPRC")) != DBNull.Value)
                                            {
                                                zAIKO_PROEQual.MKPRC = reader["MKPRC"].ToString();
                                            }
                                            zAIKO_PROEQual.MODQUA = 0;
                                            zAIKO_PROEQual.MODKND = "2";
                                            zAIKO_PROEQual.MODRSN = "廃棄　EK20009";
                                            if (reader.GetValue(reader.GetOrdinal("COSTCD")) != DBNull.Value)
                                            {
                                                zAIKO_PROEQual.COSTCD = reader["COSTCD"].ToString();
                                            }
                                            zAIKO_PROEQual.STOTP = "C";
                                            if (reader.GetValue(reader.GetOrdinal("APPDT")) != DBNull.Value)
                                            {
                                                zAIKO_PROEQual.APPDT = reader["APPDT"].ToString();
                                            }
                                            zAIKO_PROEQual.PROKND = "1";
                                            zAIKO_PROEQual.QUANNTExecute = crtdt.QUANNTExecute;
                                            listZAIKO_PROEQual.Add(zAIKO_PROEQual);
                                        }    
                                        else if(int.Parse(reader["QUANT"].ToString()) > crtdt.QUANNTExecute)
                                        {
                                            ZAIKO_PROELess zAIKO_PROELess = new ZAIKO_PROELess();
                                            zAIKO_PROELess.INS_PS = "ZAIKO300";
                                            zAIKO_PROELess.UP_PS = "ZAIKO300";
                                            zAIKO_PROELess.INS_HST = "ZAIKO300";
                                            zAIKO_PROELess.UP_HST = "ZAIKO300";

                                            if (reader.GetValue(reader.GetOrdinal("CRTDT")) != DBNull.Value)
                                            {
                                                zAIKO_PROELess.CRTDT = reader["CRTDT"].ToString();
                                            }
                                            if (reader.GetValue(reader.GetOrdinal("CMDCD")) != DBNull.Value)
                                            {
                                                zAIKO_PROELess.CMDCD = reader["CMDCD"].ToString();
                                            }
                                            if (reader.GetValue(reader.GetOrdinal("LOTNO")) != DBNull.Value)
                                            {
                                                zAIKO_PROELess.LOTNO = reader["LOTNO"].ToString();
                                            }

                                            if (reader.GetValue(reader.GetOrdinal("LIFTM")) != DBNull.Value)
                                            {
                                                zAIKO_PROELess.LIFTM = reader["LIFTM"].ToString();
                                            }
                                            if (reader.GetValue(reader.GetOrdinal("DIRPGNO")) != DBNull.Value)
                                            {
                                                zAIKO_PROELess.DIRPGNO = reader["DIRPGNO"].ToString();
                                            }
                                            if (reader.GetValue(reader.GetOrdinal("MKDT")) != DBNull.Value)
                                            {
                                                zAIKO_PROELess.MKDT = reader["MKDT"].ToString();
                                            }

                                            string mkcntValue = reader["MKCNT"].ToString();
                                            int mkcnt;
                                            if (int.TryParse(mkcntValue, out mkcnt))
                                            {
                                                zAIKO_PROELess.MKCNT = mkcnt;
                                            }
                                            else
                                            {
                                                zAIKO_PROELess.MKCNT = null;
                                            }

                                            if (reader.GetValue(reader.GetOrdinal("STOCD")) != DBNull.Value)
                                            {
                                                zAIKO_PROELess.STOCD = reader["STOCD"].ToString();
                                            }
                                            if (reader.GetValue(reader.GetOrdinal("MKPRC")) != DBNull.Value)
                                            {
                                                zAIKO_PROELess.MKPRC = reader["MKPRC"].ToString();
                                            }
                                            zAIKO_PROELess.MODQUA = -1;
                                            zAIKO_PROELess.MODKND = "0";
                                            zAIKO_PROELess.MODRSN = "廃棄　EK20009";
                                            if (reader.GetValue(reader.GetOrdinal("COSTCD")) != DBNull.Value)
                                            {
                                                zAIKO_PROELess.COSTCD = reader["COSTCD"].ToString();
                                            }
                                            zAIKO_PROELess.STOTP = "3";
                                            if (reader.GetValue(reader.GetOrdinal("APPDT")) != DBNull.Value)
                                            {
                                                zAIKO_PROELess.APPDT = reader["APPDT"].ToString();
                                            }
                                            zAIKO_PROELess.PROKND = "0";
                                            zAIKO_PROELess.QUANNTExecute = crtdt.QUANNTExecute;
                                            listZAIKO_PROELess.Add(zAIKO_PROELess);
                                        }    
                                    }
                                }
                            }
                        }

                        if((listZAIKO_PROEQual.Count + listZAIKO_PROELess.Count) != listCRTDTExcel.Count)
                        {
                            connection.Dispose();
                            MessageBox.Show("Số lượng không trùng khớp", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        if (listZAIKO_PROEQual.Count > 0)
                        {
                            int rowsAffected = 0;
                            foreach (var item in listZAIKO_PROEQual)
                            {
                                using (SqlCommand command = new SqlCommand(sqlQueryINSERTZAIKO_PMODEQual, connection))
                                {
                                    command.Parameters.AddWithValue("@CRTDT", item.CRTDT ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@CMDCD", item.CMDCD ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@LOTNO", item.LOTNO ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@LIFTM", item.LIFTM ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@DIRPGNO", item.DIRPGNO ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@MKDT", item.MKDT ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@MKCNT", item.MKCNT ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@STOCD", item.STOCD ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@MODQUA", item.MODQUA ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@MODKND", item.MODKND ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@MODRSN", item.MODRSN ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@STOTP", item.STOTP ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@APPDT", item.APPDT ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@PROKND", item.PROKND ?? (object)DBNull.Value);

                                    rowsAffected = command.ExecuteNonQuery();
                                    if (rowsAffected > 1)
                                        return;
                                }

                                using (SqlCommand command = new SqlCommand(sqlQueryUPDATEZAIKO_PROEQual, connection))
                                {
                                    command.Parameters.AddWithValue("@STOTP", item.STOTP ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@PROKND", item.PROKND ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@CRTDT", item.CRTDT ?? (object)DBNull.Value);

                                    rowsAffected = command.ExecuteNonQuery();
                                    if (rowsAffected > 1)
                                        return;
                                }

                                using (SqlCommand command = new SqlCommand(sqlQueryINSERTHAIKI_PLANQual, connection))
                                {
                                    decimal mathFloor = Math.Floor(Convert.ToDecimal(item.MKPRC) * Convert.ToDecimal(item.QUANNTExecute) + 0.5m);
                                    command.Parameters.AddWithValue("@CRTDT", item.CRTDT ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@CMDCD", item.CMDCD ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@LOTNO", item.LOTNO ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@MKDT", item.MKDT ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@MKCNT", item.MKCNT ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@LIFTM", item.LIFTM ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@DIRPGNO", item.DIRPGNO ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@STOCD", item.STOCD ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@APPDT", item.APPDT ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@STOTP", item.STOTP ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@MKPRC", item.MKPRC ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@HQUANT", item.QUANNTExecute ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@MathFloor", mathFloor);
                                    command.Parameters.AddWithValue("@MODRSN", item.MODRSN ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@COSTCD", item.COSTCD ?? (object)DBNull.Value);

                                    rowsAffected = command.ExecuteNonQuery();
                                    if (rowsAffected > 1)
                                        return;
                                }
                            }
                        }

                        connection.Dispose();

                    }
                    scope.Complete();
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

        #region  Trường hợp bằng nhau

        private string sqlQueryINSERTZAIKO_PMODEQual = @"
                INSERT INTO ZAIKO_PMOD 
                (INS_DT, UP_DT, INS_PS, UP_PS, INS_HST, UP_HST, MDDT, CRTDT, CMDCD, LOTNO, LIFTM, DIRPGNO, MKDT, MKCNT, STOCD, MODQUA, MODKND, MODRSN, STOTP, APPDT, PROKND) 
                VALUES  
                (GETDATE(), GETDATE(), 'ZAIKO340', 'ZAIKO340', 'V000266', 'V000266', GETDATE(), 
                CONVERT(DATETIME, @CRTDT), @CMDCD, @LOTNO, @LIFTM, @DIRPGNO, @MKDT, @MKCNT, @STOCD, @MODQUA, @MODKND, @MODRSN, @STOTP, CONVERT(DATETIME, @APPDT), @PROKND)";

        private string sqlQueryUPDATEZAIKO_PROEQual = @"
                UPDATE ZAIKO_PRO 
                SET 
                    UP_DT = GETDATE(), 
                    UP_PS = 'ZAIKO340', 
                    UP_HST = 'V000266', 
                    STOTP = @STOTP,
                    PROKND = @PROKND
                WHERE 
                    CRTDT = CONVERT(DATETIME, @CRTDT) 
                    AND (
                        UP_HST = 'V000266' 
                        OR (
                            UP_HST <> 'V000266' 
                            AND UP_DT <= CONVERT(DATETIME, '2023-12-07 12:30:59.300')
                        )
                    )";
        // 2023-12-07 thời gian hiện tại
        private string sqlQueryINSERTHAIKI_PLANQual = @"
                INSERT INTO HAIKI_PLAN 
                (INS_DT, UP_DT, INS_PS, UP_PS, INS_HST, UP_HST, HKDT, CRTDT, CMDCD, LOTNO, MKDT, MKCNT, LIFTM, DIRPGNO, STOCD, APPDT, STOTP, PROKND, MKPRC, HQUANT, SUMVAL, HAIKIRSN, WKCD, EMPNO, RELCRTDT, FILE_DT, RINNO, HKEIJYODT) 
                VALUES 
                (GETDATE(), GETDATE(), 'ZAIKO340', 'ZAIKO340', 'V000266', 'V000266', CONVERT(DATETIME, '2023/12/07'), 
                CONVERT(DATETIME, @CRTDT), @CMDCD, @LOTNO, @MKDT, @MKCNT, @LIFTM, @DIRPGNO, @STOCD, 
                CONVERT(DATETIME, @APPDT), @STOTP, '1', @MKPRC, @HQUANT, @MathFloor, @MODRSN, @COSTCD, '99999', 
                NULL, NULL, NULL, NULL)";

        #endregion

        public class CRTDTExcel
        {
            public string CRTDT { get; set; }
            public int QUANNTExecute { get; set; }
        }
        public class ZAIKO_PROEQual
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
            public int? MKCNT { get; set; }
            public string STOCD { get; set; }
            public string MKPRC { get; set; }
            public int? MODQUA { get; set; }
            public string MODKND { get; set; }
            public string MODRSN { get; set; }
            public string COSTCD { get; set; }
            public string STOTP { get; set; }
            public string APPDT { get; set; }
            public string PROKND { get; set; }
            public int? QUANNTExecute { get; set; }
        }
        public class ZAIKO_PROELess
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
            public int? MKCNT { get; set; }
            public string STOCD { get; set; }
            public string MKPRC { get; set; }
            public int? MODQUA { get; set; }
            public string MODKND { get; set; }
            public string MODRSN { get; set; }
            public string COSTCD { get; set; }
            public string STOTP { get; set; }
            public string APPDT { get; set; }
            public string PROKND { get; set; }
            public int? QUANNTExecute { get; set; }
        }
    }
}