using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using System.Globalization;

namespace TKCIM
{
    public partial class frmDAILYREPORTPACKAGESUB : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();


        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        string MB001;
        string MB002;
        string TARGETPROTA001;
        string TARGETPROTA002;

        public frmDAILYREPORTPACKAGESUB()
        {
            InitializeComponent();
        }

        public frmDAILYREPORTPACKAGESUB(string SUBTARGETPROTA001,string SUBTARGETPROTA002)
        {
            InitializeComponent();

            TARGETPROTA001 = SUBTARGETPROTA001;
            TARGETPROTA002 = SUBTARGETPROTA002;

            SEARCHMOCTE();
        }


        #region FUNCTION

        public void SEARCHMOCTE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TE004 AS '品號',MB002 AS '品名',MB003 AS '規格',SUM(TE005) AS '數量'   ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTE,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE TE004=MB001");                
                sbSql.AppendFormat(@"  AND TE011='{0}' AND TE012='{1}'", TARGETPROTA001, TARGETPROTA002);
                sbSql.AppendFormat(@"  GROUP BY TE004,MB002,MB003");
                sbSql.AppendFormat(@"  ORDER BY TE004 ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                    
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];
                       

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }


        }
       

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    MB002 = row.Cells["品名"].Value.ToString();     
                    textBox1.Text= row.Cells["品名"].Value.ToString();

                }
                else
                {
                    MB002 = null;
                    textBox1.Text = null;

                }
            }
            else
            {
                MB002 = null;
                textBox1.Text = null;
            }

        }

        public string TextBoxMsg
        {
            set
            {

            }
            get
            {
                return MB002;
            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion

       
    }
}
