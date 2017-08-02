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

namespace TKCIM
{
    public partial class frmDAILYREPORTPACKAGE : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSql2 = new StringBuilder();
        StringBuilder sbSql3 = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter43 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();
        SqlDataAdapter adapter8 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder8 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();
        DataSet ds8 = new DataSet();

        DataSet ds = new DataSet();
        string tablename = null;
        int result;
        string CHECKYN = "N";


        string ID;
        string MDID;
        string TARGETPROTA001;
        string TARGETPROTA002;

        public frmDAILYREPORTPACKAGE()
        {
            InitializeComponent();

            comboBox2load();

        }

        #region FUNCTION
        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD   WHERE MD002 LIKE '新廠包裝線%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MD002";
            comboBox2.DisplayMember = "MD002";
            sqlConn.Close();


        }
        public void SERACHMOCTARGET()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'  ,MB003 AS '規格'  ");
                sbSql.AppendFormat(@"  ,MD002 AS '線別'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA WITH (NOLOCK),[TK].dbo.INVMB WITH (NOLOCK),[TK].dbo.CMSMD WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=  MD001 ");
                //sbSql.AppendFormat(@"  AND MB002 NOT LIKE '%水麵%' ");
                //sbSql.AppendFormat(@"  AND TA006 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'", comboBox2.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY TA003,TA006");
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

            sqlConn.Close();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];                   
                    TARGETPROTA001 = row.Cells["單別"].Value.ToString();
                    TARGETPROTA002 = row.Cells["單號"].Value.ToString();


                }
                else
                {
                    TARGETPROTA001 = null;
                    TARGETPROTA002 = null;

                }
            }
            else
            {
                TARGETPROTA001 = null;
                TARGETPROTA002 = null;
            }

            SEARCHMOCTE();
        }

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
                sbSql.AppendFormat(@"  AND TE004 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TE011='{0}' AND TE012='{1}'",TARGETPROTA001,TARGETPROTA002);
                sbSql.AppendFormat(@"  GROUP BY TE004,MB002,MB003");
                sbSql.AppendFormat(@"  ORDER BY TE004 ");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;

                    SETNULL();
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        SETNULL();
                        int i = 0;
                        foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                        {
                            if (i <= 5)
                            {
                                TextBox iTextBox1 = (TextBox)FindControl(this, "textBox2" + i + "1");
                                iTextBox1.Text = dr.Cells["品號"].Value.ToString();

                                TextBox iTextBox2 = (TextBox)FindControl(this, "textBox2" + i+"2");
                                iTextBox2.Text = dr.Cells["品名"].Value.ToString();

                                TextBox iTextBox3 = (TextBox)FindControl(this, "textBox2" + i + "3");
                                iTextBox3.Text = "0";

                                TextBox iTextBox4 = (TextBox)FindControl(this, "textBox2" + i + "4");
                                iTextBox4.Text = dr.Cells["數量"].Value.ToString();

                                TextBox iTextBox5 = (TextBox)FindControl(this, "textBox2" + i + "5");
                                iTextBox5.Text = "0";

                                TextBox iTextBox6 = (TextBox)FindControl(this, "textBox2" + i + "6");
                                iTextBox6.Text = "0";

                                TextBox iTextBox7 = (TextBox)FindControl(this, "textBox2" + i + "7");
                                iTextBox7.Text = "0";

                                TextBox iTextBox8 = (TextBox)FindControl(this, "textBox2" + i + "8");
                                iTextBox8.Text = "0";

                                TextBox iTextBox9 = (TextBox)FindControl(this, "textBox2" + i + "9");
                                iTextBox9.Text = "0";
                                i++;
                            }

                        }

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

        /// <summary>
        ///查找指定控件容器中，指定名字的控件
        /// </summary>
        /// <param name="i_form">控件容器对象</param>
        /// <param name="i_name">控件名称</param>
        /// <returns>Control对象，需要强制转换回相应的控件(lable)FindControl()</returns>
        public static Control FindControl(Control i_form, string i_name)
        {

            if (i_form.Name.ToString() == i_name.ToString()) return i_form;

            foreach (Control iCtrl in i_form.Controls)//遍历Panel上的所有控件
            {
                Control i_Ctrl = FindControl(iCtrl, i_name);
                if (i_Ctrl != null) return i_Ctrl;

            }
            return null;

        }

        public void SETNULL()
        {
            int i = 0;
            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                if (i <= 5)
                {
                    TextBox iTextBox1 = (TextBox)FindControl(this, "textBox2" + i + "1");
                    iTextBox1.Text =null;

                    TextBox iTextBox2 = (TextBox)FindControl(this, "textBox2" + i + "2");
                    iTextBox2.Text = null;

                    TextBox iTextBox3 = (TextBox)FindControl(this, "textBox2" + i + "3");
                    iTextBox3.Text = null;

                    TextBox iTextBox4 = (TextBox)FindControl(this, "textBox2" + i + "4");
                    iTextBox4.Text = null;

                    TextBox iTextBox5 = (TextBox)FindControl(this, "textBox2" + i + "5");
                    iTextBox5.Text = null;

                    TextBox iTextBox6 = (TextBox)FindControl(this, "textBox2" + i + "6");
                    iTextBox6.Text = null;

                    TextBox iTextBox7 = (TextBox)FindControl(this, "textBox2" + i + "7");
                    iTextBox7.Text = null;

                    TextBox iTextBox8 = (TextBox)FindControl(this, "textBox2" + i + "8");
                    iTextBox8.Text = null;

                    TextBox iTextBox9 = (TextBox)FindControl(this, "textBox2" + i + "9");
                    iTextBox9.Text = null;
                    i++;
                }

            }
        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();
        }

        #endregion

       
    }
}
