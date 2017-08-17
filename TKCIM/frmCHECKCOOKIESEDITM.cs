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
    public partial class frmCHECKCOOKIESEDITM : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();

        int result;
        string ID;

        public frmCHECKCOOKIESEDITM()
        {
            InitializeComponent();
        }

        public frmCHECKCOOKIESEDITM(string SUBID)
        {
            InitializeComponent();

            ID = SUBID;
            SERACHCHECKCOOKIESM();

        }

        #region FUNCTION
        public void SERACHCHECKCOOKIESM()
        {

        }

        public void SETVALUES()
        {

        }

        public void UPDATECHECKCOOKIESM()
        {

        }
        #endregion

        #region BUTTON

        #endregion
    }
}
