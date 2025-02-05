using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace InfoChange
{
    public partial class frmLogList : Form
    {
        private DataTable dtLogsList;
        private string txtAgreementCode;
        private string constrGIBDD;
        private string txtBankFlag;
        private decimal nResult = 0;
        private OleDbConnection con;


        public decimal ShowForm()
        {

            this.ShowDialog();
            return nResult;
        }

        public frmLogList()
        {
            InitializeComponent();
        }

        // ���� ����������� � ����������, ����� ������� �� ����� ������ ����� ����������
        public frmLogList(string txtAgreementCodeParam, string constrGIBDDParam)
        {
            txtAgreementCode = txtAgreementCodeParam;
            constrGIBDD = constrGIBDDParam;
            con = null;

            InitializeComponent();
        }
        
        // ���� ����������� � ����������, ����� ������� �� ����� ������ ����� ����������
        public frmLogList(OleDbConnection conParam, string txtAgreementCodeParam, string constrGIBDDParam)
        {
            txtAgreementCode = txtAgreementCodeParam;
            constrGIBDD = constrGIBDDParam;
            con = conParam;

            InitializeComponent();
        }
        

        public frmLogList(string txtAgreementCodeParam, string constrGIBDDParam, string txtBankFlagParam)
        {
            txtAgreementCode = txtAgreementCodeParam;
            constrGIBDD = constrGIBDDParam;
            txtBankFlag = txtBankFlagParam;
            con = null;

            InitializeComponent();
        }

        public frmLogList(OleDbConnection conParam, string txtAgreementCodeParam, string constrGIBDDParam, string txtBankFlagParam)
        {
            txtAgreementCode = txtAgreementCodeParam;
            constrGIBDD = constrGIBDDParam;
            txtBankFlag = txtBankFlagParam;
            con = conParam;

            InitializeComponent();
        }


        private DataTable SelectLogs()
        {
            string txtSql = "";
            DataTable dtRes = null;
            try
            {
                
                switch(txtAgreementCode){
                // ���� ��� �� ������� DBF
                    case "100":
                        // ���� ��������� �������� - ��� ��� ����� 2 ������ ������� (�� ��������� � �������������)
                        // ������� ����� ��������������� ������ ������� - ����� �� ���� ������
                        // PACK_STATUS = 12 ��� ��������� ����� �������
                        // PACK_STATUS = 2 ��� ������ ���������
                        txtSql = "SELECT ll.ID, ll.PACKDATE, ll.PACK_COUNT, ls.STATUS_NAME, pt.TYPE from LOCAL_LOGS ll join logs_status ls on ll.pack_status = ls.id join pack_type pt on ll.pack_type = pt.id WHERE ll.pack_type = 1 and ((PACK_STATUS = 2) or (PACK_STATUS = 12)) and CONV_CODE = '" + txtAgreementCode + "'";
                        dtRes = GetDataTableFromFB(constrGIBDD, txtSql, "Logs", IsolationLevel.Unspecified);
                        break;

                    // ���� ��� �� ����������� DBF
                    case "110":
                        txtSql = "SELECT ll.ID, ll.PACKDATE, ll.PACK_COUNT, ls.STATUS_NAME, pt.TYPE from LOCAL_LOGS ll join logs_status ls on ll.pack_status = ls.id join pack_type pt on ll.pack_type = pt.id WHERE ll.pack_type = 1 and PACK_STATUS = 2 and CONV_CODE = '" + txtAgreementCode + "'";
                        dtRes = GetDataTableFromFB(constrGIBDD, txtSql, "Logs", IsolationLevel.Unspecified);
                        break;

                    // ���� ��� ����
                    case "120":
                        txtSql = "SELECT ll.ID, ll.PACKDATE, ll.PACK_COUNT, ls.STATUS_NAME, pt.TYPE from LOCAL_LOGS ll join logs_status ls on ll.pack_status = ls.id join pack_type pt on ll.pack_type = pt.id WHERE ll.pack_type = 1 and PACK_STATUS = 2 and CONV_CODE = '" + txtAgreementCode + "'";
                        dtRes = GetDataTableFromFB(constrGIBDD, txtSql, "Logs", IsolationLevel.Unspecified);
                        break;
                    // ������ ��� �����
                    default:
                        // ������ ���� 3 ����� - FL_FIND, FL_NOFIND, FL_E_TOFIND
                        txtSql = "SELECT ll.ID, ll.PACKDATE, ll.PACK_COUNT, ls.STATUS_NAME, pt.TYPE from LOCAL_LOGS ll join logs_status ls on ll.pack_status = ls.id join pack_type pt on ll.pack_type = pt.id WHERE ll.pack_type = 1 and CONV_CODE = '" + txtAgreementCode + "'";
                        switch (txtBankFlag)
                        {
                            case "FL_FIND":
                                 txtSql += " and FL_FIND = 0";
                                 break;
                            case "FL_NOFIND":
                                txtSql += " and FL_NOFIND = 0";
                                break;
                            case "FL_E_TOFIND":
                                txtSql += " and FL_E_TOFIND = 0";
                                break;
                        }
                        
                        dtRes = GetDataTableFromFB(constrGIBDD, txtSql, "Logs", IsolationLevel.Unspecified);
                        break;
            }
            }
            catch
            {
            }
            return dtRes;
        }

        private string GetLegalNameByAgrCode(OleDbConnection con, string txtAgreementCode)
        {
            String res = "��� �������� � ���� ������";
            try
            {
                if (con != null && con.State.Equals(ConnectionState.Closed)) con.Open();

                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select first 1 AGENT_DEPT FROM MVV_AGENT_AGREEMENT WHERE AGREEMENT_CODE = '" + txtAgreementCode + "'", con, tran);
                res = Convert.ToString(cmd.ExecuteScalar());
                tran.Rollback();
                con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }
            return res;
        }


        private string GetAgrNameByAgrCode(OleDbConnection con, string txtAgreementCode)
        {
            String res = "��� �������� � ���� ������";
            try
            {
                if (con != null && con.State.Equals(ConnectionState.Closed)) con.Open();

                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select first 1 NAME_AGREEMENT FROM MVV_AGENT_AGREEMENT WHERE AGREEMENT_CODE = '" + txtAgreementCode + "'", con, tran);
                res = Convert.ToString(cmd.ExecuteScalar());
                tran.Rollback();
                con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }
            return res;
        }


        private void frmLogList_Load(object sender, EventArgs e)
        {
            dtLogsList = SelectLogs();

            dgvwLogList.DataSource = dtLogsList;

            dgvwLogList.Columns["ID"].Visible = false;

            dgvwLogList.Columns["TYPE"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvwLogList.Columns["TYPE"].FillWeight = 30;
            dgvwLogList.Columns["TYPE"].DisplayIndex = 1;
            dgvwLogList.Columns["TYPE"].HeaderText = "��� ������";

            dgvwLogList.Columns["PACKDATE"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvwLogList.Columns["PACKDATE"].FillWeight = 30;
            dgvwLogList.Columns["PACKDATE"].DisplayIndex = 2;
            dgvwLogList.Columns["PACKDATE"].HeaderText = "���� ������";

            dgvwLogList.Columns["PACK_COUNT"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvwLogList.Columns["PACK_COUNT"].FillWeight = 20;
            dgvwLogList.Columns["PACK_COUNT"].DisplayIndex = 3;
            dgvwLogList.Columns["PACK_COUNT"].HeaderText = "����������";

            dgvwLogList.Columns["STATUS_NAME"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvwLogList.Columns["STATUS_NAME"].FillWeight = 20;
            dgvwLogList.Columns["STATUS_NAME"].DisplayIndex = 4;
            dgvwLogList.Columns["STATUS_NAME"].HeaderText = "������";

            if(con != null){
                lblTitle.Text = GetLegalNameByAgrCode(con, txtAgreementCode);
                lblTitleAgr.Text = "����������: " + GetAgrNameByAgrCode(con, txtAgreementCode);
            }



        }

        private DataTable GetDataTableFromFB(string conStr, string txtSql, string tblName, IsolationLevel islLevel)
        {
            OleDbConnection conGIBDD;
            DataSet ds = new DataSet();
            DataTable tbl = ds.Tables.Add(tblName);
            try
            {
                conGIBDD = new OleDbConnection(conStr);
                conGIBDD.Open();

                OleDbTransaction tran = conGIBDD.BeginTransaction(islLevel);
                OleDbCommand cmd = new OleDbCommand(txtSql, conGIBDD, tran);

                using (OleDbDataReader rdr = cmd.ExecuteReader(CommandBehavior.Default))
                {
                    ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                    rdr.Close();
                }

                tran.Rollback();
                conGIBDD.Close();

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }
            return tbl;
        }

        private void dgvwLogList_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            decimal nResultID;
            // �� �������� ������ �� ������ ����� ID � 1 ������� (�������)
            if (dgvwLogList.CurrentRow != null)
            {
                // �������� ID ������
                nResultID = Convert.ToDecimal(dgvwLogList.CurrentRow.Cells["ID"].Value);
                
                // �������� �������� � ��������
                if(nResultID > 0)
                    nResult = nResultID;
                this.Close();
            }
            

        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            decimal nResultID;
            // �� ������ ����� ID � 1 ������� (�������) ������� ������
            if (dgvwLogList.CurrentRow != null)
            {
                // �������� ID ������
                nResultID = Convert.ToDecimal(dgvwLogList.CurrentRow.Cells["ID"].Value);

                // �������� �������� � ��������
                if (nResultID > 0)
                    nResult = nResultID;
                
                this.Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            nResult = -1;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            nResult = 0;
            this.Close();
        }

    }
}