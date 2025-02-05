using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Configuration;
using System.Collections.Specialized;
using System.Globalization;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Data.Odbc;
using System.Text.RegularExpressions;
using System.Xml;



namespace InfoChange
{
    public partial class Form1 : Form
    {

        // These are the Win32 error code for file not found or access denied.
        const int ERROR_FILE_NOT_FOUND = 2;
        const int ERROR_ACCESS_DENIED = 5;

        private OleDbConnection con;
        private OleDbConnection DBFcon;
        private OleDbCommand m_cmd;
        public String fullpath;

        String query_cred_org_path;

        String cred_org_path;
        String archive_cred_org_path;

        String sber_path;
        String archive_sber_path;

        String ktfoms_path;
        String archive_ktfoms_path;

        String pens_path;
        String archive_pens_path;

        String potd_path;
        String archive_potd_path;

        String krc_path;
        String archive_krc_path;

        String archive_folder_tofind;

        String[] Legal_List;
        String[] Legal_Name_List;
        String[] Legal_�onv_List;

        //int ktfoms_id;
        //int pens_id;
        //int potd_id;
        //int krc_id;

        Decimal ktfoms_id;
        Decimal pens_id;
        Decimal potd_id;
        Decimal krc_id;
        Decimal mvd_id;

        DateTime DatZapr1;
        DateTime DatZapr2;

        DateTime DatZapr1_sber;
        DateTime DatZapr2_sber;

        DateTime DatZapr1_ktfoms;
        DateTime DatZapr2_ktfoms;
        
        DateTime DatZapr1_pens;
        DateTime DatZapr2_pens;
        DateTime DatZapr1_potd;
        DateTime DatZapr2_potd;
        DateTime DatZapr1_gibd;
        DateTime DatZapr2_gibd;
        DateTime DatZapr1_krc;
        DateTime DatZapr2_krc;

        DataTable DT_reg;
        DataTable DT_okon;
        DataTable DT_doc;
        DataTable DT_doc_fiz;
        DataTable DT_doc_jur;
        DataTable DT_sber_reg;
        DataTable DT_ktfoms_reg;
        DataTable DT_ktfoms_doc;
        DataTable DT_pens_reg;
        DataTable DT_pens_doc;
        //DataTable DT_potd_reg;
        DataTable DT_potd_doc;
        DataTable DT_gibd_reg;
        DataTable DT_gibd_rst;
        DataTable DT_krc_reg;
        DataTable DT_krc_oplat;

        ArrayList cbxItems;
        bool bVFP_DBASE;
        bool bReadFromCopy;
        bool bDateFolderAdd;
        string constr1, constr2, constrRDB, constrGIBDD;

        //DateTime dtIntTablesDeplmntDate;


        bool OooIsInstall = true;
        //string txtSberFilialCode;

        //Sorted list of IDs from XML document
        SortedList<int, int> indexList = new SortedList<int, int>();
        SortedList<int, string> nodePathDic = new SortedList<int, string>();


        public Form1()
        {
            InitializeComponent();
            // ������� ������ � ����������� �� app.config
            Properties.Settings s = new Properties.Settings();
            // ������������ ����� ����������� �� �������� ������ �����������
            constr1 = s.ConnectionString;
            constr2 = s.ConnectionString2;
            constrRDB = s.ConnectionStringRDB;
            constrGIBDD = s.ConnectionStringGIBDD;

            con = new OleDbConnection(s.ConnectionStringRDB);

            query_cred_org_path = s.Cred_org_path; //"c:\\temp\\cred_org";
            archive_folder_tofind = s.Archive_folder; // "c:\\temp\\tofind";

            cred_org_path = s.Cred_org_path;
            sber_path = s.Sber_path;
            ktfoms_path = s.Ktfoms_path;
            ktfoms_id = Convert.ToDecimal(s.Ktfoms_id);
            pens_path = s.Pens_path;
            pens_id = Convert.ToDecimal(s.Pens_id);
            potd_path = s.Potd_path;
            potd_id = Convert.ToDecimal(s.Potd_id);
            krc_path = s.Krc_path;
            krc_id = Convert.ToDecimal(s.Krc_id);
            mvd_id = Convert.ToDecimal(s.Mvd_id);

            archive_cred_org_path = s.Archive_credorg_path;
            archive_sber_path = s.Archive_sber_path;
            archive_ktfoms_path = s.Archive_ktfoms_path;
            archive_pens_path = s.Archive_pens_path;
            archive_potd_path = s.Archive_potd_path;
            archive_krc_path = s.Archive_krc_path;

            bVFP_DBASE = s.bVFP_DBASE;
            bReadFromCopy = s.readFromCopy;
            
            bDateFolderAdd = s.bDateFolderAdd;

            Legal_List = s.Legal_list.Split(',');

            //txtSberFilialCode = s.SberFilial_num;

            DatZapr1 = (DateTime.Today).AddDays(-7);
            DatZapr2 = DateTime.Today;

            //�������� ��������� ��� ��������� �����
            //load_gibdlist();

            //***�������*GIBDD***
            //tabControl1.Controls.Remove(tabGibd);
            //tabControl1.Update();

            //***�������*KRC***
            tabControl1.Controls.Remove(tabKRC);
            tabControl1.Update();

            /*
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
            {
                byte[] infb = File.ReadAllBytes(openFileDialog1.FileName);

                FileStream fss = File.OpenRead(openFileDialog1.FileName);
                //string tss = fss.;

                //tss = File.ReadAllLines(openFileDialog1.FileName);
                //tss = ConvertDOS("��������");
            }*/

            //OOo_Writer child = new OOo_Writer();
            //OooIsInstall = child.isOOoInstalled();
        }

        // ������� �������� �� ����� ����, ���������� ����� + ������� ����, ������ �������� ���� ����� � ������� �����
        private void CreatePathWithDate(string txtPathWithoutDate){

            string txtCurrDateFolder = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString().PadLeft(2, '0') + DateTime.Today.Day.ToString().PadLeft(2, '0');

            if (Directory.Exists(string.Format(@"{0}\{1}", txtPathWithoutDate, txtCurrDateFolder)))
            {
                    // ����� ������� ����� ���� � ��������� _mmss ��� release_name
                    DateTime dtFixNowDate = DateTime.Now;
                    string suffix = "_" + dtFixNowDate.Hour.ToString().PadLeft(2, '0') + dtFixNowDate.Minute.ToString().PadLeft(2, '0') + dtFixNowDate.Second.ToString().PadLeft(2, '0');
                    DialogResult rv = MessageBox.Show("�� ���� " + string.Format(@"{0}\{1}", txtPathWithoutDate, txtCurrDateFolder) + ", ��������� � ���������������� �����, ���������� ����. ����� ���� ����� �������� � ����� " + txtCurrDateFolder + suffix, "��������", MessageBoxButtons.OK);
                    txtCurrDateFolder += suffix;
            }
            Directory.CreateDirectory(string.Format(@"{0}\{1}", txtPathWithoutDate, txtCurrDateFolder));
            this.fullpath = string.Format(@"{0}\{1}", txtPathWithoutDate, txtCurrDateFolder);

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

        private DataTable GetDataTableFromFB(string txtSql, string tblName, IsolationLevel islLevel)
        {
            DataSet ds = new DataSet();
            DataTable tbl = ds.Tables.Add(tblName);
            try
            {
                // ��������� ����������� - � �� ����� ������� ��� �� �������
                if (con != null && con.State != ConnectionState.Closed) con.Close();

                con.Open();
                OleDbTransaction tran = con.BeginTransaction(islLevel);
                OleDbCommand cmd = new OleDbCommand(txtSql, con, tran);

                using (OleDbDataReader rdr = cmd.ExecuteReader(CommandBehavior.Default))
                {
                    ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                    rdr.Close();
                }

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
            return tbl;
        }

        private DataTable GetDataTableFromFB(string txtSql, string tblName){
            DataSet ds = new DataSet();
            DataTable tbl = ds.Tables.Add(tblName);
            try
            {
                // ��������� ����������� - � �� ����� ������� ��� �� �������
                if (con != null && con.State != ConnectionState.Closed) con.Close();

                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.RepeatableRead);
                OleDbCommand cmd = new OleDbCommand(txtSql, con, tran);

                using (OleDbDataReader rdr = cmd.ExecuteReader(CommandBehavior.Default))
                {
                    ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                    rdr.Close();
                }

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
            return tbl;
        }

        private DataTable GetDBFTable(string txtSql, string tblName, string filename)
        {
            DataSet ds = new DataSet();
            DataTable tbl = ds.Tables.Add(tblName);
            try
            {
                DBFcon = new OleDbConnection();
                DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + filename + ";Mode=Read;Collating Sequence=RUSSIAN");
                DBFcon.Open();
                m_cmd = new OleDbCommand();
                m_cmd.Connection = DBFcon;
                m_cmd.CommandText = txtSql;
                using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                {
                    ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                    rdr.Close();
                }

                DBFcon.Close();
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
            return ds.Tables[tblName];
        }


        private int InsertKTFOMSZapros(DateTime dat1, DateTime dat2)
        {
            int vid_d = 1;
            DateTime dtDate;
            DateTime DatZapr = DateTime.Today;
            String uid = System.Guid.NewGuid().ToString();
            String osp_name = GetOSP_Name();
            uid = cutEnd(uid, 30);
            Int32 iCnt = 0;
            OleDbTransaction tran;

            string txtNomspi = "";
            int nNomspi = 0;
            Int32 iBadYearBorn = 0;
            bool bFizBadYear = false;
            string txtStatus = "���������";
            string txtText = "";
            int nResult = 0;
            int nBirthYear = 0;


            prbWritingDBF.Value = 0;


            int iDocCnt = 0;
            if (DT_ktfoms_doc != null) iDocCnt = DT_ktfoms_doc.Rows.Count;
            prbWritingDBF.Maximum = iDocCnt;
            
            prbWritingDBF.Step = 1;


            try
            {
                string txtKtfomsName = cutEnd((GetLegal_Name(ktfoms_id)).Trim(), 200);

                string txtKtfomsConv, txtKtfomsConv1, txtKtfomsConv2;

                txtKtfomsConv = GetLegal_Conv(ktfoms_id);

                if (txtKtfomsConv.Trim().Equals(""))
                {
                    txtKtfomsConv = "����������� �� ������ ����������� �� ������������� ������������ ����������� �� 1 ������ 2006 ����";
                }

                // ����������� txtKtfomsName �� ������� ���� �� ������ 120 �������� � 2-� �� ������ 200 ��������

                if (txtKtfomsConv.Length > 120)
                {
                    txtKtfomsConv1 = txtKtfomsConv.Substring(0, 120);
                    txtKtfomsConv2 = txtKtfomsConv.Substring(120);
                    if (txtKtfomsConv2.Length > 200) txtKtfomsConv2 = txtKtfomsConv2.Remove(200);
                }
                else
                {
                    txtKtfomsConv1 = txtKtfomsConv;
                    txtKtfomsConv2 = "";
                }


                if (con != null && con.State != ConnectionState.Closed) con.Close();

                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                #region "ktfoms_doc"
                if (DT_ktfoms_doc != null)
                {
                    foreach (DataRow row in DT_ktfoms_doc.Rows)
                    {
                        bFizBadYear = false;
                        nResult = 0;
                        txtStatus = "���������";
                        txtText = "";
                        nBirthYear = 0;

                        m_cmd = new OleDbCommand();
                        m_cmd.Connection = con;
                        m_cmd.Transaction = tran;

                        m_cmd.CommandText = "INSERT INTO ZAPROS (PK, NUM_ZAPR, SUB_DIV, FIO_SPI, NUM_IP, DATE_ZAPR, VID_D, INN_D, NAME_D, DATE_R, NUM_RES, DATE_RES, RESULT, FK_DOC, FK_IP, FK_ID, NUM_ID, TEXT, DATE_BEG, DATE_END, ADRESS, NUM_PACK, NUM_ZAPR_IN_PACK, STATUS, DATE_SEND, TEXT_ERROR, USCODE, CONVENTION, WHYRESPONS, WHYPREPARE, PASSPORT, SUMM, TEMP, FK_LEGAL, ADRESAT)";
                        m_cmd.CommandText += " VALUES(GEN_ID(S_N_ZAPROS, 1), :NUM_ZAPR, :SUB_DIV, :FIO_SPI, :NUM_IP, :DATE_ZAPR, :VID_D, :INN_D, :NAME_D, :DATE_R, :NUM_RES, :DATE_RES, :RESULT, :FK_DOC, :FK_IP, :FK_ID, :NUM_ID, :TEXT, :DATE_BEG, :DATE_END, :ADRESS, :NUM_PACK, :NUM_ZAPR_IN_PACK, :STATUS, :DATE_SEND, :TEXT_ERROR, :USCODE, :CONVENTION, :WHYRESPONS, :WHYPREPARE, :PASSPORT, :SUMM, :TEMP, :FK_LEGAL, :ADRESAT)";


                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR", Convert.ToString(row["ZAPROS"]).Trim()));

                        m_cmd.Parameters.Add(new OleDbParameter(":SUB_DIV", cutEnd(osp_name, 100).Trim()));

                      
                        txtNomspi = Convert.ToString(row["NOMSPI"]).Trim();
                        if (!Int32.TryParse(txtNomspi, out nNomspi))
                        {
                            nNomspi = 0;
                        }
                        
                        OleDbCommand spi_name_cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE USCODE = " + Convert.ToString(nNomspi), con, tran);
                        String spi_name = Convert.ToString(spi_name_cmd.ExecuteScalar());

                        m_cmd.Parameters.Add(new OleDbParameter(":FIO_SPI", cutEnd(spi_name, 100).Trim()));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_IP", cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40)));


                        //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// ������� ��� ���� ����������� ��
                        //{
                        //    DatZapr = DateTime.Today;
                        //}
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                        // � ��� ��� ���� ������ �������������

                        vid_d = 1; // ���. ����
                        //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                        //{
                        //    vid_d = 1;// ���. ����
                        //}

                        m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                        //m_cmd.Parameters.Add(new OleDbParameter(":INN_D", Convert.ToString(row["INNORG"])));
                        m_cmd.Parameters.Add(new OleDbParameter(":INN_D", ""));

                        m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 100)));

                        // TODO: ������� �������� �� ������������ ���� �������� � ���� ������ - �� ������ ������ � ����
                        if (vid_d == 1) // �������� ������ ��� ���. ���
                        {
                            nBirthYear = parseBirthDate(Convert.ToString(row["DATROZHD"]));
                            if (nBirthYear == 0)
                            {
                                nResult = 1;
                                txtStatus = "������ � �������";
                                txtText = "����������� ��������� ��������� ���� ��� ��� �������� (������ ##.##.#### ��� ####)";
                                bFizBadYear = true;
                            }
                        }
                        if (!DateTime.TryParse(Convert.ToString(row["DATROZHD"]), out dtDate))
                        {
                            dtDate = DateTime.MaxValue;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// ��� ����� ������ �� ������� ���-��

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// ���� ������, �� ������ �������� ��� ������

                        m_cmd.Parameters.Add(new OleDbParameter(":RESULT", nResult));

                        //m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ������� �� ���������������, 1 - ��� ���. �� ��������, ������ 1 - ���� ���-� �� ��������) (��� ��� � FIND - ��� ������ 1, ���������� 0)

                        Int32 iKey = -1;
                        //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                        //{
                        //    iKey = 0;
                        //}

                        // TODO: �������, ��� ��� ������� ������ �� DOCUMENTS!!!
                        // ���� ��� � ������� DOCUMENTS ��������� ������, ���� �� �����
                        m_cmd.Parameters.Add(new OleDbParameter(":FK_DOC", iKey));



                        if (!Int32.TryParse(Convert.ToString(row["FK_IP"]), out iKey))
                        {
                            iKey = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":FK_IP", iKey));


                        if (!Int32.TryParse(Convert.ToString(row["FK_ID"]), out iKey))
                        {
                            iKey = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":FK_ID", iKey));


                        OleDbCommand num_id_cmd = new OleDbCommand("Select NUM_ID from ID where ID.PK = " + Convert.ToString(row["FK_ID"]), con, tran);
                        String num_id = Convert.ToString(num_id_cmd.ExecuteScalar());

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ID", cutEnd(num_id.Trim(), 30)));

                        m_cmd.Parameters.Add(new OleDbParameter(":TEXT", txtText));// ����� ������ ��-��������� ������
                        //m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// ����� ������ - ������

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", dat1));
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", dat2));

                        m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250)));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                        //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                        m_cmd.Parameters.Add(new OleDbParameter(":STATUS", txtStatus));
                        //m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "���������"));

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                        if (txtKtfomsConv2.Length > 0)
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", txtKtfomsConv2));
                        }
                        else
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", txtText)); // ���� ��� ������ �� ����� ������, � ���� ���� �� ����� text
                        }

           
                        if (!(Int32.TryParse(Convert.ToString(row["NOMSPI"]), out iKey)))
                        {
                            iKey = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":USCODE", iKey));

                        //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));


                        m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", txtKtfomsConv1));

                        m_cmd.Parameters.Add(new OleDbParameter(":WHYRESPONS", ""));

                        m_cmd.Parameters.Add(new OleDbParameter(":WHYPREPARE", ""));

                        // ����� � ������� persons � ��� ���� ������ ������
                        // ������� persons �� ip.fk; � ������� physical �� person.tablename + person.FK
                        // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                        //

                        m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// ����� ������ ����. ������

                        Double sum = 0;
                        //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                        //{
                        //    sum = 0;
                        //}
                        m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                        // TODO: ����� ��� ��� NOMIP � tofind
                        // ������ NOMIP ����� ����� �� IPNO_NUM
                        String txtNOMIP = Convert.ToString(row["ZAPROS"]).Trim();
                        if (txtNOMIP.Trim() != "")
                        {
                            String[] strings = txtNOMIP.Split('/');
                            if (!(Int32.TryParse(Convert.ToString(strings[2]), out iKey)))
                            {
                                iKey = 0;
                            }
                        }
                        else
                        {
                            iKey = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":TEMP", iKey));

                        //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));
                        //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                        m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", ktfoms_id));
                        //m_cmd.Parameters[":FK_LEGAL"].Value = Convert.ToInt32(Legal_List[i].Trim());

                        m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", txtKtfomsName));
                        //m_cmd.Parameters[":ADRESAT"].Value = Convert.ToString(Legal_Name_List[i]);
                        if (m_cmd.ExecuteNonQuery() != -1)
                        {
                            if (bFizBadYear) { iBadYearBorn++; }
                            else
                            {
                                iCnt++;
                            }
                            prbWritingDBF.PerformStep();
                            prbWritingDBF.Refresh();
                            System.Windows.Forms.Application.DoEvents();
                        }
                    }
                }
                #endregion

                if(tran != null) tran.Commit();
                if(m_cmd != null) m_cmd.Dispose();
                if(con.State.Equals(ConnectionState.Open)) con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    //tran.Rollback();
                    if (con.State != System.Data.ConnectionState.Closed) con.Close();
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con.State != System.Data.ConnectionState.Closed) con.Close();
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

            return iCnt;
        }

        private int InsertPENSZapros(DateTime dat1, DateTime dat2)
        {
            int vid_d = 1;
            DateTime dtDate;
            DateTime DatZapr = DateTime.Today;
            String uid = System.Guid.NewGuid().ToString();
            String osp_name = GetOSP_Name();
            uid = cutEnd(uid, 30);
            Int32 iCnt = 0;
            OleDbTransaction tran;
            string txtNomspi = "";
            int nNomspi = 0;
            Int32 iBadYearBorn = 0;
            bool bFizBadYear = false;
            string txtStatus = "���������";
            string txtText = "";
            int nResult = 0;
            int nBirthYear = 0;

            prbWritingDBF.Value = 0;
            if (DT_pens_doc != null)
            {
                // ��������� ������ �� ���������
                //prbWritingDBF.Maximum = DT_pens_reg.Rows.Count + DT_pens_doc.Rows.Count; // *Legal_List.Length;
                prbWritingDBF.Maximum = DT_pens_doc.Rows.Count;
            }
            // ��������� ������ �� ���������
            //else
            //{
            //    prbWritingDBF.Maximum = DT_pens_reg.Rows.Count;// + DT_ktfoms_doc.Rows.Count; // *Legal_List.Length;   
            //}
            
            prbWritingDBF.Step = 1;


            try
            {
                string txtPensName = cutEnd((GetLegal_Name(pens_id)).Trim(), 200);

                string txtPensConv = GetLegal_Conv(pens_id);
                if (txtPensConv.Trim().Equals(""))
                {
                    txtPensConv = "����������� � 10/03-08-1778 � ������� ��������������� ������ � �������������� ����� ����� �� �� � ���������� ����������� ����� �� �� �� �� 24 ����� 2008 ����";
                }
                txtPensConv = cutEnd(txtPensConv, 120);

                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                #region "penss_reg"
                // ��������� ������ �� ���������
                //foreach (DataRow row in DT_pens_reg.Rows)
                //{
                //    m_cmd = new OleDbCommand();
                //    m_cmd.Connection = con;
                //    m_cmd.Transaction = tran;

                //    m_cmd.CommandText = "INSERT INTO ZAPROS (PK, NUM_ZAPR, SUB_DIV, FIO_SPI, NUM_IP, DATE_ZAPR, VID_D, INN_D, NAME_D, DATE_R, NUM_RES, DATE_RES, RESULT, FK_DOC, FK_IP, FK_ID, NUM_ID, TEXT, DATE_BEG, DATE_END, ADRESS, NUM_PACK, NUM_ZAPR_IN_PACK, STATUS, DATE_SEND, TEXT_ERROR, USCODE, CONVENTION, WHYRESPONS, WHYPREPARE, PASSPORT, SUMM, TEMP, FK_LEGAL, ADRESAT)";
                //    m_cmd.CommandText += " VALUES(GEN_ID(S_N_ZAPROS, 1), :NUM_ZAPR, :SUB_DIV, :FIO_SPI, :NUM_IP, :DATE_ZAPR, :VID_D, :INN_D, :NAME_D, :DATE_R, :NUM_RES, :DATE_RES, :RESULT, :FK_DOC, :FK_IP, :FK_ID, :NUM_ID, :TEXT, :DATE_BEG, :DATE_END, :ADRESS, :NUM_PACK, :NUM_ZAPR_IN_PACK, :STATUS, :DATE_SEND, :TEXT_ERROR, :USCODE, :CONVENTION, :WHYRESPONS, :WHYPREPARE, :PASSPORT, :SUMM, :TEMP, :FK_LEGAL, :ADRESAT)";


                //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR", Convert.ToString(row["ZAPROS"]).Trim()));

                //    m_cmd.Parameters.Add(new OleDbParameter(":SUB_DIV", cutEnd(osp_name, 100).Trim()));

                //    txtNomspi = Convert.ToString(row["NOMSPI"]).Trim();
                //    if (!Int32.TryParse(txtNomspi, out nNomspi))
                //    {
                //        nNomspi = 0;
                //    }


                //    OleDbCommand spi_name_cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE USCODE = " + Convert.ToString(nNomspi), con, tran);
                //    String spi_name = Convert.ToString(spi_name_cmd.ExecuteScalar());

                //    m_cmd.Parameters.Add(new OleDbParameter(":FIO_SPI", cutEnd(spi_name, 100).Trim()));

                //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_IP", cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40)));


                //    //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// ������� ��� ���� ����������� ��
                //    //{
                //    //    DatZapr = DateTime.Today;
                //    //}
                //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                //    // � ��� ��� ���� ������ �������������

                //    vid_d = 1; // ���. ����
                //    //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                //    //{
                //    //    vid_d = 1;// ���. ����
                //    //}

                //    m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                //    //m_cmd.Parameters.Add(new OleDbParameter(":INN_D", Convert.ToString(row["INNORG"])));
                //    m_cmd.Parameters.Add(new OleDbParameter(":INN_D", ""));

                //    m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 100)));

                //    if (!DateTime.TryParse(Convert.ToString(row["DATROZHD"]), out dtDate))
                //    {
                //        dtDate = DateTime.MaxValue;
                //    }
                //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// ��� ����� ������ �� ������� ���-��

                //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// ���� ������, �� ������ �������� ��� ������

                //    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ������� �� ���������������, 1 - ��� ���. �� ��������, ������ 1 - ���� ���-� �� ��������) (��� ��� � FIND - ��� ������ 1, ���������� 0)

                //    Int32 iKey = -1;
                //    //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                //    //{
                //    //    iKey = 0;
                //    //}

                //    // TODO: �������, ��� ��� ������� ������ �� DOCUMENTS!!!
                //    // ���� ��� � ������� DOCUMENTS ��������� ������, ���� �� �����
                //    m_cmd.Parameters.Add(new OleDbParameter(":FK_DOC", iKey));



                //    if (!Int32.TryParse(Convert.ToString(row["FK_IP"]), out iKey))
                //    {
                //        iKey = 0;
                //    }
                //    m_cmd.Parameters.Add(new OleDbParameter(":FK_IP", iKey));


                //    if (!Int32.TryParse(Convert.ToString(row["FK_ID"]), out iKey))
                //    {
                //        iKey = 0;
                //    }
                //    m_cmd.Parameters.Add(new OleDbParameter(":FK_ID", iKey));


                //    OleDbCommand num_id_cmd = new OleDbCommand("Select NUM_ID from ID where ID.PK = " + Convert.ToString(row["FK_ID"]), con, tran);
                //    String num_id = Convert.ToString(num_id_cmd.ExecuteScalar());

                //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ID", cutEnd(num_id.Trim(), 30)));

                //    m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// ����� ������ - ������

                //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", dat1));
                //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", dat2));

                //    m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250)));

                //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                //    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                //    m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "���������"));

                //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                //    m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", ""));

                //    if (!(Int32.TryParse(Convert.ToString(row["NOMSPI"]).Trim(), out iKey)))
                //    {
                //        iKey = 0;
                //    }
                //    m_cmd.Parameters.Add(new OleDbParameter(":USCODE", iKey));

                //    //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));


                //    m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", txtPensConv));

                //    m_cmd.Parameters.Add(new OleDbParameter(":WHYRESPONS", ""));

                //    m_cmd.Parameters.Add(new OleDbParameter(":WHYPREPARE", ""));

                //    // ����� � ������� persons � ��� ���� ������ ������
                //    // ������� persons �� ip.fk; � ������� physical �� person.tablename + person.FK
                //    // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                //    //

                //    m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// ����� ������ ����. ������

                //    Double sum = 0;
                //    //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                //    //{
                //    //    sum = 0;
                //    //}
                //    m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                //    // TODO: ����� ��� ��� NOMIP � tofind
                //    String txtNOMIP = Convert.ToString(row["ZAPROS"]).Trim();
                //    if (txtNOMIP.Trim() != "")
                //    {
                //        String[] strings = txtNOMIP.Split('/');
                //        if (!(Int32.TryParse(Convert.ToString(strings[2]), out iKey)))
                //        {
                //            iKey = 0;
                //        }
                //    }
                //    else
                //    {
                //        iKey = 0;
                //    }
                //    m_cmd.Parameters.Add(new OleDbParameter(":TEMP", iKey));

                //    //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));
                //    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                //    m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", pens_id));
                //    //m_cmd.Parameters[":FK_LEGAL"].Value = Convert.ToInt32(Legal_List[i].Trim());

                //    m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", txtPensName));
                //    //m_cmd.Parameters[":ADRESAT"].Value = Convert.ToString(Legal_Name_List[i]);
                //    if (m_cmd.ExecuteNonQuery() != -1)
                //    {
                //        iCnt++;
                //        prbWritingDBF.PerformStep();
                //    }
                //}
                #endregion

                #region "penss_doc"
                if (DT_pens_doc != null)
                {
                    foreach (DataRow row in DT_pens_doc.Rows)
                    {
                        bFizBadYear = false;
                        nResult = 0;
                        txtStatus = "���������";
                        txtText = "";
                        nBirthYear = 0;

                        m_cmd = new OleDbCommand();
                        m_cmd.Connection = con;
                        m_cmd.Transaction = tran;

                        m_cmd.CommandText = "INSERT INTO ZAPROS (PK, NUM_ZAPR, SUB_DIV, FIO_SPI, NUM_IP, DATE_ZAPR, VID_D, INN_D, NAME_D, DATE_R, NUM_RES, DATE_RES, RESULT, FK_DOC, FK_IP, FK_ID, NUM_ID, TEXT, DATE_BEG, DATE_END, ADRESS, NUM_PACK, NUM_ZAPR_IN_PACK, STATUS, DATE_SEND, TEXT_ERROR, USCODE, CONVENTION, WHYRESPONS, WHYPREPARE, PASSPORT, SUMM, TEMP, FK_LEGAL, ADRESAT)";
                        m_cmd.CommandText += " VALUES(GEN_ID(S_N_ZAPROS, 1), :NUM_ZAPR, :SUB_DIV, :FIO_SPI, :NUM_IP, :DATE_ZAPR, :VID_D, :INN_D, :NAME_D, :DATE_R, :NUM_RES, :DATE_RES, :RESULT, :FK_DOC, :FK_IP, :FK_ID, :NUM_ID, :TEXT, :DATE_BEG, :DATE_END, :ADRESS, :NUM_PACK, :NUM_ZAPR_IN_PACK, :STATUS, :DATE_SEND, :TEXT_ERROR, :USCODE, :CONVENTION, :WHYRESPONS, :WHYPREPARE, :PASSPORT, :SUMM, :TEMP, :FK_LEGAL, :ADRESAT)";


                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR", Convert.ToString(row["ZAPROS"]).Trim()));

                        m_cmd.Parameters.Add(new OleDbParameter(":SUB_DIV", cutEnd(osp_name, 100).Trim()));


                        txtNomspi = Convert.ToString(row["NOMSPI"]).Trim();
                        if (!Int32.TryParse(txtNomspi, out nNomspi))
                        {
                            nNomspi = 0;
                        }

                        OleDbCommand spi_name_cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE USCODE = " + Convert.ToString(nNomspi), con, tran);
                        String spi_name = Convert.ToString(spi_name_cmd.ExecuteScalar());

                        m_cmd.Parameters.Add(new OleDbParameter(":FIO_SPI", cutEnd(spi_name, 100).Trim()));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_IP", cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40)));


                        //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// ������� ��� ���� ����������� ��
                        //{
                        //    DatZapr = DateTime.Today;
                        //}
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                        // � ��� ��� ���� ������ �������������

                        vid_d = 1; // ���. ����
                        //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                        //{
                        //    vid_d = 1;// ���. ����
                        //}

                        m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                        //m_cmd.Parameters.Add(new OleDbParameter(":INN_D", Convert.ToString(row["INNORG"])));
                        m_cmd.Parameters.Add(new OleDbParameter(":INN_D", ""));

                        m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 100)));

                        // TODO: ������� �������� �� ������������ ���� �������� � ���� ������ - �� ������ ������ � ����
                        if (vid_d == 1) // �������� ������ ��� ���. ���
                        {
                            nBirthYear = parseBirthDate(Convert.ToString(row["DATROZHD"]));
                            if (nBirthYear == 0)
                            {
                                nResult = 1; // � ����� � ��� ��� ����� ����
                                txtStatus = "������ � �������";
                                txtText = "����������� ��������� ��������� ���� ��� ��� �������� (������ ##.##.#### ��� ####)";
                                bFizBadYear = true;
                            }
                        }

                        if (!DateTime.TryParse(Convert.ToString(row["DATROZHD"]), out dtDate))
                        {
                            dtDate = DateTime.MaxValue;
                        }

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// ��� ����� ������ �� ������� ���-��

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// ���� ������, �� ������ �������� ��� ������

                        m_cmd.Parameters.Add(new OleDbParameter(":RESULT", nResult));
                        //m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ������� �� ���������������, 1 - ��� ���. �� ��������, ������ 1 - ���� ���-� �� ��������) (��� ��� � FIND - ��� ������ 1, ���������� 0)

                        Int32 iKey = -1;
                        //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                        //{
                        //    iKey = 0;
                        //}

                        // TODO: �������, ��� ��� ������� ������ �� DOCUMENTS!!!
                        // ���� ��� � ������� DOCUMENTS ��������� ������, ���� �� �����
                        m_cmd.Parameters.Add(new OleDbParameter(":FK_DOC", iKey));



                        if (!Int32.TryParse(Convert.ToString(row["FK_IP"]), out iKey))
                        {
                            iKey = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":FK_IP", iKey));


                        if (!Int32.TryParse(Convert.ToString(row["FK_ID"]), out iKey))
                        {
                            iKey = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":FK_ID", iKey));


                        OleDbCommand num_id_cmd = new OleDbCommand("Select NUM_ID from ID where ID.PK = " + Convert.ToString(row["FK_ID"]), con, tran);
                        String num_id = Convert.ToString(num_id_cmd.ExecuteScalar());

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ID", cutEnd(num_id.Trim(), 30)));

                        m_cmd.Parameters.Add(new OleDbParameter(":TEXT", txtText));// ����� ������ ��-��������� ������
                        //m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// ����� ������ - ������

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", DatZapr1_pens));
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", DatZapr2_pens));

                        m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250)));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                        //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                        m_cmd.Parameters.Add(new OleDbParameter(":STATUS", txtStatus));
                        //m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "���������"));

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                        m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", txtText)); // ���� ��� ������ �� ����� ������, � ���� ���� �� ����� text
                        //m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", ""));

                        if (!(Int32.TryParse(Convert.ToString(row["NOMSPI"]), out iKey)))
                        {
                            iKey = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":USCODE", iKey));

                        //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));

                        m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", txtPensConv));

                        m_cmd.Parameters.Add(new OleDbParameter(":WHYRESPONS", ""));

                        m_cmd.Parameters.Add(new OleDbParameter(":WHYPREPARE", ""));

                        // ����� � ������� persons � ��� ���� ������ ������
                        // ������� persons �� ip.fk; � ������� physical �� person.tablename + person.FK
                        // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                        //

                        m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// ����� ������ ����. ������

                        Double sum = 0;
                        //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                        //{
                        //    sum = 0;
                        //}
                        m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                        // TODO: ����� ��� ��� NOMIP � tofind
                        String txtNOMIP = Convert.ToString(row["ZAPROS"]).Trim();
                        if (txtNOMIP.Trim() != "")
                        {
                            String[] strings = txtNOMIP.Split('/');
                            if (!(Int32.TryParse(Convert.ToString(strings[2]), out iKey)))
                            {
                                iKey = 0;
                            }
                        }
                        else
                        {
                            iKey = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":TEMP", iKey));

                        //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));
                        //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                        m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", pens_id));
                        //m_cmd.Parameters[":FK_LEGAL"].Value = Convert.ToInt32(Legal_List[i].Trim());

                        m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", txtPensName));
                        //m_cmd.Parameters[":ADRESAT"].Value = Convert.ToString(Legal_Name_List[i]);
                        if (m_cmd.ExecuteNonQuery() != -1)
                        {
                            if (bFizBadYear) { iBadYearBorn++; }
                            else
                            {
                                iCnt++;
                            }
                            prbWritingDBF.PerformStep();
                        }
                    }
                }
                #endregion

                tran.Commit();
                con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    //tran.Rollback();
                    con.Close();
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

            return iCnt;
        }

        private int InsertPOTDZapros(DateTime dat1, DateTime dat2)
        {
            int vid_d = 1;
            DateTime dtDate;
            DateTime DatZapr = DateTime.Today;
            String uid = System.Guid.NewGuid().ToString();
            String osp_name = GetOSP_Name();
            uid = cutEnd(uid, 30);
            Int32 iCnt = 0;
            OleDbTransaction tran;
            Int32 iBadYearBorn = 0;
            bool bFizBadYear = false;
            string txtStatus = "���������";
            string txtText = "";
            int nResult = 0;
            int nBirthYear = 0;

            string txtNomspi = "";
            int nNomspi = 0;

            prbWritingDBF.Value = 0;
            prbWritingDBF.Maximum = DT_potd_doc.Rows.Count;// + DT_ktfoms_doc.Rows.Count; // *Legal_List.Length;
            prbWritingDBF.Step = 1;


            try
            {
                string txtPotdName = cutEnd((GetLegal_Name(potd_id)).Trim(), 200);

                string txtPotdConv = GetLegal_Conv(potd_id);
                if (txtPotdConv.Trim().Equals(""))
                {
                    txtPotdConv = "����������� �� ������ ����������� �� ������� ���������� ��������� ������ �� 27 ����� 2008 �.";
                }
                
                if (con != null && con.State != ConnectionState.Closed) con.Close();

                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                #region "potd_doc"

                foreach (DataRow row in DT_potd_doc.Rows)
                {
                    bFizBadYear = false;
                    nResult = 0;
                    txtStatus = "���������";
                    txtText = "";
                    nBirthYear = 0;

                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = con;
                    m_cmd.Transaction = tran;

                    m_cmd.CommandText = "INSERT INTO ZAPROS (PK, NUM_ZAPR, SUB_DIV, FIO_SPI, NUM_IP, DATE_ZAPR, VID_D, INN_D, NAME_D, DATE_R, NUM_RES, DATE_RES, RESULT, FK_DOC, FK_IP, FK_ID, NUM_ID, TEXT, DATE_BEG, DATE_END, ADRESS, NUM_PACK, NUM_ZAPR_IN_PACK, STATUS, DATE_SEND, TEXT_ERROR, USCODE, CONVENTION, WHYRESPONS, WHYPREPARE, PASSPORT, SUMM, TEMP, FK_LEGAL, ADRESAT)";
                    m_cmd.CommandText += " VALUES(GEN_ID(S_N_ZAPROS, 1), :NUM_ZAPR, :SUB_DIV, :FIO_SPI, :NUM_IP, :DATE_ZAPR, :VID_D, :INN_D, :NAME_D, :DATE_R, :NUM_RES, :DATE_RES, :RESULT, :FK_DOC, :FK_IP, :FK_ID, :NUM_ID, :TEXT, :DATE_BEG, :DATE_END, :ADRESS, :NUM_PACK, :NUM_ZAPR_IN_PACK, :STATUS, :DATE_SEND, :TEXT_ERROR, :USCODE, :CONVENTION, :WHYRESPONS, :WHYPREPARE, :PASSPORT, :SUMM, :TEMP, :FK_LEGAL, :ADRESAT)";


                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR", Convert.ToString(row["ZAPROS"]).Trim()));

                    m_cmd.Parameters.Add(new OleDbParameter(":SUB_DIV", cutEnd(osp_name, 100).Trim()));

                    txtNomspi = Convert.ToString(row["NOMSPI"]).Trim();
                    if (!Int32.TryParse(txtNomspi, out nNomspi))
                    {
                        nNomspi = 0;
                    }

                    OleDbCommand spi_name_cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE USCODE = " + Convert.ToString(nNomspi), con, tran);
                    String spi_name = Convert.ToString(spi_name_cmd.ExecuteScalar());

                    m_cmd.Parameters.Add(new OleDbParameter(":FIO_SPI", cutEnd(spi_name, 100).Trim()));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_IP", cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40)));


                    //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// ������� ��� ���� ����������� ��
                    //{
                    //    DatZapr = DateTime.Today;
                    //}
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                    // � ��� ��� ���� ������ �������������

                    vid_d = 1; // ���. ����
                    //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                    //{
                    //    vid_d = 1;// ���. ����
                    //}

                    m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                    //m_cmd.Parameters.Add(new OleDbParameter(":INN_D", Convert.ToString(row["INNORG"])));
                    m_cmd.Parameters.Add(new OleDbParameter(":INN_D", ""));

                    m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 100)));

                    if (vid_d == 1) // �������� ������ ��� ���. ���
                    {
                        nBirthYear = parseBirthDate(Convert.ToString(row["DATROZHD"]));
                        if (nBirthYear == 0)
                        {
                            nResult = 1; // � ����� � ��� ��� ����� ����
                            txtStatus = "������ � �������";
                            txtText = "����������� ��������� ��������� ���� ��� ��� �������� (������ ##.##.#### ��� ####)";
                            bFizBadYear = true;
                        }
                    }

                    if (!DateTime.TryParse(Convert.ToString(row["DATROZHD"]), out dtDate))
                    {
                        dtDate = DateTime.MaxValue;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// ��� ����� ������ �� ������� ���-��

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// ���� ������, �� ������ �������� ��� ������

                    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", nResult));
                    //m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ������� �� ���������������, 1 - ��� ���. �� ��������, ������ 1 - ���� ���-� �� ��������) (��� ��� � FIND - ��� ������ 1, ���������� 0)

                    Int32 iKey = -1;
                    //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                    //{
                    //    iKey = 0;
                    //}

                    // TODO: �������, ��� ��� ������� ������ �� DOCUMENTS!!!
                    // ���� ��� � ������� DOCUMENTS ��������� ������, ���� �� �����
                    m_cmd.Parameters.Add(new OleDbParameter(":FK_DOC", iKey));



                    if (!Int32.TryParse(Convert.ToString(row["FK_IP"]), out iKey))
                    {
                        iKey = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":FK_IP", iKey));


                    if (!Int32.TryParse(Convert.ToString(row["FK_ID"]), out iKey))
                    {
                        iKey = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":FK_ID", iKey));


                    OleDbCommand num_id_cmd = new OleDbCommand("Select NUM_ID from ID where ID.PK = " + Convert.ToString(row["FK_ID"]), con, tran);
                    String num_id = Convert.ToString(num_id_cmd.ExecuteScalar());

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ID", cutEnd(num_id.Trim(), 30)));

                    m_cmd.Parameters.Add(new OleDbParameter(":TEXT", txtText));// ����� ������ ��-��������� ������
                    //m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// ����� ������ - ������

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", dat1));
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", dat2));

                    m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250)));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                    m_cmd.Parameters.Add(new OleDbParameter(":STATUS", txtStatus));
                    //m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "���������"));

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                    m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", txtText)); // ���� ��� ������ �� ����� ������, � ���� ���� �� ����� text
                    //m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", ""));

                    if (!(Int32.TryParse(Convert.ToString(row["NOMSPI"]), out iKey)))
                    {
                        iKey = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":USCODE", iKey));

                    //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));


                    m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", txtPotdConv));

                    m_cmd.Parameters.Add(new OleDbParameter(":WHYRESPONS", ""));

                    m_cmd.Parameters.Add(new OleDbParameter(":WHYPREPARE", ""));

                    // ����� � ������� persons � ��� ���� ������ ������
                    // ������� persons �� ip.fk; � ������� physical �� person.tablename + person.FK
                    // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                    //

                    m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// ����� ������ ����. ������

                    Double sum = 0;
                    //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                    //{
                    //    sum = 0;
                    //}
                    m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                    // TODO: ����� ��� ��� NOMIP � tofind
                    String txtNOMIP = Convert.ToString(row["ZAPROS"]).Trim();
                    if (txtNOMIP.Trim() != "")
                    {
                        String[] strings = txtNOMIP.Split('/');
                        if (!(Int32.TryParse(Convert.ToString(strings[2]), out iKey)))
                        {
                            iKey = 0;
                        }
                    }
                    else
                    {
                        iKey = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":TEMP", iKey));

                    //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));
                    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                    m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", potd_id));
                    //m_cmd.Parameters[":FK_LEGAL"].Value = Convert.ToInt32(Legal_List[i].Trim());

                    m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", txtPotdName));
                    //m_cmd.Parameters[":ADRESAT"].Value = Convert.ToString(Legal_Name_List[i]);
                    if (m_cmd.ExecuteNonQuery() != -1)
                    {
                        if (bFizBadYear) { iBadYearBorn++; }
                        else
                        {
                            iCnt++;
                        }
                        prbWritingDBF.PerformStep();
                    }
                }
                #endregion

                tran.Commit();
                con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    //tran.Rollback();
                    con.Close();
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

            return iCnt;
        }

        private int InsertKRCZapros(DateTime dat1, DateTime dat2)
        {
            int vid_d = 1;
            DateTime dtDate;
            DateTime DatZapr = DateTime.Today;
            String uid = System.Guid.NewGuid().ToString();
            String osp_name = GetOSP_Name();
            uid = cutEnd(uid, 30);
            Int32 iCnt = 0;
            OleDbTransaction tran;

            string txtNomspi = "";
            int nNomspi = 0;

            prbWritingDBF.Value = 0;
            prbWritingDBF.Maximum = DT_krc_reg.Rows.Count;// + DT_ktfoms_doc.Rows.Count; // *Legal_List.Length;
            prbWritingDBF.Step = 1;


            try
            {
                string txtKrcName = cutEnd((GetLegal_Name(krc_id)).Trim(), 200);

                string txtKrcConv = GetLegal_Conv(krc_id);
                if (txtKrcConv.Trim().Equals(""))
                {
                    txtKrcConv = "����������� �� ������ ����������� � ������, ���������� ���������� �� �������������� ����������.";
                }


                if (con != null && con.State != ConnectionState.Closed) con.Close();

                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                #region "krc_doc"

                foreach (DataRow row in DT_krc_reg.Rows)
                {
                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = con;
                    m_cmd.Transaction = tran;

                    m_cmd.CommandText = "INSERT INTO ZAPROS (PK, NUM_ZAPR, SUB_DIV, FIO_SPI, NUM_IP, DATE_ZAPR, VID_D, INN_D, NAME_D, DATE_R, NUM_RES, DATE_RES, RESULT, FK_DOC, FK_IP, FK_ID, NUM_ID, TEXT, DATE_BEG, DATE_END, ADRESS, NUM_PACK, NUM_ZAPR_IN_PACK, STATUS, DATE_SEND, TEXT_ERROR, USCODE, CONVENTION, WHYRESPONS, WHYPREPARE, PASSPORT, SUMM, TEMP, FK_LEGAL, ADRESAT)";
                    m_cmd.CommandText += " VALUES(GEN_ID(S_N_ZAPROS, 1), :NUM_ZAPR, :SUB_DIV, :FIO_SPI, :NUM_IP, :DATE_ZAPR, :VID_D, :INN_D, :NAME_D, :DATE_R, :NUM_RES, :DATE_RES, :RESULT, :FK_DOC, :FK_IP, :FK_ID, :NUM_ID, :TEXT, :DATE_BEG, :DATE_END, :ADRESS, :NUM_PACK, :NUM_ZAPR_IN_PACK, :STATUS, :DATE_SEND, :TEXT_ERROR, :USCODE, :CONVENTION, :WHYRESPONS, :WHYPREPARE, :PASSPORT, :SUMM, :TEMP, :FK_LEGAL, :ADRESAT)";


                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR", Convert.ToString(row["ZAPROS"]).Trim()));

                    m_cmd.Parameters.Add(new OleDbParameter(":SUB_DIV", cutEnd(osp_name, 100).Trim()));

                    txtNomspi = Convert.ToString(row["NOMSPI"]).Trim();
                    if (!Int32.TryParse(txtNomspi, out nNomspi))
                    {
                        nNomspi = 0;
                    }

                    OleDbCommand spi_name_cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE USCODE = " + Convert.ToString(nNomspi), con, tran);
                    String spi_name = Convert.ToString(spi_name_cmd.ExecuteScalar());

                    m_cmd.Parameters.Add(new OleDbParameter(":FIO_SPI", cutEnd(spi_name, 100).Trim()));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_IP", cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40)));


                    //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// ������� ��� ���� ����������� ��
                    //{
                    //    DatZapr = DateTime.Today;
                    //}
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                    // � ��� ��� ���� ������ �������������

                    vid_d = 1; // ���. ����
                    //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                    //{
                    //    vid_d = 1;// ���. ����
                    //}

                    m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                    //m_cmd.Parameters.Add(new OleDbParameter(":INN_D", Convert.ToString(row["INNORG"])));
                    m_cmd.Parameters.Add(new OleDbParameter(":INN_D", ""));

                    m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", cutEnd(Convert.ToString(row["NAMEDOL"]).Trim(), 100)));

                    if (!DateTime.TryParse(Convert.ToString(row["BORN"]), out dtDate))
                    {
                        dtDate = DateTime.MaxValue;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// ��� ����� ������ �� ������� ���-��

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// ���� ������, �� ������ �������� ��� ������

                    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ������� �� ���������������, 1 - ��� ���. �� ��������, ������ 1 - ���� ���-� �� ��������) (��� ��� � FIND - ��� ������ 1, ���������� 0)

                    Int32 iKey = -1;
                    //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                    //{
                    //    iKey = 0;
                    //}

                    // TODO: �������, ��� ��� ������� ������ �� DOCUMENTS!!!
                    // ���� ��� � ������� DOCUMENTS ��������� ������, ���� �� �����
                    m_cmd.Parameters.Add(new OleDbParameter(":FK_DOC", iKey));



                    if (!Int32.TryParse(Convert.ToString(row["FK_IP"]), out iKey))
                    {
                        iKey = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":FK_IP", iKey));


                    if (!Int32.TryParse(Convert.ToString(row["FK_ID"]), out iKey))
                    {
                        iKey = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":FK_ID", iKey));


                    OleDbCommand num_id_cmd = new OleDbCommand("Select NUM_ID from ID where ID.PK = " + Convert.ToString(row["FK_ID"]), con, tran);
                    String num_id = Convert.ToString(num_id_cmd.ExecuteScalar());

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ID", cutEnd(num_id.Trim(), 30)));

                    m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// ����� ������ - ������

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", dat1));
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", dat2));

                    m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250)));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                    m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "���������"));

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                    m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", ""));

                    if (!(Int32.TryParse(Convert.ToString(row["NOMSPI"]), out iKey)))
                    {
                        iKey = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":USCODE", iKey));

                    //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));


                    m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", txtKrcConv));

                    m_cmd.Parameters.Add(new OleDbParameter(":WHYRESPONS", ""));

                    m_cmd.Parameters.Add(new OleDbParameter(":WHYPREPARE", ""));

                    // ����� � ������� persons � ��� ���� ������ ������
                    // ������� persons �� ip.fk; � ������� physical �� person.tablename + person.FK
                    // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                    //

                    m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// ����� ������ ����. ������

                    Double sum = 0;
                    //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                    //{
                    //    sum = 0;
                    //}
                    m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                    // TODO: ����� ��� ��� NOMIP � tofind
                    String txtNOMIP = Convert.ToString(row["ZAPROS"]).Trim();
                    if (txtNOMIP.Trim() != "")
                    {
                        String[] strings = txtNOMIP.Split('/');
                        if (!(Int32.TryParse(Convert.ToString(strings[2]), out iKey)))
                        {
                            iKey = 0;
                        }
                    }
                    else
                    {
                        iKey = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":TEMP", iKey));

                    //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));
                    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                    m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", krc_id));
                    //m_cmd.Parameters[":FK_LEGAL"].Value = Convert.ToInt32(Legal_List[i].Trim());

                    m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", txtKrcName));
                    //m_cmd.Parameters[":ADRESAT"].Value = Convert.ToString(Legal_Name_List[i]);
                    if (m_cmd.ExecuteNonQuery() != -1)
                    {
                        iCnt++;
                        prbWritingDBF.PerformStep();
                    }
                }
                #endregion

                tran.Commit();
                con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    //tran.Rollback();
                    con.Close();
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

            return iCnt;
        }

        //private int InsertZapros()
        private int InsertZapros(DataTable dtTable, DateTime DatZapr1_param, DateTime DatZapr2_param)
        {
            // ������� ��� ��� ���� ������ - ��� ������ ������� ����������� ������ ��� ��������� - ���� �� ������ � Legal_List
            int vid_d = 1;
            DateTime dtDate;
            DateTime DatZapr;
            String uid = System.Guid.NewGuid().ToString();
            String osp_name = GetOSP_Name();
            Decimal osp_num = GetOSP_Num();
            String osp_h_pristav = GetOSP_H_Pristav();
            String legal_branch = GetLegal_Branch(Convert.ToInt32(Legal_List[0].Trim())); // �������� branch ��� �����
            uid = cutEnd(uid, 30);
            Int32 iCnt = 0;
            OleDbTransaction tran;

            prbWritingDBF.Value = 0;
            int iRowCnt = 0;
            if(dtTable != null) iRowCnt = dtTable.Rows.Count;
            prbWritingDBF.Maximum = iRowCnt; // *Legal_List.Length;
            prbWritingDBF.Step = 1;
            
            string txtTextError = "";
            string txtZaprosText = "";
            Int32 nBirthYear = 0;


            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                foreach (DataRow row in dtTable.Rows)
                {
                    bool bInserted = true;
                    bool bNotFullFIO = false;
                    txtZaprosText = "NULL";
                    
                    string txtStatus = "������ ����� � ��������";
                    
                    txtTextError = "";
                    
                    if (Legal_�onv_List[0].Length > 100)
                    {
                        txtTextError = Legal_�onv_List[0].Substring(100);
                    }
                    
                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = con;
                    m_cmd.Transaction = tran;
                    //m_cmd.CommandText = "INSERT INTO ZAPROS (PK, NUM_ZAPR, SUB_DIV, FIO_SPI, NUM_IP, DATE_ZAPR, VID_D, INN_D, NAME_D, DATE_R, NUM_RES, DATE_RES, RESULT, FK_DOC, FK_IP, FK_ID, NUM_ID, TEXT, DATE_BEG, DATE_END, ADRESS, NUM_PACK, NUM_ZAPR_IN_PACK, STATUS, DATE_SEND, TEXT_ERROR, USCODE, CONVENTION, WHYRESPONS, WHYPREPARE, PASSPORT, SUMM, TEMP, FK_LEGAL, ADRESAT)";
                    //m_cmd.CommandText += " VALUES(GEN_ID(S_N_ZAPROS, 1), :NUM_ZAPR, :SUB_DIV, :FIO_SPI, :NUM_IP, :DATE_ZAPR, :VID_D, :INN_D, :NAME_D, :DATE_R, :NUM_RES, :DATE_RES, :RESULT, :FK_DOC, :FK_IP, :FK_ID, :NUM_ID, :TEXT, :DATE_BEG, :DATE_END, :ADRESS, :NUM_PACK, :NUM_ZAPR_IN_PACK, :STATUS, :DATE_SEND, :TEXT_ERROR, :USCODE, :CONVENTION, :WHYRESPONS, :WHYPREPARE, :PASSPORT, :SUMM, :TEMP, :FK_LEGAL, :ADRESAT)";

                    m_cmd.CommandText = "INSERT INTO ZAPROS (PK, NUM_ZAPR, SUB_DIV, FIO_SPI, NUM_IP, DATE_ZAPR, VID_D, INN_D, NAME_D, DATE_R, NUM_RES, DATE_RES, RESULT, FK_DOC, FK_IP, FK_ID,TEMP, NUM_ID, TEXT, TYPE_ZAPR, DATE_BEG, DATE_END, ADRESS, NUM_PACK, NUM_ZAPR_IN_PACK, ADRESAT, STATUS, DATE_SEND, TEXT_ERROR, USCODE, FK_LEGAL, CONVENTION,  WHYRESPONS, WHYPREPARE, DATE_RESP, PASSPORT, SUMM, FILENAME, FILE_IMP_NAME, FILE_EXP_NAME, DATE_PRINT, BRANCH_OFFICE, NUM_SUB_DIV, H_PRISTAV, DATE_ID, DOLG_SURNAME, DOLG_NAME, DOLG_SECONDNAME, DOLG_BIRTH_YEAR, TOWN_BORN, OSB_NAME, OSB_ADDR, OSB_NUM, OSB_TEL, OSB_BIC)";
                    
                    m_cmd.CommandText += " VALUES (GEN_ID(S_N_ZAPROS, 1), :NUM_ZAPR, :SUB_DIV, :FIO_SPI, :NUM_IP, :DATE_ZAPR, :VID_D, :INN_D,";
                    
                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR", Convert.ToString(row["ZAPROS"]).Trim()));

                    m_cmd.Parameters.Add(new OleDbParameter(":SUB_DIV", cutEnd(osp_name, 100).Trim()));

                    //OleDbCommand spi_name_cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE USCODE = '" + Convert.ToString(row["NOMSPI"]) + "'", con, tran);
                    //String spi_name = Convert.ToString(spi_name_cmd.ExecuteScalar());

                    m_cmd.Parameters.Add(new OleDbParameter(":FIO_SPI", cutEnd(Convert.ToString(row["FIO_SPI"]).Trim(), 100).Trim()));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_IP", cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40)));


                    if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// ������� ��� ���� ����������� ��
                    {
                        DatZapr = DateTime.Today;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DatZapr));

                    // � ��� ��� ���� ������ �������������

                    vid_d = 2;
                    if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                    {
                        vid_d = 1;// ���. ����
                    }

                    m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                    m_cmd.Parameters.Add(new OleDbParameter(":INN_D", Convert.ToString(row["INNORG"])));

                    // �������� ���� �� �������, ��� � ��������
                    // ���� ����-�� ���, �� ����� ��� staus = '�����' � text = '�� ������� ���������� �������, ��� � �������� ��������. ����������� ������ �� ����� ���� ��������� ����������. ���������� ��������� ������ � �������� ����.'
                    string txtFIO = Convert.ToString(row["FIOVK"]).Trim();
                    string[] Names;
                    Names = parseFIO(txtFIO);
                    if (!(Names.Length > 2) || (Names[0].Trim().Equals("")) || (Names[1].Trim().Equals("")) || (Names[2].Trim().Equals("")))
                    {
                        txtStatus = "�����";
                        bNotFullFIO = true;
                        txtZaprosText = "'�� ������� ���������� ������� ��� ��� ��� �������� ��������. ����������� ������ �� ����� ���� ��������� ����������. ���������� ��������� ������ � �������� ����.'";
                    }

                    m_cmd.CommandText += ":NAME_D, :DATE_BORN, NULL, NULL, NULL, :FK_DOC, :FK_IP, :FK_ID, NULL, :NUM_ID," + txtZaprosText + ", NULL, NULL, NULL,";
                    
                    m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", cutEnd(txtFIO, 100)));

                    nBirthYear = parseBirthDate(Convert.ToString(row["GOD"]));
                    if (nBirthYear == 0)
                    {
                        txtStatus = "������ � �������";
                        txtTextError = "����������� ���� �������� (������ ##.##.####)";
                    }

                    if (!DateTime.TryParse(Convert.ToString(row["GOD"]), out dtDate))
                    {
                        dtDate = DateTime.MaxValue;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_BORN", dtDate));

                    // ������ ��������� 3 �� �����
                    //m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// ��� ����� ������ �� ������� ���-��

                    //m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// ���� ������, �� ������ �������� ��� ������

                    //3.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ������� �� ���������������, 1 - ��� ���. �� ��������, ������ 1 - ���� ���-� �� ��������) (��� ��� � FIND - ��� ������ 1, ���������� 0)

                    Int32 iKey = -1;
                    //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                    //{
                    //    iKey = 0;
                    //}

                    // TODO: �������, ��� ��� ������� ������ �� DOCUMENTS!!!
                    // ���� ��� � ������� DOCUMENTS ��������� ������, ���� �� �����
                    m_cmd.Parameters.Add(new OleDbParameter(":FK_DOC", iKey));



                    if (!Int32.TryParse(Convert.ToString(row["FK_IP"]), out iKey))
                    {
                        iKey = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":FK_IP", iKey));


                    if (!Int32.TryParse(Convert.ToString(row["FK_ID"]), out iKey))
                    {
                        iKey = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":FK_ID", iKey));


                    DataSet ds_id = new DataSet();
                    DataTable tbl_id = ds_id.Tables.Add("ID");

                    OleDbCommand num_id_cmd = new OleDbCommand("Select NUM_ID, D_ID from ID where ID.PK = " + Convert.ToString(row["FK_ID"]), con, tran);
                    
                    using (OleDbDataReader rdr_id = num_id_cmd.ExecuteReader(CommandBehavior.Default))
                    {
                        ds_id.Load(rdr_id, LoadOption.OverwriteChanges, tbl_id);
                        rdr_id.Close();
                    }
                    
                    String num_id = "";
                    DateTime d_id;

                    if ((ds_id != null) && (ds_id.Tables.Count > 0) && (ds_id.Tables[0] != null) && (ds_id.Tables[0].Rows != null) && (ds_id.Tables[0].Rows.Count > 0))
                    {

                        num_id = Convert.ToString(ds_id.Tables[0].Rows[0]["NUM_ID"]);
                        if (!DateTime.TryParse(Convert.ToString(ds_id.Tables[0].Rows[0]["D_ID"]), out d_id))
                        {
                            d_id = DateTime.MaxValue;
                        }
                    }
                    else d_id = DateTime.MaxValue;

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ID", cutEnd(num_id.Trim(), 30)));

                   
                    m_cmd.CommandText += ":ADRESS, NULL, NULL, :ADRESAT, '" + txtStatus + "', NULL, :TEXT_ERROR, :USCODE, :FK_LEGAL, :CONVENTION, NULL, NULL, NULL, :PASSPORT, :SUMM,";
                    
                    m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250)));

                    //m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// ����� ������ - ������
                    
                    m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", cutEnd(Convert.ToString(Legal_Name_List[0]).Trim(), 200)));
                    //m_cmd.Parameters[":ADRESAT"].Value = Convert.ToString(Legal_Name_List[i]);

                    if (Legal_�onv_List[0].Length > 100)
                    {
                        m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", txtTextError));
                    }
                    else
                    {
                        m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", ""));
                    }

                    if (!(Int32.TryParse(Convert.ToString(row["USCODE"]), out iKey)))
                    {
                        iKey = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":USCODE", iKey));

                    m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));
                    //m_cmd.Parameters[":FK_LEGAL"].Value = Convert.ToInt32(Legal_List[i].Trim());
                    // ��������� ��� ����������� Legal_NameList[0];
                    
                    
                    if (Legal_�onv_List[0].Length > 100)
                    {
                        m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_�onv_List[0].Substring(0, 100)));
                    }
                    else
                    {
                        m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_�onv_List[0].Substring(0, Legal_�onv_List[0].Length)));
                    }

                    
                    // ����� � ������� persons � ��� ���� ������ ������
                    // ������� persons �� ip.fk; � ������� physical �� person.tablename + person.FK
                    // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                    //

                    m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// ����� ������ ����. ������
                    
                    Double sum;
                    if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                    {
                        sum = 0;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                    m_cmd.CommandText += "NULL, NULL, NULL , NULL, :BRANCH_OFFICE, :NUM_SUB_DIV, :H_PRISTAV, :DATE_ID, :FAM, :IM, :OT, :DOLG_BIRTH_YEAR, :TOWN_BORN, NULL, NULL, NULL, NULL, NULL);";
                    
                    m_cmd.Parameters.Add(new OleDbParameter(":BRANCH_OFFICE", legal_branch));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_SUB_DIV", Convert.ToString(osp_num)));

                    m_cmd.Parameters.Add(new OleDbParameter(":H_PRISTAV", osp_h_pristav));

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", d_id));

                    // ���� �������� �� FIOVK �������� �, � � �
                    // 

                    // ��� ��� �� ��� ���� �������
                    //string txtFIO = Convert.ToString(row["FIOVK"]).Trim();
                    //string[] Names;
                    //Names = parseFIO(txtFIO);

                    if (Names.Length > 0) m_cmd.Parameters.Add(new OleDbParameter(":FAM", cutEnd(Convert.ToString(Names[0]), 30)));
                    else m_cmd.Parameters.Add(new OleDbParameter(":FAM", ""));

                    if (Names.Length > 1) m_cmd.Parameters.Add(new OleDbParameter(":IM", cutEnd(Convert.ToString(Names[1]), 30)));
                    else m_cmd.Parameters.Add(new OleDbParameter(":IM", ""));

                    if (Names.Length > 2)
                    {
                        // ��� ��� �������� - ��������. ���������, �������� �� 30 �������� � � ����
                        string txtOt = "";
                        
                        for (int j = 2; j < Names.Length; j++)
                        {
                            txtOt += Names[j] + ' ';
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":OT", cutEnd(txtOt.TrimEnd(), 30)));

                        //if (Names.Length > 3)
                        //{
                        //    MessageBox.Show("�������� ��������� �����! Message:" + txtOt, "��������!", MessageBoxButtons.OK);

                        //}
                    }
                    else m_cmd.Parameters.Add(new OleDbParameter(":OT", ""));

                    if (nBirthYear > 0)
                    {
                        m_cmd.Parameters.Add(new OleDbParameter(":DOLG_BIRTH_YEAR", Convert.ToString(nBirthYear)));
                    }
                    else{
                        m_cmd.Parameters.Add(new OleDbParameter(":DOLG_BIRTH_YEAR", ""));
                    }

                    m_cmd.Parameters.Add(new OleDbParameter(":TOWN_BORN", ""));

                    //String txtNOMIP = Convert.ToString(row["NOMIP"]).Trim();
                    //if (txtNOMIP.Trim() != "")
                    //{
                    //    String[] strings = txtNOMIP.Split('/');
                    //    if (!(Int32.TryParse(Convert.ToString(strings[2]), out iKey)))
                    //    {
                    //        iKey = 0;
                    //    }
                    //}
                    //else
                    //{
                    //    iKey = 0;
                    //}
                    //m_cmd.Parameters.Add(new OleDbParameter(":TEMP", iKey));
                    

                    if (m_cmd.ExecuteNonQuery() == -1) bInserted = false;
                    
                    if (bInserted)
                    {
                        iCnt++;
                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }

                tran.Commit();
                con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    //tran.Rollback();
                    con.Close();
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

            return iCnt;
        }


        private int InsertZapros_kred_org(DataTable dtTable, DateTime DatZapr1_param, DateTime DatZapr2_param, bool bRegTable)
        {
            int vid_d = 1;
            DateTime dtDate;
            DateTime DatZapr;
            String uid = System.Guid.NewGuid().ToString();
            String osp_name = GetOSP_Name();
            uid = cutEnd(uid, 30);
            Int32 iCnt = 0;
            Int32 iBadInnCnt = 0;
            Int32 iBadYearBorn = 0;
            OleDbTransaction tran;
            bool bJurBadInn = false;
            bool bFizBadYear = false;

            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                osp_name = cutEnd(osp_name, 100);
                foreach (DataRow row in dtTable.Rows)
                {
                    string txtLitzdolg, txtInnd, txtZapros, txtFIO_SPI, txtName_d, txtNum_Id, txtAddr, txtNOMIP;
                    Int32 iFK_DOC = -1;
                    Int32 iFK_IP, iFK_ID, iUSCODE;
                    Double sum;
                    bool bInserted = true;
                    Int32 iNPP = 0;
                    string txtText = "";
                    int nResult= 0;
                    string txtStatus = "���������";
                    bJurBadInn = false;
                    bFizBadYear = false;
                    int nBirthYear = 0;
                    
                    for (int i = 1; i < Legal_List.Length; i++)
                    {
                        // ��� doc ��� ���� ����� ��������� ���� �� ��� � ��. ���� � ���� ��� ��
                        // �������� ������, result � text
                        // ��� �������� � DBF ��������� �������� ���-��
                        // � Insert �������� ���� TEXT, RESULT, (STATUS ��� ����)
                        // NULL, 0 ��-���������

                        txtText = "";
                        nResult = 0;
                        nBirthYear = 0;
                        txtStatus = "���������";
                        
                        m_cmd = new OleDbCommand();
                        m_cmd.Connection = con;
                        m_cmd.Transaction = tran;
                        m_cmd.CommandText = "INSERT INTO ZAPROS (PK, NUM_ZAPR, SUB_DIV, FIO_SPI, NUM_IP, DATE_ZAPR, VID_D, INN_D, NAME_D, DATE_R, NUM_RES, DATE_RES, RESULT, FK_DOC, FK_IP, FK_ID, NUM_ID, TEXT, DATE_BEG, DATE_END, ADRESS, NUM_PACK, NUM_ZAPR_IN_PACK, STATUS, DATE_SEND, TEXT_ERROR, USCODE, CONVENTION, WHYRESPONS, WHYPREPARE, PASSPORT, SUMM, TEMP, FK_LEGAL, ADRESAT)";
                        m_cmd.CommandText += " VALUES(GEN_ID(S_N_ZAPROS, 1), :NUM_ZAPR, :SUB_DIV, :FIO_SPI, :NUM_IP, :DATE_ZAPR, :VID_D, :INN_D, :NAME_D, :DATE_R, :NUM_RES, :DATE_RES, :RESULT, :FK_DOC, :FK_IP, :FK_ID, :NUM_ID, :TEXT, :DATE_BEG, :DATE_END, :ADRESS, :NUM_PACK, :NUM_ZAPR_IN_PACK, :STATUS, :DATE_SEND, :TEXT_ERROR, :USCODE, :CONVENTION, :WHYRESPONS, :WHYPREPARE, :PASSPORT, :SUMM, :TEMP, :FK_LEGAL, :ADRESAT)";

                        txtZapros = cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40);
                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR", txtZapros));

                        m_cmd.Parameters.Add(new OleDbParameter(":SUB_DIV", osp_name));

                        //OleDbCommand spi_name_cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE USCODE = '" + Convert.ToString(row["NOMSPI"]) + "'", con, tran);
                        //String spi_name = Convert.ToString(spi_name_cmd.ExecuteScalar());
                        
                        txtFIO_SPI = cutEnd(Convert.ToString(row["FIO_SPI"]).Trim(), 100);
                        
                        m_cmd.Parameters.Add(new OleDbParameter(":FIO_SPI", txtFIO_SPI));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_IP", txtZapros));


                        if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// ������� ��� ���� ����������� ��
                        {
                            DatZapr = DateTime.Today;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DatZapr));

                        // � ��� ��� ���� ������ �������������

                        vid_d = 2;
                        txtLitzdolg = Convert.ToString(row["LITZDOLG"]).Trim();
                        if (txtLitzdolg.StartsWith("/1/"))
                        {
                            vid_d = 1;// ���. ����
                        }

                        m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                        txtInnd = Convert.ToString(row["INNORG"]).Trim();
                        // ���� ��� REG � ��. ����
                        if (bRegTable && (vid_d == 2))
                        {
                            if (txtInnd.Trim().Length < 10)
                            {
                                // ��� ��� � ������ ��� ��� ����� ����� �������� �������� � ��� ��� � ���� ���� � ��� � ������ �� �����
                                nResult = 1; // � ����� � ��� ��� ����� ����
                                txtText = "������! � �������� - ��.���� �� �������� ���. ����������� ������ �� ����� ���� ���������. ��� �������� ������������ ������� ���������� ��������� ���� ��� � ������� ������ �����.\n";
                                txtStatus = "������ � �������";
                                bJurBadInn = true;
                            }
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":INN_D", txtInnd));

                        txtName_d = cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 100);
                        m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", txtName_d));

                        // TODO: ������� �������� �� ������������ ���� �������� � ���� ������ - �� ������ ������ � ����
                        if (vid_d == 1) // �������� ������ ��� ���. ���
                        {
                            nBirthYear = parseBirthDate(Convert.ToString(row["GOD"]));
                            if (nBirthYear == 0)
                            {
                                nResult = 1; // � ����� � ��� ��� ����� ����
                                txtStatus = "������ � �������";
                                txtText = "����������� ��������� ��������� ���� ��� ��� �������� (������ ##.##.#### ��� ####)";
                                bFizBadYear = true;
                            }
                        }

                        if (!DateTime.TryParse(Convert.ToString(row["GOD"]), out dtDate))
                        {
                            dtDate = DateTime.MaxValue;
                        }

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// ��� ����� ������ �� ������� ���-��

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// ���� ������, �� ������ �������� ��� ������

                        // m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ������� �� ���������������, 1 - ��� ���. �� ��������, ������ 1 - ���� ���-� �� ��������) (��� ��� � FIND - ��� ������ 1, ���������� 0)

                        m_cmd.Parameters.Add(new OleDbParameter(":RESULT", nResult));

                        //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                        //{
                        //    iKey = 0;
                        //}

                        // TODO: �������, ��� ��� ������� ������ �� DOCUMENTS!!!
                        // ���� ��� � ������� DOCUMENTS ��������� ������, ���� �� �����

                        m_cmd.Parameters.Add(new OleDbParameter(":FK_DOC", iFK_DOC));

                        if (!Int32.TryParse(Convert.ToString(row["FK_IP"]), out iFK_IP))
                        {
                            iFK_IP = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":FK_IP", iFK_IP));


                        if (!Int32.TryParse(Convert.ToString(row["FK_ID"]), out iFK_ID))
                        {
                            iFK_ID = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":FK_ID", iFK_ID));


                        OleDbCommand num_id_cmd = new OleDbCommand("Select NUM_ID from ID where ID.PK = " + Convert.ToString(row["FK_ID"]), con, tran);
                        txtNum_Id = cutEnd(Convert.ToString(num_id_cmd.ExecuteScalar()).Trim(), 30);

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ID", txtNum_Id));

                        m_cmd.Parameters.Add(new OleDbParameter(":TEXT", txtText));// ����� ������ ��-��������� ������

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", DatZapr1_param));
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", DatZapr2_param));

                        txtAddr = cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250);
                        m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", txtAddr));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", uid));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                        //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                        m_cmd.Parameters.Add(new OleDbParameter(":STATUS", txtStatus));

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                        if (Legal_�onv_List[i].Length > 100)
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", Legal_�onv_List[i].Substring(100)));
                        }
                        else
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", txtText)); // ���� ��� ������ �� ����� ������, � ���� ���� �� ����� text
                        }

                        if (!(Int32.TryParse(Convert.ToString(row["USCODE"]), out iUSCODE)))
                        {
                            iUSCODE = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":USCODE", iUSCODE));

                        //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));

                        if (Legal_�onv_List[i].Length > 100)
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_�onv_List[i].Substring(0, 100)));
                        }
                        else
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_�onv_List[i].Substring(0, Legal_�onv_List[i].Length)));
                        }

                        m_cmd.Parameters.Add(new OleDbParameter(":WHYRESPONS", ""));

                        m_cmd.Parameters.Add(new OleDbParameter(":WHYPREPARE", ""));

                        // ����� � ������� persons � ��� ���� ������ ������
                        // ������� persons �� ip.fk; � ������� physical �� person.tablename + person.FK
                        // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                        //

                        m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// ����� ������ ����. ������

                        if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                        {
                            sum = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));


                        txtNOMIP = Convert.ToString(row["NOMIP"]).Trim();
                        iNPP = 0;
                        if (txtNOMIP.Trim() != "")
                        {
                            String[] strings = txtNOMIP.Split('/');
                            if (!(Int32.TryParse(Convert.ToString(strings[2]), out iNPP)))
                            {
                                iNPP = 0;
                            }
                        }
                        else
                        {
                            iNPP = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":TEMP", iNPP));


                        m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[i].Trim())));
                        //m_cmd.Parameters[":FK_LEGAL"].Value = Convert.ToInt32(Legal_List[i].Trim());
                        m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", cutEnd(Convert.ToString(Legal_Name_List[i]).Trim(), 200)));
                        //m_cmd.Parameters[":ADRESAT"].Value = Convert.ToString(Legal_Name_List[i]);

                        //m_cmd.Parameters.Add(new OleDbParameter(":FILENAME", cutEnd(Convert.ToString(row["LITZDOLG"]).Trim(), 30)));

                        if (m_cmd.ExecuteNonQuery() == -1) bInserted = false;

                        // �������� ������ �� ������� ���������������� ��� �� ��.����� ���� ���� ���
                        if ((txtLitzdolg.Equals("/1/5/")) && (txtInnd.Trim().Length >= 10))
                        {

                            m_cmd = new OleDbCommand();
                            m_cmd.Connection = con;
                            m_cmd.Transaction = tran;
                            m_cmd.CommandText = "INSERT INTO ZAPROS (PK, NUM_ZAPR, SUB_DIV, FIO_SPI, NUM_IP, DATE_ZAPR, VID_D, INN_D, NAME_D, DATE_R, NUM_RES, DATE_RES, RESULT, FK_DOC, FK_IP, FK_ID, NUM_ID, TEXT, DATE_BEG, DATE_END, ADRESS, NUM_PACK, NUM_ZAPR_IN_PACK, STATUS, DATE_SEND, TEXT_ERROR, USCODE, CONVENTION, WHYRESPONS, WHYPREPARE, PASSPORT, SUMM, TEMP, FK_LEGAL, ADRESAT)";
                            m_cmd.CommandText += " VALUES(GEN_ID(S_N_ZAPROS, 1), :NUM_ZAPR, :SUB_DIV, :FIO_SPI, :NUM_IP, :DATE_ZAPR, :VID_D, :INN_D, :NAME_D, :DATE_R, :NUM_RES, :DATE_RES, :RESULT, :FK_DOC, :FK_IP, :FK_ID, :NUM_ID, :TEXT, :DATE_BEG, :DATE_END, :ADRESS, :NUM_PACK, :NUM_ZAPR_IN_PACK, :STATUS, :DATE_SEND, :TEXT_ERROR, :USCODE, :CONVENTION, :WHYRESPONS, :WHYPREPARE, :PASSPORT, :SUMM, :TEMP, :FK_LEGAL, :ADRESAT)";

                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR", txtZapros));

                            m_cmd.Parameters.Add(new OleDbParameter(":SUB_DIV", osp_name));

                            m_cmd.Parameters.Add(new OleDbParameter(":FIO_SPI", txtFIO_SPI));

                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_IP", txtZapros));

                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DatZapr));

                            // � ��� ��� ���� ������ �������������

                            vid_d = 2; // ������ � �� = ��� ��.����
                            m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                            m_cmd.Parameters.Add(new OleDbParameter(":INN_D", txtInnd));

                            m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", txtName_d));

                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// ��� ����� ������ �� ������� ���-��

                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// ���� ������, �� ������ �������� ��� ������

                            m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ������� �� ���������������, 1 - ��� ���. �� ��������, ������ 1 - ���� ���-� �� ��������) (��� ��� � FIND - ��� ������ 1, ���������� 0)

                            m_cmd.Parameters.Add(new OleDbParameter(":FK_DOC", iFK_DOC));

                            m_cmd.Parameters.Add(new OleDbParameter(":FK_IP", iFK_IP));

                            m_cmd.Parameters.Add(new OleDbParameter(":FK_ID", iFK_ID));

                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_ID", txtNum_Id));

                            m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// ����� ������ - ������

                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", DatZapr1_param));
                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", DatZapr2_param));

                            m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", txtAddr));

                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", uid));

                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                            //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                            m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "���������"));

                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                            if (Legal_�onv_List[i].Length > 100)
                            {
                                m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", Legal_�onv_List[i].Substring(100)));
                            }
                            else
                            {
                                m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", ""));
                            }

                            m_cmd.Parameters.Add(new OleDbParameter(":USCODE", iUSCODE));

                            //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));

                            if (Legal_�onv_List[i].Length > 100)
                            {
                                m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_�onv_List[i].Substring(0, 100)));
                            }
                            else
                            {
                                m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_�onv_List[i].Substring(0, Legal_�onv_List[i].Length)));
                            }

                            m_cmd.Parameters.Add(new OleDbParameter(":WHYRESPONS", ""));

                            m_cmd.Parameters.Add(new OleDbParameter(":WHYPREPARE", ""));

                            // ����� � ������� persons � ��� ���� ������ ������
                            // ������� persons �� ip.fk; � ������� physical �� person.tablename + person.FK
                            // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                            //

                            m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// ����� ������ ����. ������

                            
                            m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                            m_cmd.Parameters.Add(new OleDbParameter(":TEMP", iNPP));


                            m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[i].Trim())));
                            //m_cmd.Parameters[":FK_LEGAL"].Value = Convert.ToInt32(Legal_List[i].Trim());
                            m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", cutEnd(Convert.ToString(Legal_Name_List[i]).Trim(), 200)));
                            //m_cmd.Parameters[":ADRESAT"].Value = Convert.ToString(Legal_Name_List[i]);

                            //m_cmd.Parameters.Add(new OleDbParameter(":FILENAME", cutEnd(Convert.ToString(row["LITZDOLG"]).Trim(), 30)));

                            if (m_cmd.ExecuteNonQuery() == -1) bInserted = false;
                        }
                    }

                    if (bInserted)
                    {
                        if (bJurBadInn) { iBadInnCnt++; }
                        else if (bFizBadYear) { iBadYearBorn++; }
                        else
                        {
                            iCnt++;
                        }
                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }

                tran.Commit();
                con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    //tran.Rollback();
                    con.Close();
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

            return iCnt;
        }




        private void ReadPOTDData(DateTime dat1, DateTime dat2)
        {
            // ��������� ������������ ������������ �� ������ ��� � ����������� � �������� ������ ��������������� �� ����������� ������
            //DT_ktfoms = GetDataTableFromFB("SELECT DISTINCT a.USCODE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD,  a.name_d as FIOVK, a.DATE_BORN_D as DATROZHD, a.ADR_D as ADDR, b.PK as FK_DOC, a.PK as FK_IP, a.PK_ID as FK_ID, b.KOD, b.DATE_DOC FROM IP a LEFT JOIN DOCUMENT b ON b.FK = a.PK WHERE a.DATE_IP_OUT is not null and a.text_pp is not null and ((a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') or (b.DATE_DOC >= '" + dat1.ToShortDateString() + "' AND b.DATE_DOC <= '" + dat2.ToShortDateString() + "' AND b.KOD = 1010))  AND a.VIDD_KEY LIKE '/1/%'", "TOFIND");
            //DT_pens_reg = GetDataTableFromFB("SELECT DISTINCT a.USCODE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD, a.name_d as FIOVK, a.DATE_BORN_D as DATROZHD, a.ADR_D as ADDR, a.PK as FK_IP, a.PK_ID as FK_ID FROM IP a LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.DATE_IP_OUT is null and b.KOD != 1006 and (a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' AND a.NUM_IP NOT LIKE '%!%'", "TOFIND and a.NUM_IP not in (select a.NUM_IP from IP a LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.DATE_IP_OUT is null and b.KOD = 1006 and (b.DATE_DOC >= '" + dat1.ToShortDateString() + "' AND b.DATE_DOC <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%')");
            DT_potd_doc = GetDataTableFromFB("SELECT DISTINCT c.PRIMARY_SITE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD, UPPER(a.name_d) as FIOVK, a.DATE_BORN_D as DATROZHD, a.ADR_D as ADDR, b.PK as FK_DOC, a.PK as FK_IP, a.PK_ID as FK_ID, a.uscode, a.FIO_SPI FROM IP a left join s_users c on (a.uscode=c.uscode) LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.DATE_IP_OUT is null and b.KOD = 1045 and (b.DATE_DOC >= '" + dat1.ToShortDateString() + "' AND b.DATE_DOC <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' and a.ssd is null and a.ssv is null order by FIOVK", "TOFIND");
        }

        private void ReadKRCData(DateTime dat1, DateTime dat2)
        {            
            //DT_krc_reg = GetDataTableFromFB("select a.sdc as nomotd, (select first 1 c.PRIMARY_SITE from s_users c where c.uscode = a.uscode) as nomspi, a.num_ip as zapros, a.num_id as nomid, a.date_id_send as datid, a.name_d as namedol, a.DATE_BORN_D as born, a.sum_ as sumvz, a.adr_d as addr from ip a WHERE a.DATE_IP_OUT is null and a.name_v like '%���%' and (a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' and a.ssd is null and a.ssv is null", "TOFIND");
            //DT_krc_reg = GetDataTableFromFB("select distinct c.PRIMARY_SITE as NOMSPI, a.sdc as nomotd, a.num_ip as zapros, a.num_id as nomid, a.date_id_send as datid, a.name_d as namedol, a.DATE_BORN_D as born, a.sum_ as sumvz, a.adr_d as addr, a.PK as FK_IP, a.PK_ID as FK_ID, a.uscode, a.FIO_SPI from ip a left join s_users c on (a.uscode=c.uscode) WHERE a.DATE_IP_OUT is null and a.name_v like '%���%' and (a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' and a.ssd is null and a.ssv is null", "TOFIND");

            DT_krc_reg = GetDataTableFromFB("select distinct c.PRIMARY_SITE as NOMSPI, a.sdc as nomotd, a.num_ip as zapros, a.num_id as nomid, a.date_id_send as datid, a.name_d as namedol, a.DATE_BORN_D as born, a.sum_ as sumvz, a.adr_d as addr, a.PK as FK_IP, a.PK_ID as FK_ID, a.uscode, a.FIO_SPI from ip a left join s_users c on (a.uscode=c.uscode) WHERE a.DATE_IP_OUT is null and (a.name_v like '%���%' or a.name_v = '����������� ��������� ����� ���') and a.VIDD_KEY LIKE '/1/%' and a.ssd is null and a.ssv is null", "TOFIND");
        }

        private bool ReestrOutWord(DataTable dtReg, string dir)
        {
            Decimal nYear = DateTime.Today.Year;
            DateTime dtDate;
            string bankname = "���������� ��� N 8628 �� �� ��";
            //string bankadres = "�.������������, ��.�����������, �.2";
            string bankadres = "";
            string ospadres = GetOSP_Adres().ToUpper();
            string ospname = GetOSP_Name().ToUpper();

            DataRow[] FizRows = dtReg.Select("LITZDOLG LIKE '/1/*'", "FIOVK");

            if (File.Exists(string.Format(@"{0}\{1}", dir, DateTime.Today.ToShortDateString() + ".doc")))
                File.Delete(string.Format(@"{0}\{1}", dir, DateTime.Today.ToShortDateString() + ".doc"));

            using (StreamWriter sw = new StreamWriter(dir + "\\" + DateTime.Today.ToShortDateString() + ".doc", true, Encoding.GetEncoding(1251)))
            {

                sw.WriteLine("              ������                      " + bankname);
                sw.WriteLine("   ����������� ������ �������� ���������  " + bankadres);
                sw.WriteLine("   ���������� ���� �� ���������� �������  ");
                sw.WriteLine("");
                sw.WriteLine("  " + ospname);
                sw.WriteLine("  ");
                sw.WriteLine("  " + ospadres);
                sw.WriteLine("");
                sw.WriteLine("    ���.N ________�� _________   ;");
                sw.WriteLine("");
                sw.WriteLine("                                � � � � � �");
                sw.WriteLine("   �� ���������� � " + ospname);
                sw.WriteLine("   ��������� �������������� ��������� �� ��������� : ");
                foreach (DataRow row in FizRows)
                {
                    nYear = 0;
                    if (DateTime.TryParse(Convert.ToString(row["GOD"]), out dtDate))
                    {
                        nYear = dtDate.Year;
                    }
                    sw.WriteLine("   " + Convert.ToString(row["ZAPROS"]) + ", " + Convert.ToString(row["FIOVK"]) + ", " + Convert.ToString(row["ADDR"]) + ", " + Convert.ToString(nYear));
                }

                sw.WriteLine("  ");
                sw.WriteLine("  ");
                sw.WriteLine("  ������ ��� � ����������� ���� �������� � ������ ���������,");
                sw.WriteLine("  ������������������ � ����� �����.");
                sw.WriteLine("  ");
                sw.WriteLine("  ");
                sw.WriteLine("  �����������:  ");

                sw.Flush();
                sw.Close();
            }

            using (StreamReader sr = new StreamReader(dir + "\\" + DateTime.Today.ToShortDateString() + ".doc", Encoding.GetEncoding(1251)))
            {

                // ������ ��� �����

                Microsoft.Office.Interop.Word.ApplicationClass app = new ApplicationClass();

                object s1 = "";
                object fl = false;
                object t = WdNewDocumentType.wdNewBlankDocument;
                object fl2 = true;

                Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref s1, ref fl, ref t, ref fl2);

                Paragraph par = doc.Content.Paragraphs[1];

                par.Range.Font.Name = "Courier";
                par.Range.Font.Size = 8;
                float a = par.Range.PageSetup.RightMargin;
                float b = par.Range.PageSetup.LeftMargin;
                float c = par.Range.PageSetup.TopMargin;
                float d = par.Range.PageSetup.BottomMargin;

                par.Range.PageSetup.RightMargin = 30;
                par.Range.PageSetup.LeftMargin = 30;
                par.Range.PageSetup.TopMargin = 20;
                par.Range.PageSetup.BottomMargin = 20;

                par.Range.Text = sr.ReadToEnd();
                app.Visible = true;
            }
            return true;
        }


        private bool ReestrOut(DataTable dtReg, string dir)
        {
            Decimal nYear = DateTime.Today.Year;
            DateTime dtDate;
            string bankname = "���������� ��� N 8628 �� �� ��";
            //string bankadres = "�.������������, ��.�����������, �.2";
            string bankadres = "";
            string ospadres = GetOSP_Adres().ToUpper();
            string ospname = GetOSP_Name().ToUpper();

            DataRow[] FizRows = dtReg.Select("LITZDOLG LIKE '/1/*'", "FIOVK");

            if (File.Exists(string.Format(@"{0}\{1}", dir, DateTime.Today.ToShortDateString() + ".doc")))
                File.Delete(string.Format(@"{0}\{1}", dir, DateTime.Today.ToShortDateString() + ".doc"));

            using (StreamWriter sw = new StreamWriter(dir + "\\" + DateTime.Today.ToShortDateString() + ".doc", true, Encoding.GetEncoding(1251)))
            {

                sw.WriteLine("              ������                      " + bankname);
                sw.WriteLine("   ����������� ������ �������� ���������  " + bankadres);
                sw.WriteLine("   ���������� ���� �� ���������� �������  ");
                sw.WriteLine("");
                sw.WriteLine("  " + ospname);
                sw.WriteLine("  ");
                sw.WriteLine("  " + ospadres);
                sw.WriteLine("");
                sw.WriteLine("    ���.N ________�� _________   ;");
                sw.WriteLine("");
                sw.WriteLine("                                � � � � � �");
                sw.WriteLine("   �� ���������� � " + ospname);
                sw.WriteLine("   ��������� �������������� ��������� �� ��������� : ");
                foreach (DataRow row in FizRows)
                {
                    nYear = 0;
                    if (DateTime.TryParse(Convert.ToString(row["GOD"]), out dtDate))
                    {
                        nYear = dtDate.Year;
                    }
                    sw.WriteLine("   " + Convert.ToString(row["ZAPROS"]) + ", " + Convert.ToString(row["FIOVK"]) + ", " + Convert.ToString(row["ADDR"]) + ", " + Convert.ToString(nYear));
                }

                sw.WriteLine("  ");
                sw.WriteLine("  ");
                sw.WriteLine("  ������ ��� � ����������� ���� �������� � ������ ���������,");
                sw.WriteLine("  ������������������ � ����� �����.");
                sw.WriteLine("  ");
                sw.WriteLine("  ");
                sw.WriteLine("  �����������:  ");

                sw.Flush();
                sw.Close();
                //report = sw.ToString();

                Process proc = new Process();
                //proc.StartInfo.FileName = "winword.exe";

                proc.StartInfo.FileName = dir + "\\" + DateTime.Today.ToShortDateString() + ".doc";
                //proc.StartInfo.Arguments = dir + "\\" + DateTime.Today.ToShortDateString() + ".doc";
                proc.Start();
            }
            return true;

        }

        private Decimal GetOSP_Num()
        {
            Decimal res = 0;
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                //OleDbCommand cmd = new OleDbCommand("Select DEPARTMENT from OSP", con, tran);
                OleDbCommand cmd = new OleDbCommand("select osp.department from system_site left join osp on osp.osp_system_site_id = system_site.system_site_id", con, tran);
                res = Convert.ToDecimal(cmd.ExecuteScalar());
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



        private String GetSpiName(decimal USCODE, OleDbConnection conn)
        {
            String res = "";
            try
            {
                if (conn.State == ConnectionState.Open)
                {
                    OleDbCommand cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE USCODE = '" + Convert.ToString((int)USCODE) + "'", conn);
                    res = Convert.ToString(cmd.ExecuteScalar());
                }
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

        private String GetSpiName2(decimal USCODE)
        {
            String res = "";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();

                con.Open();

                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                //OleDbCommand cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE USCODE = '" + Convert.ToString((int)USCODE) + "'", FBcon, tran);
                OleDbCommand cmd = new OleDbCommand("select suser_fio from spi left join sys_users on spi.suser_id = sys_users.suser_id where spi.SPI_ZONENUM = " + Convert.ToString((int)USCODE), con, tran);
                res = Convert.ToString(cmd.ExecuteScalar());
                tran.Rollback();
                con.Close();

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", System.Windows.Forms.MessageBoxButtons.OK);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", System.Windows.Forms.MessageBoxButtons.OK);
            }
            return res;
        }

        private string PKOSP_GetOrgConvention(decimal org_id){
            string res = "< ��� �������� � ���� ������ >";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();

                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("select head_post from entity where entity.entt_id = " + Convert.ToString(org_id), con, tran);
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

        private String PK_OSP_GetSPI_Name(Decimal code)
        {
            String res = "��� �������� � ���� ������";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();

                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("select suser_fio from spi left join sys_users on spi.suser_id = sys_users.suser_id where spi.SPI_ZONENUM =  " + Convert.ToString(code), con, tran);
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

        private String GetSpiName3(decimal PRIMARY_SITE)
        {
            String res = "";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();

                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE PRIMARY_SITE = '" + Convert.ToString((int)PRIMARY_SITE) + "'", con, tran);
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

        private String GetOSP_H_Pristav()
        {
            String res = "";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select first 1 H_PRISTAV from S_SUBDIVIDINGS", con, tran);
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

        private ArrayList GetLoadedReestrs(OleDbConnection ConParam){
            
            OleDbTransaction tran = null;
            DataSet dsReestr_params;
            DataTable dtReestr_params;
            ArrayList result = new ArrayList();

            dsReestr_params = new DataSet();
            dtReestr_params = dsReestr_params.Tables.Add("Reestr_params");
            

            try
            {

                if ((ConParam == null) || (ConParam.State == ConnectionState.Closed))
                    ConParam.Open();

                tran = ConParam.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmdIP = new OleDbCommand();
                cmdIP.Connection = ConParam;
                cmdIP.Transaction = tran;
                // ���������� select ��-������
                cmdIP.CommandText = "select distinct ish_number from GIBDD_PLATEZH";
                using (OleDbDataReader rdr = cmdIP.ExecuteReader(CommandBehavior.Default))
                {
                    dsReestr_params.Load(rdr, LoadOption.OverwriteChanges, dtReestr_params);
                    rdr.Close();
                }

                tran.Rollback();
                ConParam.Close();

                foreach (DataRow dataRow in dtReestr_params.Rows)
                {
                    result.Add(Convert.ToString(dataRow[0]));
                }

                return result;
                
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
            ConParam.Close();
            return result;
        }
        
        private DataTable GetPackParams(OleDbConnection con, Decimal nZaprosID, Decimal nContrID)
        {
            OleDbTransaction tran = null;
            DataSet dsIP_params;
            DataTable dtIP_params;

            dsIP_params = new DataSet();
            dtIP_params = dsIP_params.Tables.Add("IP_params");

            try
            {

                if((con == null) || (con.State == ConnectionState.Closed))
                    con.Open();

                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmdIP = new OleDbCommand();
                cmdIP.Connection = con;
                cmdIP.Transaction = tran;
                // ���������� select ��-������
                //cmdIP.CommandText = "select s.dx_pack_id, s.agent_id, s.agreement_id, s.agent_dept_id from sendlist s where sendlist_o_id = :ZAPROS_ID and SENDLIST_CONTR = :CONTR_ID";
                cmdIP.CommandText = "select d_p.id as dx_pack_id, d_p.agent_id, d_p.agent_dept_id, d_p.agreement_id, d_p.agent_code, d_p.agent_dept_code, d_p.agreement_code from SENDLIST_DBT_REQUEST_TYPE s_d_r_t join dx_pack d_p on s_d_r_t.outer_agreement_id = d_p.agreement_id join ext_request e_r on e_r.pack_id = d_p.id  where e_r.req_id = :ZAPROS_ID and s_d_r_t.SNDL_CONTR_ID = :CONTR_ID";
                cmdIP.Parameters.Add(new OleDbParameter(":ZAPROS_ID", Convert.ToDecimal(nZaprosID)));
                cmdIP.Parameters.Add(new OleDbParameter(":CONTR_ID", Convert.ToDecimal(nContrID)));
                using (OleDbDataReader rdr = cmdIP.ExecuteReader(CommandBehavior.Default))
                {
                    dsIP_params.Load(rdr, LoadOption.OverwriteChanges, dtIP_params);
                    rdr.Close();
                }

                tran.Rollback();
                con.Close();

                return dsIP_params.Tables[0];
                
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
            con.Close();
            return null;
        }

        private bool FindSendlist(decimal nID, decimal nSendlistContr)
        {
            Decimal res = 0;
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();

                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("select sendlist_o_id from sendlist where sendlist_o_id = " + nID.ToString() + " and SENDLIST_CONTR = " + nSendlistContr.ToString(), con, tran);
                res = Convert.ToDecimal(cmd.ExecuteScalar());
                tran.Rollback();
                con.Close();
                if ((res == nID) && (nID != 0))
                {
                    return true;
                }
                else return false;
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
            return false;
        }

        private bool FindZapros(Decimal nID){
            Decimal res = 0;
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select ID from document where id = " + nID.ToString(), con, tran);
                res = Convert.ToDecimal(cmd.ExecuteScalar());
                tran.Rollback();
                con.Close();
                if ((res == nID) && (nID != 0))
                {
                    return true;
                }
                else return false;
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
            return false;
        }

        private decimal GetAgentDept_ID(Decimal nCode)
        {
            decimal res = 0;
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("select first 1 AGENT_DEPT_ID from mvv_agent_agreement agr join mvv_agent_dept agent on agr.agent_dept_id = agent.exad_id where agr.id = " + nCode.ToString(), con, tran);
                res = Convert.ToDecimal(cmd.ExecuteScalar());
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

        private String GetAgentDept_Code(Decimal nCode)
        {
            String res = "";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("select first 1 DEPARTAMENT_CODE from mvv_agent_agreement agr join mvv_agent_dept agent on agr.agent_dept_id = agent.exad_id where agr.id = " + nCode.ToString(), con, tran);
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

        private Decimal GetAgr_by_Org(Decimal nOrgCode)
        {
            Decimal res = -1;
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("select outer_agreement_id from sendlist_dbt_request_type where sndl_contr_id = " + nOrgCode.ToString(), con, tran);
                res = Convert.ToDecimal(cmd.ExecuteScalar());
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

        private decimal GetAgent_ID(Decimal nCode)
        {
            decimal res = 0;
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("select first 1 AGENT_ID from mvv_agent_agreement agr join mvv_agent agent on agr.agent_id = agent.id where agr.id = " + nCode.ToString(), con, tran);
                res = Convert.ToDecimal(cmd.ExecuteScalar());
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

        private String GetAgent_Code(Decimal nCode)
        {
            String res = "";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("select first 1 ORGANIZATION_CODE from mvv_agent_agreement agr join mvv_agent agent on agr.agent_id = agent.id where agr.id = " + nCode.ToString(), con, tran);
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

        private String GetAgreement_Code(Decimal nAgreementCode)
        {
            String res = "";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("select first 1 AGREEMENT_CODE from mvv_agent_agreement agr where agr.id = " + nAgreementCode.ToString(), con, tran);
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

        private String GetOSP_Name()
        {
            String res = "";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("select first 1 osp.div_name || ' ' || osp.to_name from system_site left join osp on osp.osp_system_site_id = system_site.system_site_id", con, tran);
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

        private String GetOSP_Adres()
        {
            String res = "";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select first 1 P_ADDRESS from S_SUBDIVIDINGS", con, tran);
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

        private String GetLegal_Branch(int code)
        {
            String res = "";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select PODR from LEGAL WHERE PK = " + Convert.ToString(code), con, tran);
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



        private String GetLegal_Name(int code)
        {
            String res = "��� �������� � ���� ������";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select ENTT_FULL_NAME from ENTITY WHERE ENTT_ID = " + Convert.ToString(code), con, tran);
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


        private String GetLegal_Conv(int code)
        {
            String res = "��� �������� � ���� ������";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select HEAD_POST from ENTITY WHERE ENTT_ID = " + Convert.ToString(code), con, tran);
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


        private string GetIPNum(string txtCode)
        {
            
            Decimal code;
            if(!Decimal.TryParse(txtCode, out code)){
                code = -1;
            }
            string res = "";
            try
            {
                if (code != -1)
                {
                    if (con != null && con.State != ConnectionState.Closed) con.Close();
                    con.Open();
                    OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                    OleDbCommand cmd = new OleDbCommand("Select ipno from O_IP WHERE ID = " + Convert.ToString(code), con, tran);
                    res = Convert.ToString(cmd.ExecuteScalar());
                    tran.Rollback();
                    con.Close();
                }
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


        private String GetLegal_Name(Decimal code)
        {
            String res = "��� �������� � ���� ������";
            try
            {
                if(con != null && con.State.Equals(ConnectionState.Closed)) con.Open();
                
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select ENTT_FULL_NAME from ENTITY WHERE ENTT_ID = " + Convert.ToString(code), con, tran);
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

        private string GetLegalNameByAgrCode(string txtAgreementCode){
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
        


        private String GetLegal_Adr(Decimal code)
        {
            String res = "��� �������� � ���� ������";
            try
            {
                if (con != null && con.State.Equals(ConnectionState.Closed)) con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select ENTT_ADDRESS from ENTITY WHERE ENTT_ID = " + Convert.ToString(code), con, tran);
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
        
        private String GetLegal_Conv(Decimal code)
        {
            String res = "��� �������� � ���� ������";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("Select HEAD_POST from ENTITY WHERE ENTT_ID = " + Convert.ToString(code), con, tran);
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


        
        
        private string[] parseFIO(string txtFIO)
        {
            string[] Names;

            try
            {
                // ���� �������� �� FIOVK �������� �, � � �
                int i = 0;
                if (txtFIO.Trim() != "")
                {
                while (txtFIO.IndexOf("  ") != -1)
                    {
                        txtFIO = txtFIO.Replace("  ", " ");
                        i++;
                        if (i > 200)
                        {
                            break;
                        }
                    }
                    Names = txtFIO.Split(' ');
                }
                else
                {
                    Names = new string[] { "", "", "" };
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ��� ������� ������� ������ ��� �� 3 �����. Message: " + ex.Message + "Source: " + ex.Source, "��������!", MessageBoxButtons.OK);
                Names = new string[] { "", "", "" };
            }

            return Names;
            
        }

        private int parseBirthDate(string txtDateBornD)
        {

            int nYear = 0;
            DateTime dtDate;
            
            try
            {
                txtDateBornD = txtDateBornD.TrimEnd();
                if (DateTime.TryParse(txtDateBornD, out dtDate))
                {
                    nYear = dtDate.Year;
                }
                else
                {
                    // ����� ������, ���� ������� � ���������� ����� �� �������������, �� �������� ��� ����� 
                    // � ���� ���������� 4 ����� � ��� ������ � �������� 1900 - 9999
                    
                    // �������� �� ������ ��� �����, ����������� ������������� ��� ���
                    string[] strData = txtDateBornD.Split('.');
                    txtDateBornD = "";
                    foreach (string str in strData)
                    {
                        txtDateBornD += str.Trim();
                    }

                    // �������� ���� � ������ ������ - ��� ��� ���� ��� ��� ��������, �� ��� ��� �� ����� - ��� ����������
                    txtDateBornD = txtDateBornD.TrimStart('0');

                    if (txtDateBornD.Length == 4)
                    {
                        MatchCollection myMatchColl = Regex.Matches(txtDateBornD, @"\b[12]\d\d\d");
                        if (myMatchColl.Count > 0)
                        {
                            nYear = Convert.ToInt32(txtDateBornD);
                        }
                    }
                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ��� ������� ���������� ��� �������� ��������. Message: " + ex.Message + "Source: " + ex.Source, "��������!", MessageBoxButtons.OK);
            }

            if ((nYear < 1900) || (nYear > 2900)) { nYear = 0; }

            return nYear;
        }

        //private int parseBirthDate(string txtDateBornD)
        //{

        //    int nYear = 0;
        //    DateTime dtDate;
        //    try
        //    {
        //        DateTime dtDatrozhd = DateTime.MaxValue;
        //        txtDateBornD = txtDateBornD.TrimEnd();

        //        // ����� ������, ���� ������� � ���������� ����� �� �������������, �� �������� ��� ����� 
        //        // � ���� ���������� 4 ����� � ��� ������ � �������� 1900 - 9999

        //        // ������, ��� ��� �������� ���������� � ����� ������, ������� ���������� � "  .  ."


        //        if (txtDateBornD.Length == 10)
        //        {
        //            if (txtDateBornD.Substring(0, 6) == "  .  .")
        //            {
        //                txtDateBornD = txtDateBornD.Substring(6, 4);
        //            }
        //        }
        //        else
        //        {
        //            // �������� �� ������ ��� �����, ����������� ������������� ��� ���
        //            string[] strData = txtDateBornD.Split('.');
        //            txtDateBornD = "";
        //            foreach (string str in strData)
        //            {
        //                txtDateBornD += str.Trim();
        //            }
        //        }

        //        if (txtDateBornD.Length == 4)
        //        {
        //            MatchCollection myMatchColl = Regex.Matches(txtDateBornD, @"\b[12]\d\d\d");
        //            if (myMatchColl.Count > 0)
        //            {
        //                nYear = Convert.ToInt32(txtDateBornD);
        //            }
        //        }
        //        else
        //        {
        //            if (txtDateBornD.Length == 10)
        //            {
        //                if (DateTime.TryParse(txtDateBornD, out dtDate))
        //                {
        //                    nYear = dtDate.Year;
        //                    dtDatrozhd = dtDate;
        //                }
        //                else
        //                {
        //                    //dtDatrozhd = DateTime.MinValue;
        //                    dtDate = DateTime.MaxValue;
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex) {
        //        MessageBox.Show("������ ��� ������� ���������� ��� �������� ��������. Message: " + ex.Message + "Source: " + ex.Source, "��������!", MessageBoxButtons.OK);
        //    }

        //    if (nYear < 1900) { nYear = 0; }
            
        //    return nYear;
        //}

        public void FolderExist(string m_fullpath)
        {
            if (!Directory.Exists(m_fullpath))
                Directory.CreateDirectory(m_fullpath);
            this.fullpath = m_fullpath;
        }

        private void CreateKtfomsToFind_DBF(bool bVFP, string m_fullpath, string tofind_name)
        {
            try
            {
                DBFcon = new OleDbConnection();
                string DBF_Table_Query = "";
                if (bVFP)
                {
                    DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMSPI numeric(2,0), ZAPROS char(25), NOMIP numeric(9,0), DATZAPR date, FAM char(40), IM char(40), OT char(40), DD_R date, ADDR char(120), ADRES char(150), PRIZ char(1), TYPE_DOG char(150), N_DOG char(10), NAMELONG char(140), FIO_BOSS char(150), TEL_BOSS char(15), ADR_PR char(250))";
                }
                else
                {
                    DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " (NOMOSP numeric(2,0),LITZDOLG numeric(1,0),FIOVK char(100),ZAPROS char(40),GOD numeric(4,0),NOMSPI numeric(2,0),NOMIP numeric(5,0),SUMMA numeric(14,2),VIDVZISK char(100),INNORG char(12), DATZAPR date,ADDR char(120),FLZPRSPI numeric(1,0),DATZAPR1 date,DATZAPR2 date,FL_OKON numeric(1,0),OSNOKON char(250))";
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " (NOMSPI numeric(3,0), ZAPROS char(25), NOMIP numeric(10,0), DATZAPR date, FAM char(40), IM char(40), OT char(40), DD_R date, ADDR char(120), ADRES char(150), PRIZ char(1), TYPE_DOG char(150), N_DOG char(10), NAMELONG char(140), FIO_BOSS char(150), TEL_BOSS char(15), ADR_PR char(250))";
                }
                DBFcon.Open();
                OleDbCommand cmd = new OleDbCommand(DBF_Table_Query, DBFcon);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                DBFcon.Close();
                DBFcon.Dispose();
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
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

        }

        private void CreatePensToFind_DBF(bool bVFP, string m_fullpath, string tofind_name)
        {
            try
            {
                DBFcon = new OleDbConnection();
                string DBF_Table_Query = "";
                if (bVFP)
                {

                    DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMOTD numeric(19,5),NOMSPI numeric(19,5), NOMZAP char(40),ZAPROS char(40),NOMIP numeric(19,5), DATZAPR1 date, DATZAPR2 date, DATZAP date, NAME char(40), FNAME char(40),SNAME char(40), BORN date, ADDR char(120), ADRES char(150), PRIZ char(1), VIDPENS char(150), SUMMA numeric(19,5), KODPFR numeric(19,5), TABNOMIP char(11))";
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMOTD numeric(19,5),NOMSPI numeric(19,5), NOMZAP char(40),ZAPROS char(40),NOMIP numeric(19,5), DATZAP date, NAME char(40), FNAME char(40),SNAME char(40), BORN date, ADDR char(120), ADRES char(150), PRIZ char(1), VIDPENS char(150), SUMMA numeric(19,5), KODPFR numeric(19,5), TABNOMIP char(11))";
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMOTD numeric(1,0),NOMSPI numeric(2,0), NOMZAP char(40),ZAPROS char(40),NOMIP numeric(9,0), DATZAP date, NAMEDOL char(40), FNAMEDOL char(40),SNAMEDOL char(40), BORN date, SUMVZ numeric(12,2), ADDR char(120), ADRES char(150), PRIZ char(1), SUMMA numeric(20,2), KODPFR numeric(20,0), TABNOMIP char(11), KOMMENT char(254))";
                }
                else
                {
                    //DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                    DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBase IV;Data Source={0}", m_fullpath);
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " (NOMOSP numeric(2,0),LITZDOLG numeric(1,0),FIOVK char(100),ZAPROS char(40),GOD numeric(4,0),NOMSPI numeric(2,0),NOMIP numeric(5,0),SUMMA numeric(14,2),VIDVZISK char(100),INNORG char(12), DATZAPR date,ADDR char(120),FLZPRSPI numeric(1,0),DATZAPR1 date,DATZAPR2 date,FL_OKON numeric(1,0),OSNOKON char(250))";
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " (NOMOTD numeric(2,0),NOMSPI numeric(2,0), NOMZAP char(40),ZAPROS char(40),NOMIP numeric(5,0), DATZAPR1 date, DATZAPR2 date, DATZAP date,  NAME char(40), FNAME char(40),SNAME char(40), BORN date, ADDR char(120), ADRES char(150), PRIZ char(1), VIDPENS char(150), SUMMA numeric(20,5), KODPFR numeric(20,5), TABNOMIP char(11))";
                }
                DBFcon.Open();
                OleDbCommand cmd = new OleDbCommand(DBF_Table_Query, DBFcon);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                DBFcon.Close();
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
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }

                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

        }

        private void CreatePotdToFind_DBF(bool bVFP, string m_fullpath, string tofind_name)
        {
            try
            {
                DBFcon = new OleDbConnection();
                string DBF_Table_Query = "";
                bVFP = true;
                if (bVFP)
                {

                    DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMOTD numeric(19,5),NOMSPI numeric(19,5), NOMZAP char(40),ZAPROS char(40),NOMIP numeric(19,5), DATZAPR1 date, DATZAPR2 date, DATZAP date,  NAME char(40), FNAME char(40),SNAME char(40), BORN date, ADDR char(120), ADRES char(150),NAMEORG char(140), ADRORG char(140), DATST date, DATFN date,KODPFR numeric(19,5),TABNOMIP char(11),KOMMENT char(254))";
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMOTD numeric(1,0),NOMSPI numeric(2,0), NOMZAP char(40),ZAPROS char(40),NOMIP numeric(9,0), DATZAP date, FNAMEDOL char(40), NAMEDOL char(40), SNAMEDOL char(40), BORN date, ADDR char(120), ADRES char(150),NAMEORG char(140), ADRORG char(140), DATST date, DATFN date,KODPFR numeric(20,0),TABNOMIP char(11),KOMMENT char(254))";                    

                }
                else
                {
                    //DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                    DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBase IV;Data Source={0}", m_fullpath);
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " (NOMOSP numeric(2,0),LITZDOLG numeric(1,0),FIOVK char(100),ZAPROS char(40),GOD numeric(4,0),NOMSPI numeric(2,0),NOMIP numeric(5,0),SUMMA numeric(14,2),VIDVZISK char(100),INNORG char(12), DATZAPR date,ADDR char(120),FLZPRSPI numeric(1,0),DATZAPR1 date,DATZAPR2 date,FL_OKON numeric(1,0),OSNOKON char(250))";
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " (NOMOTD numeric(2,0),NOMSPI numeric(3,0), NOMZAP char(40),ZAPROS char(40),NOMIP numeric(10,0), DATZAPR1 date, DATZAPR2 date, DATZAP date,  NAME char(40), FNAME char(40),SNAME char(40), BORN date, ADDR char(120), ADRES char(150),NAMEORG char(140), ADRORG char(140), DATST date, DATFN date,KODPFR numeric(3,0),TABNOMIP char(11),KOMMENT char(254))";
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " (NOMOTD numeric(2,0),NOMSPI numeric(3,0), NOMZAP char(40),ZAPROS char(40),NOMIP numeric(10,0), DATZAP date, FNAMEDOL char(40), NAMEDOL char(40), SNAMEDOL char(40), BORN date, ADDR char(120), ADRES char(150),NAMEORG char(140), ADRORG char(140), DATST date, DATFN date,KODPFR numeric(20,0),TABNOMIP char(11),KOMMENT char(254))";                    
                }
                DBFcon.Open();
                OleDbCommand cmd = new OleDbCommand(DBF_Table_Query, DBFcon);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                DBFcon.Close();
                DBFcon.Dispose();
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
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

        }

        private void CreateKrcToFind_DBF(bool bVFP, string m_fullpath, string tofind_name)
        {
            try
            {
                DBFcon = new OleDbConnection();
                string DBF_Table_Query = "";
                if (bVFP)
                {

                    DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMOTD numeric(19,5),NOMSPI numeric(19,5), NOMZAP char(40),ZAPROS char(40),NOMIP numeric(19,5), DATZAPR1 date, DATZAPR2 date, DATZAP date, NAME char(40), FNAME char(40),SNAME char(40), BORN date, ADDR char(120), ADRES char(150), PRIZ char(1), VIDPENS char(150), SUMMA numeric(19,5), KODPFR numeric(19,5), TABNOMIP char(11))";
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMOTD numeric(19,5),NOMSPI numeric(19,5), NOMZAP char(40),ZAPROS char(40),NOMIP numeric(19,5), DATZAP date, NAME char(40), FNAME char(40),SNAME char(40), BORN date, ADDR char(120), ADRES char(150), PRIZ char(1), VIDPENS char(150), SUMMA numeric(19,5), KODPFR numeric(19,5), TABNOMIP char(11))";
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMOTD numeric(1,0),NOMSPI numeric(2,0), ZAPROS char(40),NOMID char(20),DATID date,NAMEDOL char(40),BORN date,SUMVZ numeric(12,2), ADDR char(120),DATZAPR1 date,DATZAPR2 date,SUMPL numeric(12,2),DATPL date,ADRES char(150),KOMMENT char(254))";
                }
                else
                {
                    //DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                    DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBase IV;Data Source={0}", m_fullpath);
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " (NOMOSP numeric(2,0),LITZDOLG numeric(1,0),FIOVK char(100),ZAPROS char(40),GOD numeric(4,0),NOMSPI numeric(2,0),NOMIP numeric(5,0),SUMMA numeric(14,2),VIDVZISK char(100),INNORG char(12), DATZAPR date,ADDR char(120),FLZPRSPI numeric(1,0),DATZAPR1 date,DATZAPR2 date,FL_OKON numeric(1,0),OSNOKON char(250))";
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " (NOMOTD numeric(2,0),NOMSPI numeric(3,0), ZAPROS char(40),NOMID char(20),DATID date,NAMEDOL char(40),BORN date,SUMVZ numeric(14,2), ADDR char(120),DATZAPR1 date,DATZAPR2 date,SUMPL numeric(14,2),DATPL date,ADRES char(150),KOMMENT char(254))";
                }
                DBFcon.Open();
                OleDbCommand cmd = new OleDbCommand(DBF_Table_Query, DBFcon);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                DBFcon.Close();
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
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }

                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

        }


        private void CreateToFind_DBF(bool bVFP, string m_fullpath, string tofind_name)
        {
            try
            {
                DBFcon = new OleDbConnection();
                string DBF_Table_Query = "";
                if (bVFP)
                {
                    DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMOSP numeric(1,0) NOT NULL,LITZDOLG numeric(0,0) NOT NULL,FIOVK char(50) NOT NULL,ZAPROS char(25) NOT NULL,GOD numeric(3,0) NOT NULL,NOMSPI numeric(2,0) NOT NULL,NOMIP numeric(9,0) NOT NULL,SUMMA numeric(12,2) NOT NULL,VIDVZISK char(100) NOT NULL,INNORG char(12) NOT NULL,DATZAPR date NOT NULL,ADDR char(120) NOT NULL,FLZPRSPI numeric(0,0) NOT NULL,DATZAPR1 date NOT NULL,DATZAPR2 date NOT NULL,FL_OKON numeric(0,0) NOT NULL,OSNOKON char(250) NOT NULL, DATROZHD date NOT NULL)";
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMOSP numeric(1,0) NOT NULL,LITZDOLG numeric(0,0) NOT NULL,FIOVK char(254) NOT NULL,ZAPROS char(25) NOT NULL,GOD numeric(3,0) NOT NULL,NOMSPI numeric(13,0) NOT NULL,NOMIP numeric(9,0) NOT NULL,SUMMA numeric(16,2) NOT NULL,VIDVZISK char(254) NOT NULL,INNORG char(12) NOT NULL,DATZAPR date NOT NULL,ADDR char(254) NOT NULL,FLZPRSPI numeric(0,0) NOT NULL,DATZAPR1 date NOT NULL,DATZAPR2 date NOT NULL,FL_OKON numeric(0,0) NOT NULL,OSNOKON char(250) NOT NULL, DATROZHD date NOT NULL)";
                    //DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (NOMOSP numeric(1,0),LITZDOLG numeric(0,0),FIOVK char(50),ZAPROS char(25),GOD numeric(3,0),NOMSPI numeric(2,0),NOMIP numeric(9,0),SUMMA numeric(13,2),VIDVZISK char(100),INNORG char(12),DATZAPR date,ADDR char(120),FLZPRSPI numeric(0,0),DATZAPR1 date,DATZAPR2 date,FL_OKON numeric(0,0),OSNOKON char(250))";
                }
                else
                {
                    DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                    //DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE III;Data Source={0}", m_fullpath);
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " (NOMOSP numeric(2,0),LITZDOLG numeric(1,0),FIOVK char(50),ZAPROS char(25),GOD numeric(4,0),NOMSPI numeric(3,0),NOMIP numeric(10,0),SUMMA numeric(14,2),VIDVZISK char(100),INNORG char(12),DATZAPR date,ADDR char(120),FLZPRSPI numeric(1,0),DATZAPR1 date,DATZAPR2 date,FL_OKON numeric(1,0),OSNOKON char(250), DATROZHD date)";
                }
                DBFcon.Open();
                OleDbCommand cmd = new OleDbCommand(DBF_Table_Query, DBFcon);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                DBFcon.Close();
                DBFcon.Dispose();
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

            if (DBFcon.State == System.Data.ConnectionState.Open)
            {
                DBFcon.Close();
                DBFcon.Dispose();
            }
        }


        private void CreateToFind_SBER_DBF(bool bVFP, string m_fullpath, string tofind_name)
        {
            try
            {
                // TODO: ������-�� ����� �������������
                DBFcon = new OleDbConnection();
                string DBF_Table_Query = "";
                if (bVFP)
                {
                    DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " FREE CODEPAGE = 866 (FIOVK char(100),ZAPROS char(40),GOD numeric(3,0),NOMSPI numeric(2,0),NOMIP numeric(9,0), DATZAPR date,ADDR char(120),FLZPRSPI numeric(0,0),DATZAPR1 date,DATZAPR2 date)";
                }
                else
                {
                    DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                    DBF_Table_Query = "CREATE TABLE " + tofind_name + " (FIOVK char(100),ZAPROS char(40),GOD numeric(4,0),NOMSPI numeric(3,0),NOMIP numeric(10,0), DATZAPR date,ADDR char(120),FLZPRSPI numeric(1,0),DATZAPR1 date,DATZAPR2 date)";
                }
                DBFcon.Open();
                OleDbCommand cmd = new OleDbCommand(DBF_Table_Query, DBFcon);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                DBFcon.Close();
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

            
            if (DBFcon.State == System.Data.ConnectionState.Open)
            {
                DBFcon.Close();
                DBFcon.Dispose();
            }


        }

        private Int64 WriteToDBF_SBER(DataTable dtTable, bool bVFP, string m_fullpath, string tofind_name, DateTime DatZapr1_param, DateTime DatZapr2_param, string release_name)
        {
            Int64 iCnt = 0;
            DataRow[] FizRows = dtTable.Select("LITZDOLG LIKE '/1/*'", "FIOVK");
            int nom_rows = FizRows.Length;
            prbWritingDBF.Value = 0;
            prbWritingDBF.Maximum = nom_rows;
            prbWritingDBF.Step = 1;

            int col_files = (nom_rows / 700);
            int i = 0;

            archive_folder_tofind = archive_sber_path;

            if (!Directory.Exists(archive_folder_tofind))
            {
                MessageBox.Show("�� ���������� ���� � ���������������� ����� ����������� ���������� ��� ������� ������. � ��� ����� ���� ������� �� �����.", "��������", MessageBoxButtons.OK);
                archive_folder_tofind = "";
            }


            try
            {
                if (col_files != 0)
                {
                    for (i = 0; (i <= col_files); i++)
                    {

                        if (!Directory.Exists(string.Format(@"{0}\{1}", m_fullpath, i.ToString())))
                            Directory.CreateDirectory(string.Format(@"{0}\{1}", m_fullpath, i.ToString()));

                        //m_cmd = new OleDbCommand();

                        if (File.Exists(string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), tofind_name)))
                            File.Delete(string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), tofind_name));

                        if (File.Exists(string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), release_name)))
                            File.Delete(string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), release_name));

                        // ���� ���� tofind  � �����, �� ������� ���
                        if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                            File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));

                        if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)))
                            File.Delete(string.Format(@"{0}\{1}", m_fullpath, release_name));



                        CreateToFind_SBER_DBF(bVFP, m_fullpath + "\\" + i.ToString(), tofind_name);

                        DBFcon = new OleDbConnection();
                        if (bVFP) DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + "\\" + i.ToString() + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                        else DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath + "\\" + i.ToString());
                        DBFcon.Open();

                        for (int j = i * 700; (j < (i + 1) * 700) && (j < nom_rows); j++)
                        {
                            if (InsertRowToDBF_SBER(FizRows[j], 1, DatZapr1_param, DatZapr2_param, tofind_name)) iCnt++;
                            prbWritingDBF.PerformStep();
                            prbWritingDBF.Refresh();
                            System.Windows.Forms.Application.DoEvents();
                        }
                        DBFcon.Close();
                        DBFcon.Dispose();
                        // ��� ��� ���� ����� ��������������!!!
                        Process proc = new Process();
                        proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");
                        //proc.StartInfo.FileName = string.Format(@"{0}\{1}", "C:\\Program Files\\SSP\\InstallInfoChange", "fox622.exe ");

                        proc.StartInfo.Arguments = string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), tofind_name) + " " + string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), release_name);
                        proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";

                        proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                        proc.Start();

                        System.Threading.Thread.Sleep(5000);// ���� 5 ������ ����� �������������� ����������� ��������������

                        DateTime tm;
                        tm = DateTime.Now;
                        while (!File.Exists(string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), release_name)) || (File.GetLastWriteTime(string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), release_name)).AddMilliseconds(100) > tm)) // ���� �� �������� ����������������� ����
                        {
                            System.Threading.Thread.Sleep(1000);// ���� ������� ����� �������������� ����������� ��������������
                            tm = DateTime.Now;
                        }


                        if (File.Exists(string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), tofind_name)))
                            File.Delete(string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), tofind_name));

                        if (!archive_folder_tofind.Equals(""))
                        {
                            if (!Directory.Exists(string.Format(@"{0}\{1}", archive_folder_tofind, DateTime.Today.ToShortDateString())))
                                Directory.CreateDirectory(string.Format(@"{0}\{1}", archive_folder_tofind, DateTime.Today.ToShortDateString()));

                            Copy(string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), release_name), string.Format(@"{0}\{1}\{2}", archive_folder_tofind, DateTime.Today.ToShortDateString(), i.ToString()));

                        }
                    }
                    // ������ ��� ���������� ����������� �� ���������� ���
                    while (Directory.Exists(string.Format(@"{0}\{1}", m_fullpath, i.ToString())))
                    {
                        Directory.Delete(string.Format(@"{0}\{1}", m_fullpath, i.ToString()), true);
                        i++;
                    }


                }
                else
                {
                    // ������ ��� ���������� ����������� �� ���������� ���
                    i = 0;
                    while (Directory.Exists(string.Format(@"{0}\{1}", m_fullpath, i.ToString())))
                    {
                        Directory.Delete(string.Format(@"{0}\{1}", m_fullpath, i.ToString()), true);
                        i++;
                    }

                    if (!Directory.Exists(string.Format(@"{0}", m_fullpath)))
                        Directory.CreateDirectory(string.Format(@"{0}", m_fullpath));


                    //m_cmd = new OleDbCommand();


                    if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                        File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));


                    if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)))
                        File.Delete(string.Format(@"{0}\{1}", m_fullpath, release_name));

                    CreateToFind_SBER_DBF(bVFP, m_fullpath, tofind_name);


                    DBFcon = new OleDbConnection();
                    if (bVFP) DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                    else DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                    DBFcon.Open();

                    for (int j = 0; j < nom_rows; j++)
                    {
                        if (InsertRowToDBF_SBER(FizRows[j], 1, DatZapr1_param, DatZapr2_param, tofind_name)) iCnt++;
                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    DBFcon.Close();
                    DBFcon.Dispose();

                    // ��� ��� ���� ����� ��������������!!!
                    Process proc = new Process();
                    proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");
                    //proc.StartInfo.FileName = string.Format(@"{0}\{1}", "C:\\Program Files\\SSP\\InstallInfoChange", "fox622.exe ");
                    proc.StartInfo.Arguments = string.Format(@"{0}\{1}", m_fullpath, tofind_name) + " " + string.Format(@"{0}\{1}", m_fullpath, release_name);
                    proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";
                    proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                    proc.Start();

                    System.Threading.Thread.Sleep(5000);// ���� 5 ������ ����� �������������� ����������� ��������������

                    DateTime tm;
                    tm = DateTime.Now;
                    while (!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)) || (File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name)).AddMilliseconds(100) > tm)) // ���� �� �������� ����������������� ����
                    {
                        System.Threading.Thread.Sleep(1000);// ���� ������� ����� �������������� ����������� ��������������
                        tm = DateTime.Now;
                    }

                    if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                        File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));

                    if (!archive_folder_tofind.Equals(""))
                    {
                        Copy(string.Format(@"{0}\{1}", m_fullpath, release_name), archive_folder_tofind);
                    }
                }
                prbWritingDBF.PerformStep();
                prbWritingDBF.Refresh();
                System.Windows.Forms.Application.DoEvents();
            }
            catch (OleDbException ole_ex)
            {
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                DBFcon.Close();
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Win32Exception e)
            {
                if (e.NativeErrorCode == ERROR_FILE_NOT_FOUND)
                {
                    MessageBox.Show("������ ����������. ��������� ���� ������� � �����: " + e.Message, "��������!", MessageBoxButtons.OK);
                }

                else if (e.NativeErrorCode == ERROR_ACCESS_DENIED)
                {
                    MessageBox.Show("������ ����������. ������ � ����� ��������: " + e.Message, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }


            return iCnt;

        }


        //private Int64 WriteToDBF_SBER(DataTable dtTable, bool bVFP, string m_fullpath, DateTime DatZapr1_param, DateTime DatZapr2_param)
        //{
        //    string tofind_name = "";
        //    string release_extension = "";
        //    string release_name = makenewSberFileName() + ".dbf";

        //    Int64 iCnt = 0;
        //    DataRow[] FizRows = dtTable.Select("LITZDOLG LIKE '/1/*'", "FIOVK");
        //    int nom_rows = FizRows.Length;
        //    prbWritingDBF.Value = 0;
        //    prbWritingDBF.Maximum = nom_rows;
        //    prbWritingDBF.Step = 1;

        //    int col_files = (nom_rows / 700);
        //    int i = 0;

        //    // �������� �� �����
        //    //archive_folder_tofind = archive_sber_path;
        //    //if (!Directory.Exists(archive_folder_tofind))
        //    //{
        //    //    MessageBox.Show("�� ���������� ���� � ���������������� ����� ����������� ���������� ��� ������� ������. � ��� ����� ���� ������� �� �����.", "��������", MessageBoxButtons.OK);
        //    //    archive_folder_tofind = "";
        //    //}


        //    try
        //    {
        //        // ��������� ������� ����������
        //        if (Directory.Exists(string.Format(@"{0}", m_fullpath)))                    
        //            Directory.Delete(string.Format(@"{0}", m_fullpath),true);
                
        //        // ������� ������� �� �������
        //        Directory.CreateDirectory(string.Format(@"{0}", m_fullpath));
                
        //        for (i = 0; (i <= col_files); i++)
        //        {
        //            // TODO: ��� ��� � ����� ����������� ����� ��� ����� - ����������.
        //            // release_name - ���� �������� ����� ��� �����.

        //            tofind_name = i.ToString() + '_' + release_name;
                    
        //            if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
        //                File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));

        //            if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)))
        //                File.Delete(string.Format(@"{0}\{1}", m_fullpath, release_name));


        //            CreateToFind_SBER_DBF(bVFP, m_fullpath, tofind_name);

        //            DBFcon = new OleDbConnection();
        //            if (bVFP) DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
        //            else DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
        //            DBFcon.Open();

        //            for (int j = i * 700; (j < (i + 1) * 700) && (j < nom_rows); j++)
        //            {
        //                if (InsertRowToDBF_SBER(FizRows[j], 1, DatZapr1_param, DatZapr2_param, tofind_name)) iCnt++;
        //                prbWritingDBF.PerformStep();
        //                prbWritingDBF.Refresh();
        //                System.Windows.Forms.Application.DoEvents();
        //            }
        //            DBFcon.Close();
        //            DBFcon.Dispose();
        //            // ��� ��� ���� ����� ��������������!!!
        //            Process proc = new Process();
        //            proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");

        //            proc.StartInfo.Arguments = string.Format(@"{0}\{1}", m_fullpath, tofind_name) + " " + string.Format(@"{0}\{1}", m_fullpath, release_name);
        //            proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";
                    
        //            proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
        //            proc.Start();

        //            System.Threading.Thread.Sleep(5000);// ���� 5 ������ ����� �������������� ����������� ��������������

        //            DateTime tm;
        //            tm = DateTime.Now;
        //            while (!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)) || (File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name)).AddMilliseconds(100) > tm)) // ���� �� �������� ����������������� ����
        //            {
        //                System.Threading.Thread.Sleep(1000);// ���� ������� ����� �������������� ����������� ��������������
        //                tm = DateTime.Now;
        //            }

        //            // ������������� ���� � ���������� ���
        //            File.Move(string.Format(@"{0}\{1}", m_fullpath, release_name), string.Format(@"{0}\{1}", m_fullpath, makenewSberFileName() + makenewSberFileExt(i)));

        //            if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
        //                File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));

        //            if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)))
        //                File.Delete(string.Format(@"{0}\{1}", m_fullpath, release_name));

        //            // �� ����� - ����� ������ ���
        //            //if (!archive_folder_tofind.Equals(""))
        //            //{
        //            //    if (!Directory.Exists(string.Format(@"{0}\{1}", archive_folder_tofind, DateTime.Today.ToShortDateString())))
        //            //        Directory.CreateDirectory(string.Format(@"{0}\{1}", archive_folder_tofind, DateTime.Today.ToShortDateString()));

        //            //    Copy(string.Format(@"{0}\{1}", m_fullpath, release_name), string.Format(@"{0}\{1}", archive_folder_tofind, DateTime.Today.ToShortDateString()));

        //            //}
        //        }

        //        prbWritingDBF.PerformStep();
        //        prbWritingDBF.Refresh();
        //        System.Windows.Forms.Application.DoEvents();
        //    }
        //    catch (OleDbException ole_ex)
        //    {
        //        if (DBFcon.State == System.Data.ConnectionState.Open)
        //        {
        //            DBFcon.Close();
        //            DBFcon.Dispose();
        //        }
        //        DBFcon.Close();
        //        foreach (OleDbError err in ole_ex.Errors)
        //        {
        //            MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
        //        }
        //    }
        //    catch (Win32Exception e)
        //    {
        //        if (e.NativeErrorCode == ERROR_FILE_NOT_FOUND)
        //        {
        //            MessageBox.Show("������ ����������. ��������� ���� ������� � �����: " + e.Message, "��������!", MessageBoxButtons.OK);
        //        }

        //        else if (e.NativeErrorCode == ERROR_ACCESS_DENIED)
        //        {
        //            MessageBox.Show("������ ����������. ������ � ����� ��������: " + e.Message, "��������!", MessageBoxButtons.OK);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        if (DBFcon.State == System.Data.ConnectionState.Open)
        //        {
        //            DBFcon.Close();
        //            DBFcon.Dispose();
        //        }
        //        MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
        //    }
            

        //    return iCnt;

        //}


        private Int64 WriteKrfomsToDBF(bool bVFP, string m_fullpath, string tofind_name)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);

            Decimal LLogID = 0;

            Decimal nErrorPackID = 0;
            Int64 iCnt = 0;
            string tablename = tofind_name.Substring(0, tofind_name.Length - 4);
            try
            {
                //FolderExist(query_cred_org_path);

                archive_folder_tofind = archive_ktfoms_path;


                if (!Directory.Exists(archive_folder_tofind))
                {
                    MessageBox.Show("�� ���������� ���� � ���������������� ����� ����������� ���������� ��� ������� ������. � ��� ����� ���� ������� �� �����.", "��������", MessageBoxButtons.OK);
                    archive_folder_tofind = "";
                }

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                {
                    DialogResult rv = MessageBox.Show("�� ���� " + string.Format(@"{0}\{1}", m_fullpath, tofind_name) + ", ��������� � ���������������� �����, ���������� ����. �������� ����������, �������� ����� �������, ���� ��� ����������.", "��������", MessageBoxButtons.OK);
                    return iCnt; // ��������� ��������� �������
                }
                
                CreateKtfomsToFind_DBF(bVFP, m_fullpath, tofind_name);

                // TODO: �������� ������� ������ � ��� ��� ������
                DBFcon = new OleDbConnection();
                if (bVFP) DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                else DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                DBFcon.Open();

                //Decimal nOsp = GetOSP_Num();

                prbWritingDBF.Value = 0;

                int iDocCnt = 0; 
                if (DT_ktfoms_doc != null) iDocCnt = DT_ktfoms_doc.Rows.Count; ;

                prbWritingDBF.Maximum = 2*iDocCnt;// 2 ���� ����� update ������� ����� ��� ������� �� ���

                prbWritingDBF.Step = 1;

                // ��������� ������ �� �����
                ////���������� � DBF ������������ ��
                //foreach (DataRow row in DT_ktfoms_reg.Rows)
                //{
                //    // TODO: ���� ����� ������ - ��� KTFOMS
                //    if (InsertKtfomsRowToDBF(row, DatZapr1_ktfoms, DatZapr2_ktfoms, false, tablename)) iCnt++;
                //    prbWritingDBF.PerformStep();
                //    prbWritingDBF.Refresh();
                //    System.Windows.Forms.Application.DoEvents();
                //}

                decimal nPackID, nID;
                string txtPackID, txtID;

                int nAgreementID = 0;
                string txtAgreementCode = "";

                if ((DT_ktfoms_doc != null) && (DT_ktfoms_doc.Rows.Count > 0))
                {
                    // �������� ��� ���������� - � ���������� �� ����� ����� ���������� ��� pens_id ����� pens_agr_code.
                    nAgreementID = Convert.ToInt32(DT_ktfoms_doc.Rows[0]["AGREEMENT_ID"]);
                    txtAgreementCode = GetAgreement_Code(nAgreementID);

                    //TODO: ������� local_log LocalLogID
                    // 1 - c����� �����
                    // 1 - ��� ������ ������
                    LLogID = CreateLLog(conGIBDD, 1, 1, txtAgreementCode, 0, "����� ��������.");

                    foreach (DataRow row in DT_ktfoms_doc.Rows)
                    {
                        nPackID = 0;
                        txtPackID = Convert.ToString(row["pack_id"]);
                        if (!Decimal.TryParse(txtPackID, out nPackID))
                        {
                            nPackID = -1;
                        }

                        nID = 0;
                        txtID = Convert.ToString(row["ext_request_id"]);
                        if (!Decimal.TryParse(txtID, out nID))
                        {
                            nID = -1;
                        }

                        if (InsertKtfomsRowToDBF(row, DatZapr1_ktfoms, DatZapr2_ktfoms, true, tablename, ref nErrorPackID))
                        {
                            // ��������� ���������
                            iCnt++;

                            // �������� � ���
                            //WritePackLog(con, nPackID, "��������� ������ # " + iCnt.ToString() + " ext_request_id = " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");
                            WriteLLog(conGIBDD,LLogID, "��������� ������ # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");

                            
                        }
                        else
                        {
                            //WritePackLog(con, nPackID, "������! ������ # " + nID.ToString() + " ���������� �� �������.\n");
                            WriteLLog(conGIBDD, LLogID, "������! ������ ext_request_id = " + nID.ToString() + " ���������� �� �������.\n");
                            row["GOD"] = -1;
                        }
                        
                        //if (InsertRowToDBF(row, nOsp, 0, 1, DatZapr1_param, DatZapr2_param, tofind_name, true)) iCnt++;
                        prbWritingDBF.PerformStep();
                    }
                    
                    // �������� ���������� � local_log
                    UpdateLLogCount(conGIBDD, LLogID, Convert.ToInt32(iCnt));
                }


                prbWritingDBF.PerformStep();
                DBFcon.Close();
                DBFcon.Dispose();

                //DataTable dt = GetDBFTable("SELECT NOMOSP, LITZDOLG, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FL_OKON, OSNOKON FROM TOFIND", "TOFIND1", string.Format(@"{0}\{1}", fullpath, "tofind.dbf"));
                //DBF.Save(dt, fullpath);


                if (!archive_folder_tofind.Equals(""))
                {
                    // ������� ���� � ������ � ������� ����������
                    Copy(string.Format(@"{0}\{1}", m_fullpath, tofind_name), archive_folder_tofind);
                }
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
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

            // ���� ���� ������� � ��������
            if (nErrorPackID > 0)
            {
                // 2 - ����������
                UpdateLLogStatus(conGIBDD, nErrorPackID, 2);
                WriteLLog(conGIBDD, nErrorPackID, DateTime.Now + " �������� ������ �������� ���������.\n");
            }

            // 2 - ����������
            UpdateLLogStatus(conGIBDD, LLogID, 2);
            WriteLLog(conGIBDD, LLogID, DateTime.Now + " ����� �������� � ����: " + m_fullpath + "\\" + tofind_name + "\n����� � ���� ��������� ��������: " + iCnt.ToString() + "\n");

            // ��������������� ������ ��� ������ ������ �������
            //// �������� ������ �������
            //string[] cols = new string[] {"pack_id" };
            //DataTable PackList = SelectDistinct(DT_ktfoms_doc, cols);

            //decimal nRowPackID;
            //string txtRowPackID;


            //if (PackList != null)
            //{
            //    foreach (DataRow row in PackList.Rows)
            //    {
            //        nRowPackID = 0;
            //        txtRowPackID = Convert.ToString(row["pack_id"]);
            //        if (!Decimal.TryParse(txtRowPackID, out nRowPackID))
            //        {
            //            nRowPackID = -1;
            //        }
            //        WritePackLog(con, nRowPackID, DateTime.Now + " ����� �������� � ����: " + m_fullpath + "\\" + tofind_name+ "\n");
            //        WritePackLog(con, nRowPackID, "����� � ���� ��������� ��������: " + iCnt.ToString() + "\n");
            //    }
            //}


            if (DT_ktfoms_doc != null)
            {
                foreach (DataRow row in DT_ktfoms_doc.Rows)
                {
                    //UpdatePackRequest(row);
                    UpdateExtRequestRow(row);

                    prbWritingDBF.PerformStep();
                    prbWritingDBF.Refresh();
                    System.Windows.Forms.Application.DoEvents();
                }
            }

            return iCnt;

        }


        private static DataTable SelectDistinct(DataTable SourceTable, params string[] FieldNames)
        {
            object[] lastValues;
            DataTable newTable;
            DataRow[] orderedRows;

            if (FieldNames == null || FieldNames.Length == 0)
                throw new ArgumentNullException("FieldNames");

            lastValues = new object[FieldNames.Length];
            newTable = new DataTable();

            foreach (string fieldName in FieldNames)
                newTable.Columns.Add(fieldName, SourceTable.Columns[fieldName].DataType);

            orderedRows = SourceTable.Select("", string.Join(", ", FieldNames));

            foreach (DataRow row in orderedRows)
            {
                if (!fieldValuesAreEqual(lastValues, row, FieldNames))
                {
                    newTable.Rows.Add(createRowClone(row, newTable.NewRow(), FieldNames));

                    setLastValues(lastValues, row, FieldNames);
                }
            }

            return newTable;
        }

        private static bool fieldValuesAreEqual(object[] lastValues, DataRow currentRow, string[] fieldNames)
        {
            bool areEqual = true;

            for (int i = 0; i < fieldNames.Length; i++)
            {
                if (lastValues[i] == null || !lastValues[i].Equals(currentRow[fieldNames[i]]))
                {
                    areEqual = false;
                    break;
                }
            }

            return areEqual;
        }

        private static DataRow createRowClone(DataRow sourceRow, DataRow newRow, string[] fieldNames)
        {
            foreach (string field in fieldNames)
                newRow[field] = sourceRow[field];

            return newRow;
        }

        private static void setLastValues(object[] lastValues, DataRow sourceRow, string[] fieldNames)
        {
            for (int i = 0; i < fieldNames.Length; i++)
                lastValues[i] = sourceRow[fieldNames[i]];
        } 

        private Int64 WritePensToDBF(bool bVFP, string m_fullpath, string tofind_name)
        {
            
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);

            Int64 iCnt = 0;
            
            Decimal nErrorPackID = 0;
            
            Decimal LLogID = 0;

            //string tablename = tofind_name.Substring(0, tofind_name.Length - 4);
            //string tablename = tofind_name.Substring(0, tofind_name.Length - 14) + ".dbf";
            string release_name = "tofind.dbf";
            string tablename = "tofind1.dbf";
            string endtofindname = tofind_name;
            //string tablename = tofind_name;
            //tofind_name = tofind_name.Substring(0, tofind_name.Length - 14) + ".dbf";
            tofind_name = "tofind1.dbf";

            try
            {
                //FolderExist(query_cred_org_path);

                archive_folder_tofind = archive_pens_path;

                if (!Directory.Exists(archive_folder_tofind))
                {
                    MessageBox.Show("�� ���������� ���� � ���������������� ����� ����������� ���������� ��� ������� ������. � ��� ����� ���� ������� �� �����.", "��������", MessageBoxButtons.OK);
                    archive_folder_tofind = "";
                }
                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                {
                    
                    DialogResult rv = MessageBox.Show("�� ���� " + string.Format(@"{0}\{1}", m_fullpath, tofind_name) + ", ��������� � ���������������� �����, ���������� ����. �������� ����������, �������� ����� �������, ���� ��� ����������.", "��������", MessageBoxButtons.OK);
                    return iCnt; // ��������� ��������� �������
                    
                }
                
                CreatePensToFind_DBF(bVFP, m_fullpath, tofind_name);
                

                // TODO: �������� ������� ������ � ��� ��� ������
                DBFcon = new OleDbConnection();
                if (bVFP) DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                else DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                DBFcon.Open();

                //Decimal nOsp = GetOSP_Num();

                prbWritingDBF.Value = 0;
                if (DT_pens_doc != null)
                {
                    prbWritingDBF.Maximum = DT_pens_doc.Rows.Count * 2; //  � 2 ���� ������ ����� ����� ��� update �������
                }
                else
                {
                    prbWritingDBF.Maximum = 0;
                }
                prbWritingDBF.Step = 1;

                decimal nPackID, nID;
                string txtPackID, txtID;
                
                int nAgreementID = 0;
                string txtAgreementCode = "";
                                
                if ((DT_pens_doc != null) && (DT_pens_doc.Rows.Count > 0))
                {
                    // �������� ��� ���������� - � ���������� �� ����� ����� ���������� ��� pens_id ����� pens_agr_code.
                    nAgreementID = Convert.ToInt32(DT_pens_doc.Rows[0]["AGREEMENT_ID"]);
                    txtAgreementCode = GetAgreement_Code(nAgreementID);

                    //TODO: ������� local_log LocalLogID
                    // 1 - c����� �����
                    // 1 - ��� ������ ������
                    LLogID = CreateLLog(conGIBDD, 1, 1, txtAgreementCode, 0, "����� ��������.");
                    
                    foreach (DataRow row in DT_pens_doc.Rows)
                    {
                        // ������ �� ����� ���� �������� - ��� ���� ������ LLogID
                        //nPackID = 0;
                        //txtPackID = Convert.ToString(row["pack_id"]);
                        //if (!Decimal.TryParse(txtPackID, out nPackID))
                        //{
                        //    nPackID = -1;
                        //}

                        nID = 0;
                        txtID = Convert.ToString(row["ext_request_id"]);
                        if (!Decimal.TryParse(txtID, out nID))
                        {
                            nID = -1;
                        }


                        if (InsertPensRowToDBF(row, DatZapr1_pens, DatZapr2_pens, true, tablename, ref nErrorPackID))
                        {
                            // ��������� ���������
                            iCnt++;

                            // TODO: �������� � local_log
                            WriteLLog(conGIBDD, LLogID, "��������� ������ # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");
                            
                            //// �������� � ���
                            //WritePackLog(con, nPackID, "��������� ������ # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");

                        }
                        else
                        {

                            // TODO: �������� � local_log
                            WriteLLog(conGIBDD, LLogID, "������! ������ ext_request_id " + nID.ToString() + " ���������� �� �������.\n");
                            //WritePackLog(con, nPackID, "������! ������ # " + nID.ToString() + " ���������� �� �������.\n");
                            row["GOD"] = -1;
                        }

                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }

                    // �������� ���������� � local_log
                    UpdateLLogCount(conGIBDD, LLogID, Convert.ToInt32(iCnt));
                }


                prbWritingDBF.PerformStep();
                prbWritingDBF.Refresh();
                System.Windows.Forms.Application.DoEvents();
                DBFcon.Close();
                DBFcon.Dispose();

                //DataTable dt = GetDBFTable("SELECT NOMOSP, LITZDOLG, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FL_OKON, OSNOKON FROM TOFIND", "TOFIND1", string.Format(@"{0}\{1}", fullpath, "tofind.dbf"));
                //DBF.Save(dt, fullpath);


                //if (!archive_folder_tofind.Equals(""))
                //{
                //    // ������� ���� � ������ � ������� ����������
                //    Copy(string.Format(@"{0}\{1}", m_fullpath, tofind_name), archive_folder_tofind);
                //}

                //release_name = endtofindname;
                // ��� ��� ���� ����� ��������������!!!
                Process proc = new Process();
                proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");
                proc.StartInfo.Arguments = string.Format(@"{0}\{1}", m_fullpath, tofind_name) + " " + string.Format(@"{0}\{1}", m_fullpath, release_name);
                proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                proc.Start();


                System.Threading.Thread.Sleep(5000);// ���� 5 ������ ����� �������������� ����������� ��������������.
                Int32 iCounter = 0;

                while (!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name))) 
                {
                    System.Threading.Thread.Sleep(1000);
                    iCounter++;
                    if (iCounter == 600)
                    {
                        // ���� ������ 10 �����
                        Exception ex = new Exception("������. ������� � ��� �� ���� ��������������� � ������ Fox 2.x � �� ���������� �� ������ ''����������''.");
                        throw ex;
                    }
                }

                //DateTime tm = File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name));
                //System.Threading.Thread.Sleep(1000);

                //while (File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name)) == tm)
                //{
                //    tm = File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name));
                //    System.Threading.Thread.Sleep(1000);
                //}

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                    File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));

                if (!archive_folder_tofind.Equals(""))
                {
                    Copy(string.Format(@"{0}\{1}", m_fullpath, release_name), archive_folder_tofind);
                }

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, endtofindname)))
                    File.Delete(string.Format(@"{0}\{1}", m_fullpath, endtofindname));

                File.Move(string.Format(@"{0}\{1}", m_fullpath, release_name),string.Format(@"{0}\{1}", m_fullpath, endtofindname));
            
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
                if ((DBFcon != null) &&( DBFcon.State == System.Data.ConnectionState.Open))
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

            // TODO: ������ ������ ������� ����� ������������ ��� ������������ LocalLogID

            // ���� ���� ������� � ��������
            if (nErrorPackID > 0)
            {
                // 2 - ����������
                UpdateLLogStatus(conGIBDD, nErrorPackID, 2);
                WriteLLog(conGIBDD, nErrorPackID, DateTime.Now + " �������� ������ �������� ���������.\n");
            }

            // 2 - ����������
            UpdateLLogStatus(conGIBDD, LLogID, 2);
            WriteLLog(conGIBDD, LLogID, DateTime.Now + " ����� �������� � ����: " + m_fullpath + "\\" + tofind_name + "\n����� � ���� ��������� ��������: " + iCnt.ToString() + "\n");
            
            // ��������������� ������ ��� ������ ������ �������

            //// �������� ������ �������
            //string[] cols = new string[] { "pack_id" };
            //DataTable PackList = SelectDistinct(DT_pens_doc, cols);

            //decimal nRowPackID;
            //string txtRowPackID;


            //if (PackList != null)
            //{
            //    foreach (DataRow row in PackList.Rows)
            //    {
            //        nRowPackID = 0;
            //        txtRowPackID = Convert.ToString(row["pack_id"]);
            //        if (!Decimal.TryParse(txtRowPackID, out nRowPackID))
            //        {
            //            nRowPackID = -1;
            //        }
            //        WritePackLog(con, nRowPackID, DateTime.Now + " ����� �������� � ����: " + m_fullpath + "\\" + tofind_name + "\n");
            //        WritePackLog(con, nRowPackID, "����� � ���� ��������� ��������: " + iCnt.ToString() + "\n");
            //    }
            //}

            if (DT_pens_doc != null)
            {
                foreach (DataRow row in DT_pens_doc.Rows)// select ������ ���� ����������
                {
                    // �������� ������ � �����
                    // UpdatePackRequest(row);

                    // ������� update ������� ext_request
                    UpdateExtRequestRow(row);
                    
                    prbWritingDBF.PerformStep();
                    prbWritingDBF.Refresh();
                    System.Windows.Forms.Application.DoEvents();
                }
            }

            return iCnt;
        }

        private Int64 WritePotdToDBF(bool bVFP, string m_fullpath, string tofind_name)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);
            Decimal LLogID = 0;

            Int64 iCnt = 0;
            Decimal nErrorPackID = 0;
            //string tablename = tofind_name.Substring(0, tofind_name.Length - 4);
            //string tablename = tofind_name.Substring(0, tofind_name.Length - 14) + ".dbf";
            string release_name = "tofind.dbf";
            string tablename = "tofind1.dbf";
            string endtofindname = tofind_name;
            //string tablename = tofind_name;
            //tofind_name = tofind_name.Substring(0, tofind_name.Length - 14) + ".dbf";
            tofind_name = "tofind1.dbf";
            try
            {
                //FolderExist(query_cred_org_path);

                archive_folder_tofind = archive_potd_path;


                if (!Directory.Exists(archive_folder_tofind))
                {
                    MessageBox.Show("�� ���������� ���� � ���������������� ����� ����������� ���������� ��� ������� ������. � ��� ����� ���� ������� �� �����.", "��������", MessageBoxButtons.OK);
                    archive_folder_tofind = "";
                }
                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                {
                    DialogResult rv = MessageBox.Show("�� ���� " + string.Format(@"{0}\{1}", m_fullpath, tofind_name) + ", ��������� � ���������������� �����, ���������� ����. �������� ����������, �������� ����� �������, ���� ��� ����������.", "��������", MessageBoxButtons.OK);
                    return iCnt; // ��������� ��������� �������
                        
                }
                
                CreatePotdToFind_DBF(bVFP, m_fullpath, tofind_name);

                // TODO: �������� ������� ������ � ��� ��� ������
                DBFcon = new OleDbConnection();
                if (bVFP) DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                else DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                DBFcon.Open();

                //Decimal nOsp = GetOSP_Num();

                prbWritingDBF.Value = 0;
                prbWritingDBF.Maximum = DT_potd_doc.Rows.Count;
                prbWritingDBF.Step = 1;

                decimal nPackID, nID;
                string txtPackID, txtID;

                int nAgreementID = 0;
                string txtAgreementCode = "";

                if ((DT_potd_doc != null) && (DT_potd_doc.Rows.Count > 0))
                {

                    // �������� ��� ���������� - � ���������� �� ����� ����� ���������� ��� pens_id ����� pens_agr_code.
                    nAgreementID = Convert.ToInt32(DT_potd_doc.Rows[0]["AGREEMENT_ID"]);
                    txtAgreementCode = GetAgreement_Code(nAgreementID);

                    //TODO: ������� local_log LocalLogID
                    // 1 - c����� �����
                    // 1 - ��� ������ ������
                    LLogID = CreateLLog(conGIBDD, 1, 1, txtAgreementCode, 0, "����� ��������.");
                    

                    //���������� � DBF
                    foreach (DataRow row in DT_potd_doc.Rows)
                    {
                        nPackID = 0;
                        txtPackID = Convert.ToString(row["pack_id"]);
                        if (!Decimal.TryParse(txtPackID, out nPackID))
                        {
                            nPackID = -1;
                        }

                        nID = 0;
                        txtID = Convert.ToString(row["ext_request_id"]);
                        if (!Decimal.TryParse(txtID, out nID))
                        {
                            nID = -1;
                        }

                        if (InsertPotdRowToDBF(row, DatZapr1_potd, DatZapr2_potd, true, tablename, ref nErrorPackID))
                        {
                            // ��������� ���������
                            iCnt++;

                            // �������� � ���
                            //WritePackLog(con, nPackID, "��������� ������ # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");
                            WriteLLog(conGIBDD, LLogID, "��������� ������ # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");

                        }
                        else
                        {
                            //WritePackLog(con, nPackID, "������! ������ # " + nID.ToString() + " ���������� �� �������.\n");
                            WriteLLog(conGIBDD, LLogID, "������! ������ ext_request_id " + nID.ToString() + " ���������� �� �������.\n");
                            row["GOD"] = -1;
                        }

                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    // �������� ���������� � local_log
                    UpdateLLogCount(conGIBDD, LLogID, Convert.ToInt32(iCnt));
                }


                prbWritingDBF.PerformStep();
                DBFcon.Close();
                DBFcon.Dispose();

                //DataTable dt = GetDBFTable("SELECT NOMOSP, LITZDOLG, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FL_OKON, OSNOKON FROM TOFIND", "TOFIND1", string.Format(@"{0}\{1}", fullpath, "tofind.dbf"));
                //DBF.Save(dt, fullpath);


                //if (!archive_folder_tofind.Equals(""))
                //{
                //    // ������� ���� � ������ � ������� ����������
                //    Copy(string.Format(@"{0}\{1}", m_fullpath, tofind_name), archive_folder_tofind);
                //}

                // ��� ��� ���� ����� ��������������!!!
                Process proc = new Process();
                proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");
                proc.StartInfo.Arguments = string.Format(@"{0}\{1}", m_fullpath, tofind_name) + " " + string.Format(@"{0}\{1}", m_fullpath, release_name);
                proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                proc.Start();

                System.Threading.Thread.Sleep(5000);
                Int32 iCounter = 0;

                while (!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)))
                {
                    System.Threading.Thread.Sleep(1000);
                    iCounter++;
                    if (iCounter == 600)
                    {
                        // ���� ������ 10 �����
                        Exception ex = new Exception("������. ����������� ������� � ��� �� ���� ��������������� � ������ Fox 2.x � �� ���������� �� ������ ''����������''.");
                        throw ex;
                    }   
                }
                
                //System.Threading.Thread.Sleep(5000);// ���� 5 ������ ����� �������������� ����������� ��������������.

                //DateTime tm;
                //tm = DateTime.Now;
                //while (!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)) || (File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name)).AddMilliseconds(3000) > tm)) // ���� �� �������� ����������������� ����
                //{
                //    System.Threading.Thread.Sleep(1000);// ���� ������� ����� �������������� ����������� ��������������.
                //    tm = DateTime.Now;
                //}

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                    File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));

                if (!archive_folder_tofind.Equals(""))
                {
                    Copy(string.Format(@"{0}\{1}", m_fullpath, release_name), archive_folder_tofind);
                }

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, endtofindname)))
                    File.Delete(string.Format(@"{0}\{1}", m_fullpath, endtofindname));

                File.Move(string.Format(@"{0}\{1}", m_fullpath, release_name), string.Format(@"{0}\{1}", m_fullpath, endtofindname));            
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
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

            // ���� ���� ������� � ��������
            if (nErrorPackID > 0)
            {
                // 2 - ����������
                UpdateLLogStatus(conGIBDD, nErrorPackID, 2);
                WriteLLog(conGIBDD, nErrorPackID, DateTime.Now + " �������� ������ �������� ���������.\n");
            }

            // 2 - ����������
            UpdateLLogStatus(conGIBDD, LLogID, 2);
            WriteLLog(conGIBDD, LLogID, DateTime.Now + " ����� �������� � ����: " + m_fullpath + "\\" + tofind_name + "\n����� � ���� ��������� ��������: " + iCnt.ToString() + "\n");

            // ��������������� ������ ��� ������ ������ �������


            //// �������� ������ �������
            //string[] cols = new string[] { "pack_id" };
            //DataTable PackList = SelectDistinct(DT_potd_doc, cols);

            //decimal nRowPackID;
            //string txtRowPackID;


            //if (PackList != null)
            //{
            //    foreach (DataRow row in PackList.Rows)
            //    {
            //        nRowPackID = 0;
            //        txtRowPackID = Convert.ToString(row["pack_id"]);
            //        if (!Decimal.TryParse(txtRowPackID, out nRowPackID))
            //        {
            //            nRowPackID = -1;
            //        }
            //        WritePackLog(con, nRowPackID, DateTime.Now + " ����� �������� � ����: " + m_fullpath + "\\" + tofind_name + "\n");
            //        WritePackLog(con, nRowPackID, "����� � ���� ��������� ��������: " + iCnt.ToString() + "\n");
            //    }
            //}

            if (DT_potd_doc != null)
            {
                foreach (DataRow row in DT_potd_doc.Rows)// select ������ ���� ����������
                {
                    UpdateExtRequestRow(row);
                    prbWritingDBF.PerformStep();
                    prbWritingDBF.Refresh();
                    System.Windows.Forms.Application.DoEvents();
                }
            }

            return iCnt;

        }

        private Int64 WriteKrcToDBF(bool bVFP, string m_fullpath, string tofind_name)
        {

            Int64 iCnt = 0;
            //string tablename = tofind_name.Substring(0, tofind_name.Length - 4);
            //string tablename = tofind_name.Substring(0, tofind_name.Length - 14) + ".dbf";
            string release_name = "tofind.dbf";
            string tablename = "tofind1.dbf";
            string endtofindname = tofind_name;
            //string tablename = tofind_name;
            //tofind_name = tofind_name.Substring(0, tofind_name.Length - 14) + ".dbf";
            tofind_name = "tofind1.dbf";

            try
            {
                //FolderExist(query_cred_org_path);

                archive_folder_tofind = archive_krc_path;


                if (!Directory.Exists(archive_folder_tofind))
                {
                    MessageBox.Show("�� ���������� ���� � ���������������� ����� ����������� ���������� ��� ������� ������. � ��� ����� ���� ������� �� �����.", "��������", MessageBoxButtons.OK);
                    archive_folder_tofind = "";
                }
                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                {
                    DialogResult rv = MessageBox.Show("�� ���� " + string.Format(@"{0}\{1}", m_fullpath, tofind_name) + ", ��������� � ���������������� �����, ���������� ����. ������� ���?", "��������", MessageBoxButtons.YesNo);
                    if (rv == DialogResult.Yes)
                    {
                        File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));
                        CreateKrcToFind_DBF(bVFP, m_fullpath, tofind_name);
                    }
                }
                else
                {
                    CreateKrcToFind_DBF(bVFP, m_fullpath, tofind_name);
                }

                iCnt = FillDBF_Krc(bVFP, m_fullpath, iCnt, tablename);

                //DataTable dt = GetDBFTable("SELECT NOMOSP, LITZDOLG, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FL_OKON, OSNOKON FROM TOFIND", "TOFIND1", string.Format(@"{0}\{1}", fullpath, "tofind.dbf"));
                //DBF.Save(dt, fullpath);


                //if (!archive_folder_tofind.Equals(""))
                //{
                //    // ������� ���� � ������ � ������� ����������
                //    Copy(string.Format(@"{0}\{1}", m_fullpath, tofind_name), archive_folder_tofind);
                //}

                //release_name = endtofindname;
                // ��� ��� ���� ����� ��������������!!!
                Process proc = new Process();
                proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");
                proc.StartInfo.Arguments = string.Format(@"{0}\{1}", m_fullpath, tofind_name) + " " + string.Format(@"{0}\{1}", m_fullpath, release_name);
                proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                proc.Start();


                System.Threading.Thread.Sleep(5000);// ���� 5 ������ ����� �������������� ����������� ��������������.

                while (!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)))
                {
                    System.Threading.Thread.Sleep(1000);
                }

                //DateTime tm = File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name));
                //System.Threading.Thread.Sleep(1000);

                //while (File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name)) == tm)
                //{
                //    tm = File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name));
                //    System.Threading.Thread.Sleep(1000);
                //}

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                    File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));

                if (!archive_folder_tofind.Equals(""))
                {
                    Copy(string.Format(@"{0}\{1}", m_fullpath, release_name), archive_folder_tofind);
                }

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, endtofindname)))
                    File.Delete(string.Format(@"{0}\{1}", m_fullpath, endtofindname));

                File.Move(string.Format(@"{0}\{1}", m_fullpath, release_name), string.Format(@"{0}\{1}", m_fullpath, endtofindname));

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
                if ((DBFcon != null) && (DBFcon.State == System.Data.ConnectionState.Open))
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }
            return iCnt;
        }

        private Int64 FillDBF_Krc(bool bVFP, string m_fullpath, Int64 iCnt, string tablename)
        {
            // TODO: �������� ������� ������ � ��� ��� ������
            DBFcon = new OleDbConnection();
            if (bVFP) DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
            else DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
            DBFcon.Open();

            prbWritingDBF.Value = 0;
            prbWritingDBF.Maximum = DT_krc_reg.Rows.Count;
            prbWritingDBF.Step = 1;

            //���������� � DBF ������������ ��
            foreach (DataRow row in DT_krc_reg.Rows)
            {
                if (InsertKrcRowToDBF(row, Convert.ToDateTime("01.01.2005"), DateTime.Today, false, tablename)) iCnt++;

                prbWritingDBF.PerformStep();
                prbWritingDBF.Refresh();
                System.Windows.Forms.Application.DoEvents();

            }

            prbWritingDBF.PerformStep();
            prbWritingDBF.Refresh();
            System.Windows.Forms.Application.DoEvents();
            DBFcon.Close();
            DBFcon.Dispose();

            return iCnt;
        }


        private Int64 WriteToDBF(bool bVFP, string m_fullpath, string tofind_name, DateTime DatZapr1_param, DateTime DatZapr2_param, string release_name)
        {

            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);
            Decimal LLogID = 0;

            Int64 iCnt = 0;
            Decimal nErrorPackID = 0;
            try
            {
                //FolderExist(query_cred_org_path);

                archive_folder_tofind = archive_cred_org_path;

                if (!Directory.Exists(archive_folder_tofind))
                {
                    MessageBox.Show("�� ���������� ���� � ���������������� ����� ����������� ���������� ��� ������� ������. � ��� ����� ���� ������� �� �����.", "��������", MessageBoxButtons.OK);
                    archive_folder_tofind = "";
                }
                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                    File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name)); // ������� ���� tofind - �.�. �� ��� ����� �� ��������

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)))
                {
                    DialogResult rv = MessageBox.Show("�� ���� " + string.Format(@"{0}\{1}", m_fullpath, release_name) + ", ��������� � ���������������� �����, ���������� ����. �������� ����������, �������� ����� �������, ���� ��� ����������.", "��������", MessageBoxButtons.OK);
                    return iCnt; // ��������� ��������� �������
                }

                CreateToFind_DBF(bVFP, m_fullpath, tofind_name);
                    

                DBFcon = new OleDbConnection();
                if (bVFP) DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                else
                {
                    DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                }
                DBFcon.Open();

                Decimal nOsp = GetOSP_Num();
                
                // ���� ����� ������-�� 0, �� 1 - ����
                if (nOsp == 0) nOsp = 1;

                prbWritingDBF.Value = 0;

                int iDocCnt = 0;
                if (DT_doc_jur != null) iDocCnt += DT_doc_jur.Rows.Count;
                if (DT_doc_fiz != null) iDocCnt += DT_doc_fiz.Rows.Count;
                //if (DT_okon != null) iOkonCnt = DT_okon.Rows.Count;
                prbWritingDBF.Maximum = iDocCnt * 2; // + iOkonCnt 

                prbWritingDBF.Step = 1;

                decimal nPackID, nID;
                string txtPackID, txtID;

                int nAgreementID = 0;
                string txtAgreementCode = "";
                
                if ((DT_doc_fiz != null) && (DT_doc_fiz.Rows.Count > 0))
                {
                    // �������� ��� ���������� - � ���������� �� ����� ����� ���������� ��� pens_id ����� pens_agr_code.

                    // TODO: ��� ��� ����� ���������� - ������ ��� ����� ������ ���� ����� - �� ����� ���������� � ������!!!
                    nAgreementID = Convert.ToInt32(DT_doc_fiz.Rows[0]["AGREEMENT_ID"]);
                    txtAgreementCode = GetAgreement_Code(nAgreementID);

                    //TODO: ������� local_log LocalLogID
                    // 1 - c����� �����
                    // 1 - ��� ������ ������
                    LLogID = CreateLLog(conGIBDD, 1, 1, txtAgreementCode, 0, "����� ��������.");


                    foreach (DataRow row in DT_doc_fiz.Select("ZAPROS > 0", "FIOVK"))// select ������ ���� ����������
                    {

                        nPackID = 0;
                        txtPackID = Convert.ToString(row["pack_id"]);
                        if (!Decimal.TryParse(txtPackID, out nPackID))
                        {
                            nPackID = -1;
                        }

                        nID = 0;
                        txtID = Convert.ToString(row["ext_request_id"]);
                        if (!Decimal.TryParse(txtID, out nID))
                        {
                            nID = -1;
                        }

                        if (InsertRowToDBF(row, nOsp, 0, 1, DatZapr1_param, DatZapr2_param, tofind_name, true, ref nErrorPackID))
                        {
                            // ��������� ���������
                            iCnt++;

                            // �������� � ��� - ������ ��� ������ � ����� �����������
                            //WritePackLog(con, nPackID, "��������� ������ # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");
                            WriteLLog(conGIBDD, LLogID, "��������� ������ # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");

                        }
                        else
                        {
                            //WritePackLog(con, nPackID, 
                            WriteLLog(conGIBDD, LLogID, "������! ������ ext_request_id = " + nID.ToString() + " ���������� �� �������.\n");
                            row["GOD"] = -1; // ���� ������� ���� ��������� (��� ���� ��������) - �� �� ���� ����� ������ ������ �� ���������
                            // � ��� � ���� ������ - ����� ��� �� �������� ��� �������...
                        }

                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    // �������� ���������� � local_log (����� ����� ����� - ����� ������� ����� ���
                    UpdateLLogCount(conGIBDD, LLogID, Convert.ToInt32(iCnt)); // ��� ����� ��� ���� � ��� � ��.

                }



                if ((DT_doc_jur != null) && (DT_doc_jur.Rows.Count > 0))
                {
                    // �������� ��� ���������� - � ���������� �� ����� ����� ���������� ��� pens_id ����� pens_agr_code.

                    // TODO: ��� ��� ����� ���������� - ������ ��� ����� ������ ���� ����� - �� ����� ���������� � ������!!!
                    // ���� ���� ��� null � ������� ������ � �� ������� ��� ����������
                    if (nAgreementID == 0)
                    {
                        nAgreementID = Convert.ToInt32(DT_doc_jur.Rows[0]["AGREEMENT_ID"]);
                        txtAgreementCode = GetAgreement_Code(nAgreementID);

                        //TODO: ������� local_log LocalLogID
                        // 1 - c����� �����
                        // 1 - ��� ������ ������
                        LLogID = CreateLLog(conGIBDD, 1, 1, txtAgreementCode, 0, "����� ��������.");
                    }

                    foreach (DataRow row in DT_doc_jur.Select("ZAPROS > 0", "FIOVK"))// select ������ ���� ����������
                    {
                        nPackID = 0;
                        txtPackID = Convert.ToString(row["pack_id"]);
                        if (!Decimal.TryParse(txtPackID, out nPackID))
                        {
                            nPackID = -1;
                        }

                        nID = 0;
                        txtID = Convert.ToString(row["ext_request_id"]);
                        if (!Decimal.TryParse(txtID, out nID))
                        {
                            nID = -1;
                        }

                        if (InsertRowToDBF(row, nOsp, 0, 1, DatZapr1_param, DatZapr2_param, tofind_name, true, ref nErrorPackID))
                        {
                            // ��������� ���������
                            iCnt++;

                            // �������� � ��� - ������ ��� ������ � ����� �����������
                            // WritePackLog(con, nPackID, "��������� ������ # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");
                            WriteLLog(conGIBDD, LLogID, "��������� ������ �� ��. ���� # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");

                        }
                        else
                        {
                            // WritePackLog(con, nPackID,
                            WriteLLog(conGIBDD, LLogID, "������! ������ �� ��. ���� ext_request_id = " + nID.ToString() + " ���������� �� �������.\n");
                            row["GOD"] = -1; // ���� ������� ���� ��������� (��� ���� ��������) - �� �� ���� ����� ������ ������ �� ���������
                        }
                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    // �������� ���������� � local_log
                    UpdateLLogCount(conGIBDD, LLogID, Convert.ToInt32(iCnt));
                }

                prbWritingDBF.PerformStep();
                DBFcon.Close();
                DBFcon.Dispose();

                // ��� ��� ���� ����� ��������������!!!
                Process proc = new Process();
                proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");
                proc.StartInfo.Arguments = string.Format(@"{0}\{1}", m_fullpath, tofind_name) + " " + string.Format(@"{0}\{1}", m_fullpath, release_name);
                proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                proc.Start();

                System.Threading.Thread.Sleep(5000);// ���� 5 ������ ����� �������������� ����������� ��������������.

                DateTime tm;
                tm = DateTime.Now;
                Int32 iCounter = 0;
                
                while ((!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name))) || (File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name)).AddMilliseconds(100) > tm)) // ���� �� �������� ����������������� ����
                {
                    System.Threading.Thread.Sleep(1000);// ���� ������� ����� �������������� ����������� ��������������.
                    tm = DateTime.Now;
                    iCounter++;
                    if (iCounter == 600)
                    {
                        // ���� ������ 10 �����
                        Exception ex = new Exception("������. ������� �� ����� ���� ���������� � ��������. ������� ����� ��� ����������� � ������ Fox 2.x");
                        throw ex;
                    }
                }

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                    File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));

                if (!archive_folder_tofind.Equals(""))
                {
                    Copy(string.Format(@"{0}\{1}", m_fullpath, release_name), archive_folder_tofind);
                }

                // ���� ��� ������ - �� ���� ������� UPDATE ������� � �� ���������� ������ - ��������� (10)
                // ��������� ����� 1 ������ - ���� ����������� ������ ������� �� ��� ����� ����� tofind.dbf
                // �� ������ ��� ���� UPDATE

                // �� ��� ���� ������� ���������� ���:
                // ������������� ������ DT_doc_fiz.Rows �� ��������� �����
                // ��������� ������ � �������, ����� ���� ������� ������ � ����� ���� �������� ��������� �����.

                // ���� ���� ������� � �������� - ���������� ������ ���������� ��� �� ����
                if (nErrorPackID > 0)
                {
                    // 2 - ����������
                    UpdateLLogStatus(conGIBDD, nErrorPackID, 2);
                    WriteLLog(conGIBDD, nErrorPackID, DateTime.Now + " �������� ������ �������� ���������.\n");
                }

                // ���������� ������ ���������� ��� ���� �������� �������
                // 2 - ����������
                UpdateLLogStatus(conGIBDD, LLogID, 2);
                WriteLLog(conGIBDD, LLogID, DateTime.Now + " ����� �������� � ����: " + m_fullpath + "\\" + tofind_name + "\n����� � ���� ��������� ��������: " + iCnt.ToString() + "\n");

                // ������ ����� ����� ��� ����� ����������� �����������, � ��� ���� ���� ����������� ��� ���� ��������� ��. �����

                decimal nOrg_id;
                decimal nAgr_id;

                // ��� ����� �������� - �� ����� ���������� 30 (���� �����������). �� - ����� ��� ������.
                txtAgreementCode = "";

                // ����� ������ ���� ��� ������ ������������-����.����� � �������� ������ � �����

                foreach (string txtOrg_id in Legal_List)
                {
                    // �� ���� ����������� ������ ����� ���������
                    nOrg_id = Convert.ToDecimal(txtOrg_id);
                    // �������� Agreement_ID �� ������ �����������
                    nAgr_id = GetAgr_by_Org(nOrg_id);
                    txtAgreementCode = GetAgreement_Code(nAgr_id);

                    if(txtAgreementCode != "30") // ��� 30 - ����������� ���������� �� ����, �.�. �� �������� ���������.
                        CopyLLogParent(conGIBDD, LLogID, txtAgreementCode);
                }

                // ��������������� ������ ��� ������ ������ �������
                //// �������� ������ �������
                //string[] cols = new string[] { "pack_id" };
                //DataTable PackList = SelectDistinct(DT_doc_fiz, cols);

                //decimal nRowPackID;
                //string txtRowPackID;


                //if (PackList != null)
                //{
                //    foreach (DataRow row in PackList.Rows)
                //    {
                //        nRowPackID = 0;
                //        txtRowPackID = Convert.ToString(row["pack_id"]);
                //        if (!Decimal.TryParse(txtRowPackID, out nRowPackID))
                //        {
                //            nRowPackID = -1;
                //        }
                //        WritePackLog(con, nRowPackID, DateTime.Now + " ����� �������� � ����: " + m_fullpath + "\\" + tofind_name + "\n");
                //        WritePackLog(con, nRowPackID, "����� � ���� ��������� ��������: " + iCnt.ToString() + "\n");
                //    }
                //}


                if (DT_doc_fiz != null)
                {
                    foreach (DataRow row in DT_doc_fiz.Rows)// select ������ ���� ����������
                    {
                        //UpdatePackRequest(row);
                        //UpdateKredOrgRequest(row);

                        //UpdateExtRequestRow(row);
                        UpdateExtRequestThrowLegalList(row);

                        //UpdateExtRequestRow(row);
                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }

                // ��������������� - ������ ��� ������ ��� ���� �� ������� ������� � �� ���� ������� ������ �� ���
                //// �������� ������ �������
                //PackList = SelectDistinct(DT_doc_jur, cols);
                
                //if (PackList != null)
                //{
                //    foreach (DataRow row in PackList.Rows)
                //    {
                //        nRowPackID = 0;
                //        txtRowPackID = Convert.ToString(row["pack_id"]);
                //        if (!Decimal.TryParse(txtRowPackID, out nRowPackID))
                //        {
                //            nRowPackID = -1;
                //        }
                //        WritePackLog(con, nRowPackID, DateTime.Now + " ����� �������� � ����: " + m_fullpath + "\\" + tofind_name + "\n");
                //        WritePackLog(con, nRowPackID, "����� � ���� ��������� ��������: " + iCnt.ToString() + "\n");
                //    }
                //}

                if (DT_doc_jur != null)
                {
                    foreach (DataRow row in DT_doc_jur.Rows)// select ������ ���� ����������
                    {
                        //UpdatePackRequest(row);
                        //UpdateKredOrgRequest(row);
                        // ������� ���������� �������� � ��������� �����������
                        // ���� �������� � ��� - ��� �� ������� ������� ������ �� ������ �� ������ - � ��������� ���� �� ����.
                        // ��������� ������ - ��� �������� ������ ����? ����� sql ��� �����-�� ��������
                        // ���� ���� ext_request.req_id = DBF.zapros
                        // ���� ext_request.agreement_code = 

                        //UpdateExtRequestRow(row);
                        UpdateExtRequestThrowLegalList(row);

                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
            }
            catch (OleDbException ole_ex)
            {
                //if (DBFcon.State == System.Data.ConnectionState.Open)
                //{
                //    DBFcon.Close();
                //    DBFcon.Dispose();
                //}
                DBFcon.Close();
                DBFcon.Dispose();
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Win32Exception e)
            {
                if (e.NativeErrorCode == ERROR_FILE_NOT_FOUND)
                {
                    MessageBox.Show("������ ����������. ��������� ���� ������� � �����: " + e.Message, "��������!", MessageBoxButtons.OK);
                }

                else if (e.NativeErrorCode == ERROR_ACCESS_DENIED)
                {
                    MessageBox.Show("������ ����������. ������ � ����� ��������: " + e.Message, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if ((DBFcon != null) && (DBFcon.State == System.Data.ConnectionState.Open))
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }
            

            return iCnt;

        }


        private bool InsertKtfomsRowToDBF(DataRow row, DateTime dat1, DateTime dat2, bool bDoc, string tablename, ref decimal nErrorPackID)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);

            // TODO: ������� �� ���������� ��� ���������� � row
            int nAgreementID = Convert.ToInt32(row["AGREEMENT_ID"]);

            // ��� ��� ���������� ������ ���� ��������� �� ������ ������ ����� �����
            string txtAgreementCode = "";
            string txtAgentCode = "";
            string txtAgentDeptCode = "";


            DateTime DatZapr = DateTime.Today;
            DateTime dtDate;
            Decimal nYear = DateTime.Today.Year;

            string txtNOMSPI, txtZAPROS, txtNOMIP, txtDD_R, txtADDR, txtFIO;
            int nNomspi = 0;
            bool bNoYearBorn = false;

            int iRewriteState = 2;

            try
            {
                txtDD_R = Convert.ToString(row["DATROZHD"]).Trim();
                dtDate = DateTime.MaxValue;

                // �������� ���� �������� � ��� ��������
                if (!DateTime.TryParse(txtDD_R, out dtDate))
                {
                    dtDate = DateTime.MaxValue;
                    bNoYearBorn = true; // �������� � ����� ��������
                }
            
                nAgreementID = Convert.ToInt32(row["AGREEMENT_ID"]);

                if (!bNoYearBorn)
                {

                    if (DBFcon != null)
                    {


                        m_cmd = new OleDbCommand();
                        m_cmd.Connection = DBFcon;

                        m_cmd.CommandText = "INSERT INTO " + tablename + " (NOMSPI, ZAPROS, NOMIP, DATZAPR, FAM, IM, OT, DD_R, ADDR) VALUES (";

                        txtNOMSPI = Convert.ToString(row["NOMSPI"]).Trim();
                        if (!Int32.TryParse(txtNOMSPI, out nNomspi))
                        {
                            nNomspi = 0;
                        }

                        m_cmd.CommandText += Convert.ToString(nNomspi);
                        //m_cmd.CommandText += Convert.ToString(row["NOMSPI"]).Trim();


                        txtZAPROS = cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 25);
                        m_cmd.CommandText += ", '" + txtZAPROS + "'";

                        //// �m��� ������� � 20 �������� �������� ������ '86/'
                        //txtNum_IP = cutEnd(Convert.ToString(row["ZAPROS"]).Trim().Substring(3), 20);
                        //m_cmd.CommandText += ", '" + txtNum_IP + "'";

                        // ����� - ����� 5 ������... =((
                        // �������� ������������� ������ � ������ ���

                        //txtNOMIP = Convert.ToString(row["NOMIP"]).Trim();
                        //if (txtNOMIP.Trim() != "")
                        //{
                        //    String[] strings = txtNOMIP.Split('/');
                        //    m_cmd.CommandText += ", " + Convert.ToString(strings[2]);
                        //}
                        //else
                        //{
                        //    m_cmd.CommandText += ", 0";
                        //}

                        txtNOMIP = Convert.ToString(row["IPNO_NUM"]).Trim();
                        Decimal nNOMIP = 0;
                        if (!Decimal.TryParse(txtNOMIP, out nNOMIP))
                        {
                            nNOMIP = 0;
                        }

                        string txtNPP = Convert.ToInt32(nNOMIP).ToString();

                        m_cmd.CommandText += ", " + txtNPP;

                        if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))
                        {
                            DatZapr = DateTime.Today;
                        }

                        m_cmd.Parameters.Add(new OleDbParameter("DATZAPR", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR", System.Data.DataRowVersion.Original, DatZapr));
                        m_cmd.CommandText += ", ?";

                        //m_cmd.CommandText += ", " + Convert.ToString(row["NOMOTD"]);

                        // ���� �������� �� FIOVK �������� �, � � �

                        txtFIO = Convert.ToString(row["FIOVK"]).Trim();
                        String[] Names;
                        Names = parseFIO(txtFIO);

                        if (Names.Length > 0) m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(Names[0]), 40) + "'";
                        else m_cmd.CommandText += ", ''";

                        if (Names.Length > 1) m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(Names[1]), 40) + "'";
                        else m_cmd.CommandText += ", ''";

                        if (Names.Length > 2) m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(Names[2]), 40) + "'";
                        else m_cmd.CommandText += ", ''";

                        //m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 100) + "'";


                        // ���� �������� �� ��� ������ ��������
                        m_cmd.Parameters.Add(new OleDbParameter("DD_R", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "BORN", System.Data.DataRowVersion.Original, dtDate));
                        m_cmd.CommandText += ", ?";


                        txtADDR = cutEnd(Convert.ToString(row["ADDR"]).Trim(), 120);
                        m_cmd.CommandText += ", '" + txtADDR + "'";


                        m_cmd.CommandText += ')';
                        m_cmd.ExecuteNonQuery();
                        m_cmd.Dispose();

                        return true;
                    }
                    else return false;
                }
                else
                {
                    string txtResponse = "";
                    if (bNoYearBorn)
                    {
                        txtResponse += "������ �� ��� ��������� ������������, ��� ��� � �������� - ���. ���� �� ��������� ���� ��������.";
                    }

                    decimal nStatus = 15;
                    string txtZapros = Convert.ToString(row["ZAPROS"]).Trim();
                    Decimal nID = 0;
                    if (!Decimal.TryParse(txtZapros, out nID))
                    {
                        nID = -1;
                    }

                    txtResponse += " ZAPROS = " + txtZapros + "\n";
                    // TODO: �������� ����� ������� �� nID

                    if (nID > 0)
                    {
                        
                        //InsertZaprosTo_PK_OSP(con, nID, txtResponse, DateTime.Now, nStatus, ktfoms_id, ref iRewriteState);
                        // TODO: ��� ���� �������� ����� ������� - ����� ������������ ������� =)
                        // �������� ��������� - �����, agent_agreement, agent_dept_code, agent_code, enity_name
                        txtAgreementCode = GetAgreement_Code(nAgreementID);
                        txtAgentCode = GetAgent_Code(nAgreementID);
                        txtAgentDeptCode = GetAgentDept_Code(nAgreementID);

                        string txtEntityName = GetLegal_Name(ktfoms_id);

                        // �������� nAgent_id, nAgent_dept_id
                        decimal nAgent_id = GetAgent_ID(nAgreementID);
                        decimal nAgent_dept_id = GetAgentDept_ID(nAgreementID);

                        // ���� ������ ������ ��� �� ������� - �� ������
                        if (nErrorPackID == 0)
                        {
                            // ������ ������ ��������� - 70
                            //nErrorPackID = ID_CreateDX_PACK_I(con, 70, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);
                            // TODO: �������� ����� ��� � ��������
                            // -1 ������ �� ��� ���������, ������������� ����������� ����� ������� � ���������� �������
                            nErrorPackID = CreateLLog(conGIBDD, 1, -1, txtAgreementCode, 0, "����� ��������, ������� �� ���� ��������� �.�. � ��� �� ��������� ������������ ����.\n");
                            // WritePackLog(con, nErrorPackID, "���� ����� ������ ��� �������� �������, ������� ������������� ������� ��� �������� ������������ �������� (��� ���� �������� ��� ���). ����� ������ � ������ ����������� ����, ��� ����� ��������� ������������ ������ ������ ��������� ��� ����� ����� �� ������� �����.");
                        }

                        InsertResponseIntTable(con, nID, txtResponse, DateTime.Now, nStatus, ktfoms_id, ref iRewriteState, nErrorPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName);

                        // ����� � ��� ��� ������ �� ��������
                        WriteLLog(conGIBDD, nErrorPackID, txtResponse);
                        // ������� ++ ��� ���������� �������� � ���� ������
                        AppendLLogCount(conGIBDD, nErrorPackID, 1);
                    }
                    return false;
                }
        }
        catch (OleDbException ole_ex)
        {
            foreach (OleDbError err in ole_ex.Errors)
            {
                MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
            }
            return false;
        }
        catch (Exception ex)
        {
            MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            return false;
        }
        }

        private bool InsertPensRowToDBF(DataRow row, DateTime dat1, DateTime dat2, bool bDoc, string tablename, ref decimal nErrorPackID)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);
            
            // TODO: ������� �� ���������� ��� ���������� � row
            int nAgreementID = Convert.ToInt32(row["AGREEMENT_ID"]);
            
            // ��� ��� ���������� ������ ���� ��������� �� ������ ������ ����� �����
            string txtAgreementCode = "";
            string txtAgentCode = "";
            string txtAgentDeptCode = "";

            DateTime DatZapr = DateTime.Today;
            DateTime dtDate, dtDatrozhd;
            Decimal nYear = DateTime.Today.Year;
            Double sum = 0;
            bool bNoYearBorn = false;
            int nBirthYear = 0;
            string txtDateBornD = "";
            int iRewriteState = 2;
            
            string txtNomspi = "";
            int nNomspi = 0;

            txtDateBornD = Convert.ToString(row["DATROZHD"]).Trim();
            dtDatrozhd = DateTime.MaxValue;

            // �������� ���� �������� � ��� ��������
            if (!DateTime.TryParse(txtDateBornD, out dtDatrozhd))
            {
                bNoYearBorn = true; // �������� � ����� ��������
            }
            try
            {
                if (!bNoYearBorn)// ���� ��� ������� �� ���� ��������
                {

                    if (DBFcon != null)
                    {
                        
                            m_cmd = new OleDbCommand();
                            m_cmd.Connection = DBFcon;

                            m_cmd.CommandText = "INSERT INTO " + tablename + " (NOMOTD, NOMSPI, NOMZAP, ZAPROS, NOMIP, DATZAP, NAMEDOL, FNAMEDOL, SNAMEDOL, BORN, SUMVZ, ADDR) VALUES (";

                            m_cmd.CommandText += Convert.ToString(row["DIV"]);

                            txtNomspi = Convert.ToString(row["NOMSPI"]).Trim();
                            if (!Int32.TryParse(txtNomspi, out nNomspi))
                            {
                                nNomspi = 0;
                            }

                            m_cmd.CommandText += ", " + Convert.ToString(nNomspi);


                            m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40) + "'";

                            // ��������!!! ��� ��������� ������ �� ����������� ���� ����� �� zapros � nomzap
                            m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["NOMIP"]).Trim(), 40) + "'";

                            //// ����� - ����� 5 ������... =((
                            //// �������� ������������� ������ � ������ ���

                            //String txtNOMIP = Convert.ToString(row["NOMIP"]).Trim();
                            //if (txtNOMIP.Trim() != "")
                            //{
                            //    String[] strings = txtNOMIP.Split('/');
                            //    m_cmd.CommandText += ", " + Convert.ToString(strings[2]);
                            //}
                            //else
                            //{
                            //    m_cmd.CommandText += ", 0";
                            //}

                            string txtNOMIP = Convert.ToString(row["IPNO_NUM"]).Trim();
                            Decimal nNOMIP = 0;
                            if (!Decimal.TryParse(txtNOMIP, out nNOMIP))
                            {
                                nNOMIP = 0;
                            }

                            string txtNPP = Convert.ToInt32(nNOMIP).ToString();

                            m_cmd.CommandText += ", " + txtNPP;


                            dtDate = DateTime.Today;
                            m_cmd.Parameters.Add(new OleDbParameter("DATZAP", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAP", System.Data.DataRowVersion.Original, dtDate));
                            m_cmd.CommandText += ", ?";


                            // ����� �������� ����� ��� ParceFIO

                            String txtFIO = Convert.ToString(row["FIOVK"]).Trim();
                            String[] Names;

                            int i = 0;
                            if (txtFIO.Trim() != "")
                            {
                                while (txtFIO.IndexOf("  ") != -1)
                                {
                                    txtFIO = txtFIO.Replace("  ", " ");
                                    i++;
                                    if (i > 100)
                                    {
                                        break;
                                    }
                                }
                                Names = txtFIO.Split(' ');

                                if ((Names[1].Trim() == "") && (Names.Length > 2))
                                {
                                    Names[1] = Names[2];
                                }

                                if (Names.Length > 0) m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(Names[0]), 40) + "'";
                                else m_cmd.CommandText += ", ''";

                                if (Names.Length > 1) m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(Names[1]), 40) + "'";
                                else m_cmd.CommandText += ", ''";

                                if (Names.Length > 2)
                                {

                                    string txtOt = "";
                                    if (Names.Length > 3)
                                    {
                                        txtOt = "";
                                    }
                                    for (int iCnt = 2; iCnt < Names.Length; iCnt++)
                                    {
                                        txtOt += Convert.ToString(Names[iCnt]).Trim() + ' ';
                                    }
                                    m_cmd.CommandText += ", '" + cutEnd(txtOt, 40) + "'";
                                }
                                else m_cmd.CommandText += ", ''";
                            }
                            else
                            {
                                m_cmd.CommandText += ", '', '', ''";
                            }

                            m_cmd.Parameters.Add(new OleDbParameter("BORN", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "BORN", System.Data.DataRowVersion.Original, dtDatrozhd));

                            m_cmd.CommandText += ", ?";

                            if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                            {
                                sum = 0;
                            }

                            m_cmd.CommandText += ", " + sum.ToString("F2").Replace(',', '.');

                            m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["ADDR"]).Trim().ToUpper(), 120) + "'";

                            m_cmd.CommandText += ')';
                            m_cmd.ExecuteNonQuery();
                            m_cmd.Dispose();
                        
                        return true;
                    }
                    else return false;
                }
                else
                {
                    string txtResponse = "";
                    if (bNoYearBorn)
                    {
                        txtResponse += "������ �� ��� ��������� ������������, ��� ��� � �������� - ���. ���� �� ��������� ���� ��������.";
                    }

                    decimal nStatus = 15;
                    string txtZapros = Convert.ToString(row["ZAPROS"]).Trim();
                    Decimal nID = 0;
                    if (!Decimal.TryParse(txtZapros, out nID))
                    {
                        nID = -1;
                    }

                    txtResponse += " ZAPROS = " + txtZapros + "\n";
                    // TODO: �������� ����� ������� �� nID

                    if (nID > 0)
                    {
                        // TODO: ��� ���� �������� ����� ������� - ����� ������������ ������� =)
                        // �������� ��������� - �����, agent_agreement, agent_dept_code, agent_code, enity_name
                        txtAgreementCode = GetAgreement_Code(nAgreementID);
                        txtAgentCode = GetAgent_Code(nAgreementID);
                        txtAgentDeptCode = GetAgentDept_Code(nAgreementID);

                        string txtEntityName = GetLegal_Name(pens_id);

                        // �������� nAgent_id, nAgent_dept_id
                        decimal nAgent_id = GetAgent_ID(nAgreementID);
                        decimal nAgent_dept_id = GetAgentDept_ID(nAgreementID);

                        // ���� ������ ������ ��� �� ������� - �� ������
                        if (nErrorPackID == 0)
                        {
                            // TODO: �������� ����� ��� � ��������
                            nErrorPackID = CreateLLog(conGIBDD, 1, -1, txtAgreementCode, 0, "����� ��������, ������� �� ���� ��������� �.�. � ��� �� ��������� ������������ ����.\n");
                            //nErrorPackID = ID_CreateDX_PACK_I(con, 70, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);
                            //WritePackLog(con, nErrorPackID, "���� ����� ������ ��� �������� �������, ������� ������������� ������� ��� �������� ������������ �������� (��� ���� �������� ��� ���). ����� ������ � ������ ����������� ����, ��� ����� ��������� ������������ ������ ������ ��������� ��� ����� ����� �� ������� �����.");
                        }

                        // ����� � ��� ��� ������ �� ��������
                        WriteLLog(conGIBDD, nErrorPackID, txtResponse);
                        // ������� ++ ��� ���������� �������� � ���� ������
                        AppendLLogCount(conGIBDD, nErrorPackID, 1);

                        // �������� � ������������ ������� �����, ����� ������� nErrorPackID
                        InsertResponseIntTable(con, nID, txtResponse, DateTime.Now, nStatus, pens_id, ref iRewriteState, nErrorPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName);
                        
                        //InsertZaprosTo_PK_OSP(con, nID, txtResponse, DateTime.Now, nStatus, pens_id, ref iRewriteState);
                    }
                    return false;
                }
        }
        catch (Exception ex)
        {
            //if (DBFcon != null) DBFcon.Close();
            MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            return false;
        }
        }

        private bool InsertPotdRowToDBF(DataRow row, DateTime dat1, DateTime dat2, bool bDoc, string tablename, ref decimal nErrorPackID)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);

            DateTime DatZapr = DateTime.Today;
            DateTime dtDate, dtDatrozhd;
            Decimal nYear = DateTime.Today.Year;

            // TODO: ������� �� ���������� ��� ���������� � row
            int nAgreementID = Convert.ToInt32(row["AGREEMENT_ID"]);
            // ��� ��� ���������� ������ ���� ��������� �� ������ ������ ����� �����
            string txtAgreementCode = "";
            string txtAgentCode = "";
            string txtAgentDeptCode = "";

            
            string txtNomspi = "";
            int nNomspi = 0;
            
            bool gooddate = true;
            bool bNoYearBorn = false;

            int iRewriteState = 2;

            int nBirthYear = 0;
            string txtDateBornD = "";

            txtDateBornD = Convert.ToString(row["DATROZHD"]).Trim();
            dtDatrozhd = DateTime.MaxValue;

            // �������� ���� �������� � ��� ��������
            if (!DateTime.TryParse(txtDateBornD, out dtDatrozhd))
            {
                bNoYearBorn = true; // �������� � ����� ��������
            }
            try
            {
                if (!bNoYearBorn)// ���� ��� ������� �� ���� ��������
                {
                    if (DBFcon != null)
                    {
                            
                                m_cmd = new OleDbCommand();
                                m_cmd.Connection = DBFcon;

                                m_cmd.CommandText = "INSERT INTO " + tablename + " (NOMOTD, NOMSPI, NOMZAP, ZAPROS, NOMIP, DATZAP, FNAMEDOL, NAMEDOL, SNAMEDOL, BORN, ADDR) VALUES (";

                                m_cmd.CommandText += Convert.ToString(row["DIV"]);


                                txtNomspi = Convert.ToString(row["NOMSPI"]).Trim();
                                if (!Int32.TryParse(txtNomspi, out nNomspi))
                                {
                                    nNomspi = 0;
                                }

                                m_cmd.CommandText += ", " + Convert.ToString(nNomspi);

                                // m_cmd.CommandText += ", " + Convert.ToString(row["NOMSPI"]).Trim();

                                //
                                m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40) + "'";

                                /*
                                dtDate = DateTime.Today;
                                m_cmd.Parameters.Add(new OleDbParameter("DATZAP", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAP", System.Data.DataRowVersion.Original, dtDate));
                                m_cmd.CommandText += ", ?";
                                */
                                m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40) + "'";

                                // ����� - ����� 5 ������... =((
                                // �������� ������������� ������ � ������ ���

                                //String txtNOMIP = Convert.ToString(row["NOMIP"]).Trim();
                                //if (txtNOMIP.Trim() != "")
                                //{
                                //    String[] strings = txtNOMIP.Split('/');
                                //    m_cmd.CommandText += ", " + Convert.ToString(strings[2]);
                                //}
                                //else
                                //{
                                //    m_cmd.CommandText += ", 0";
                                //}

                                string txtNOMIP = Convert.ToString(row["IPNO_NUM"]).Trim();
                                Decimal nNOMIP = 0;
                                if (!Decimal.TryParse(txtNOMIP, out nNOMIP))
                                {
                                    nNOMIP = 0;
                                }

                                string txtNPP = Convert.ToInt32(nNOMIP).ToString();

                                m_cmd.CommandText += ", " + txtNPP;


                                //m_cmd.CommandText += ", ?";

                                //m_cmd.Parameters.Add(new OleDbParameter("DATZAPR1", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR1", System.Data.DataRowVersion.Original, dat1));
                                //m_cmd.CommandText += ", ?";

                                //m_cmd.Parameters.Add(new OleDbParameter("DATZAPR2", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR2", System.Data.DataRowVersion.Original, dat2));
                                //m_cmd.CommandText += ", ?";

                                dtDate = DateTime.Today;
                                m_cmd.Parameters.Add(new OleDbParameter("DATZAP", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAP", System.Data.DataRowVersion.Original, dtDate));
                                m_cmd.CommandText += ", ?";


                                // ���� �������� �� FIOVK �������� �, � � �

                                String txtFIO = Convert.ToString(row["FIOVK"]).Trim();
                                String[] Names;

                                int i = 0;
                                if (txtFIO.Trim() != "")
                                {
                                    while (txtFIO.IndexOf("  ") != -1)
                                    {
                                        txtFIO = txtFIO.Replace("  ", " ");
                                        i++;
                                        if (i > 100)
                                        {
                                            break;
                                        }
                                    }
                                    Names = txtFIO.Split(' ');


                                    if (Names.Length > 0) m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(Names[0]), 40) + "'";
                                    else m_cmd.CommandText += ", ''";

                                    if (Names.Length > 1) m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(Names[1]), 40) + "'";
                                    else m_cmd.CommandText += ", ''";

                                    if (Names.Length > 2)
                                    {

                                        string txtOt = "";
                                        if (Names.Length > 3)
                                        {
                                            txtOt = "";
                                        }
                                        for (int iCnt = 2; iCnt < Names.Length; iCnt++)
                                        {
                                            txtOt += Convert.ToString(Names[iCnt]).Trim() + ' ';
                                        }
                                        m_cmd.CommandText += ", '" + cutEnd(txtOt, 40) + "'";
                                    }
                                    else m_cmd.CommandText += ", ''";
                                }
                                else
                                {
                                    m_cmd.CommandText += ", '', '', ''";
                                }

                                //m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 100) + "'";


                                if (!(DateTime.TryParse(Convert.ToString(row["DATROZHD"]).TrimEnd(), out dtDate)))
                                {
                                    //dtDate = DateTime.Today;
                                    dtDate = DateTime.MaxValue;

                                    //OleDbTransaction tran;
                                    //con.Open();
                                    //tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                                    //m_cmd = new OleDbCommand();
                                    //m_cmd.Connection = con;
                                    //m_cmd.Transaction = tran;
                                    //m_cmd.CommandText = "UPDATE ZAPROS SET STATUS = '������ � �������', TEXT_ERROR = '� �������� �� �� ������� ���������� ���� �������� �������� � ������� ��.��.����', TEXT = '� �������� �� �� ������� ���������� ���� �������� �������� ��.��.����'";
                                    //m_cmd.CommandText += " WHERE NUM_IP = '" + Convert.ToString(row["ZAPROS"]).TrimEnd() + "'";
                                    //m_cmd.CommandText += " AND FK_LEGAL = " + Convert.ToString(potd_id).TrimEnd();

                                    //int result = m_cmd.ExecuteNonQuery();

                                    //tran.Commit();
                                    //con.Close();

                                }
                                //m_cmd.CommandText += ", " + Convert.ToString(dtDate);
                                //m_cmd.Parameters.Add(new OleDbParameter("BORN", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "BORN", System.Data.DataRowVersion.Original, dtDate));
                                m_cmd.Parameters.Add(new OleDbParameter("BORN", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "BORN", System.Data.DataRowVersion.Original, dtDatrozhd));

                                m_cmd.CommandText += ", ?";


                                m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["ADDR"]).Trim(), 120) + "'";

                                //m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["DATROZHD"]).Trim(), 10) + "'";
                                /*
                                if (bDoc)
                                {
                                    m_cmd.CommandText += ", 1";
                                }
                                else
                                {
                                    m_cmd.CommandText += ", 0";
                                }
                                * /
                                /*
                                m_cmd.Parameters.Add(new OleDbParameter("DATZAPR1", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR1", System.Data.DataRowVersion.Original, dat1));
                                m_cmd.CommandText += ", ?";

                                m_cmd.Parameters.Add(new OleDbParameter("DATZAPR2", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR2", System.Data.DataRowVersion.Original, dat2));
                                m_cmd.CommandText += ", ?";
                                */

                                //m_cmd.CommandText += ", " + sum.ToString("F2").Replace(',', '.');

                                m_cmd.CommandText += ')';
                                m_cmd.ExecuteNonQuery();
                                m_cmd.Dispose();
                           
                        return true;
                    }
                    else return false;
                }
                else
                {
                    string txtResponse = "";
                    if (bNoYearBorn)
                    {
                        txtResponse += "������ �� ��� ��������� ������������, ��� ��� � �������� - ���. ���� �� ��������� ���� ��������.";
                    }

                    decimal nStatus = 15;
                    string txtZapros = Convert.ToString(row["ZAPROS"]).Trim();
                    Decimal nID = 0;
                    if (!Decimal.TryParse(txtZapros, out nID))
                    {
                        nID = -1;
                    }
                    
                    txtResponse += " ZAPROS = " + txtZapros + "\n";
                        // TODO: �������� ����� ������� �� nID

                    if (nID > 0)
                    {
                        // TODO: ��� ���� �������� ����� ������� - ����� ������������ ������� =)
                        // �������� ��������� - �����, agent_agreement, agent_dept_code, agent_code, enity_name
                        txtAgreementCode = GetAgreement_Code(nAgreementID);
                        txtAgentCode = GetAgent_Code(nAgreementID);
                        txtAgentDeptCode = GetAgentDept_Code(nAgreementID);

                        string txtEntityName = GetLegal_Name(potd_id);

                        // �������� nAgent_id, nAgent_dept_id
                        decimal nAgent_id = GetAgent_ID(nAgreementID);
                        decimal nAgent_dept_id = GetAgentDept_ID(nAgreementID);

                        // ���� ������ ������ ��� �� ������� - �� ������
                        if (nErrorPackID == 0)
                        {
                            //nErrorPackID = ID_CreateDX_PACK_I(con, 70, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);
                            //WritePackLog(con, nErrorPackID, "���� ����� ������ ��� �������� �������, ������� ������������� ������� ��� �������� ������������ �������� (��� ���� �������� ��� ���). ����� ������ � ������ ����������� ����, ��� ����� ��������� ������������ ������ ������ ��������� ��� ����� ����� �� ������� �����.");
                            
                            // �������� ����� ��� � ��������
                            nErrorPackID = CreateLLog(conGIBDD, 1, -1, txtAgreementCode, 0, "����� ��������, ������� �� ���� ��������� �.�. � ��� �� ��������� ������������ ����.\n");

                        }

                        InsertResponseIntTable(con, nID, txtResponse, DateTime.Now, nStatus, potd_id, ref iRewriteState, nErrorPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName);

                        // ����� � ��� ��� ������ �� ��������
                        WriteLLog(conGIBDD, nErrorPackID, txtResponse);
                        // ������� ++ ��� ���������� �������� � ���� ������
                        AppendLLogCount(conGIBDD, nErrorPackID, 1);

                        //InsertZaprosTo_PK_OSP(con, nID, txtResponse, DateTime.Now, nStatus, pens_id, ref iRewriteState);

                     //   InsertZaprosTo_PK_OSP(con, nID, txtResponse, DateTime.Now, nStatus, potd_id, ref iRewriteState);

                    }
                    return false;
                }
         }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
                return false;
            }
            catch (Exception ex)
            {
                //if (DBFcon != null) DBFcon.Close();
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }
        }

        private bool InsertKrcRowToDBF(DataRow row, DateTime dat1, DateTime dat2, bool bDoc, string tablename)
        {
            DateTime DatZapr = DateTime.Today;
            DateTime dtDate;
            Decimal nYear = DateTime.Today.Year;
            Double sum = 0;

            if (DBFcon != null)
            {
                try
                {

                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = DBFcon;

                    m_cmd.CommandText = "INSERT INTO " + tablename + " (NOMOTD, NOMSPI, ZAPROS, NOMID, DATID, NAMEDOL, BORN, SUMVZ, ADDR, DATZAPR1, DATZAPR2) VALUES (";

                    m_cmd.CommandText += Convert.ToString(row["NOMOTD"]);
                    m_cmd.CommandText += ", " + Convert.ToString(row["NOMSPI"]).Trim();
                    m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40) + "'";
                    m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["NOMID"]).Trim(), 25) + "'";

                    if (!(DateTime.TryParse(Convert.ToString(row["DATID"]).TrimEnd(), out dtDate)))
                    {
                        dtDate = DateTime.MaxValue;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter("DATID", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATID", System.Data.DataRowVersion.Original, dtDate));
                    m_cmd.CommandText += ", ?";

                    // TODO: 40 - ��� ����� ����!!! � ���� � IP 500 � ZAPROS 100
                    m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["NAMEDOL"]).Trim(), 40) + "'";

                    if (!(DateTime.TryParse(Convert.ToString(row["BORN"]).TrimEnd(), out dtDate)))
                    {
                        dtDate = DateTime.MaxValue;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter("BORN", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "BORN", System.Data.DataRowVersion.Original, dtDate));
                    m_cmd.CommandText += ", ?";

                    if (!(Double.TryParse(Convert.ToString(row["SUMVZ"]), out sum)))
                    {
                        sum = 0;
                    }
                    m_cmd.CommandText += ", " + sum.ToString("F2").Replace(',', '.');
                    
                    m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["ADDR"]).Trim(), 140) + "'";

                    m_cmd.Parameters.Add(new OleDbParameter("DATZAPR1", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR1", System.Data.DataRowVersion.Original, dat1));
                    m_cmd.CommandText += ", ?";

                    m_cmd.Parameters.Add(new OleDbParameter("DATZAPR2", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR2", System.Data.DataRowVersion.Original, dat2));
                    m_cmd.CommandText += ", ?";

                    m_cmd.CommandText += ')';
                    m_cmd.ExecuteNonQuery();
                    m_cmd.Dispose();
                }
                catch (Exception ex)
                {
                    //if (DBFcon != null) DBFcon.Close();
                    MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                    return false;
                }
                return true;
            }
            else return false;
        }




        private bool SetOIpReqBank_ExternalSign(Decimal nID, string txtValue)
        {
            bool bUpdated = true;
            OleDbTransaction tran = null;

            try
            {

                if (nID != 0)
                {

                    if (con != null && con.State != ConnectionState.Closed) con.Close();

                    con.Open();
                    tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = con;
                    m_cmd.Transaction = tran;
                    m_cmd.CommandText = "UPDATE O_IP_REQ_BANK SET EXTERNAL_SIGN = :VALUE WHERE id = :ID ";


                    m_cmd.Parameters.Add(new OleDbParameter(":VALUE", txtValue));
                    m_cmd.Parameters.Add(new OleDbParameter(":ID", nID));

                    if (m_cmd.ExecuteNonQuery() == -1)
                    {
                        bUpdated = false;
                    }

                    tran.Commit();
                    con.Close();

                    if (!bUpdated)
                    {
                        Exception ex = new Exception("Error Updating  O_IP_REQ_BANK table id = " + nID.ToString());
                        throw ex;
                    }
                }
                else
                {
                    // ����� ����� ������ - ������ - ��� ���� ����� ������� ������ �� ������ ������

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                return false;
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }
            return true;

        }

        private bool SetDocument_SUSER_FIO_CHANGE(Decimal nID, decimal nValue)
        {
            bool bUpdated = true;
            OleDbTransaction tran = null;

            try
            {

                if (nID != 0)
                {
                    // UPDATE DOCUMENT SET docstatusid = nStatus WHERE id = nID 

                    if (con != null && con.State != ConnectionState.Closed) con.Close();

                    con.Open();
                    tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = con;
                    m_cmd.Transaction = tran;
                    m_cmd.CommandText = "UPDATE DOCUMENT SET SUSER_ID_CHANGE = :VALUE WHERE id = :ID ";


                    m_cmd.Parameters.Add(new OleDbParameter(":VALUE", nValue));
                    m_cmd.Parameters.Add(new OleDbParameter(":ID", nID));

                    if (m_cmd.ExecuteNonQuery() == -1)
                    {
                        bUpdated = false;
                    }

                    tran.Commit();
                    con.Close();

                    if (!bUpdated)
                    {
                        Exception ex = new Exception("Error Updating document table id = " + nID.ToString());
                        throw ex;
                    }
                }
                else
                {
                    // ����� ����� ������ - ������ - ��� ���� ����� ������� ������ �� ������ ������

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                return false;
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }
            return true;

        }

        private decimal GetPackIdFromSendlistByOrgId(Decimal nO_ID, Decimal org_id)
        {
            decimal nPackID = 0;
            OleDbTransaction tran = null;

            try
            {
                if (nO_ID != 0)
                {
                    if (con != null && con.State != ConnectionState.Closed) con.Close();

                    con.Open();
                    tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = con;
                    m_cmd.Transaction = tran;
                    m_cmd.CommandText = "select first 1 dx_pack_id from sendlist where sendlist_o_id = :O_ID and sendlist_contr = :CONTR_ID";
                    m_cmd.Parameters.Add(new OleDbParameter(":O_ID", nO_ID));
                    m_cmd.Parameters.Add(new OleDbParameter(":CONTR_ID", org_id));

                    nPackID = Convert.ToDecimal(m_cmd.ExecuteScalar());
                    
                    tran.Rollback();
                    con.Close();

                    if (nPackID == -1)
                    {
                        Exception ex = new Exception("Error selecting from sendlist table");
                        throw ex;
                    }
                }
                else
                {
                    // ����� ����� ������ - ������ - ��� ���� ����� ������� ������ �� ������ ������

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }
            
            return nPackID;

        }

        private bool SetSendlistDocumentStatus(Decimal nO_ID, Decimal nPackID, int nStatus)
        {
            // �������� �� ������ ��� ����� DOCUMENT, ������� �������� ������� �������� ������� O_ID � � ������ PACK_ID
            bool bUpdated = true;
            OleDbTransaction tran = null;

            try
            {

                if (nO_ID != 0)
                {
                    // update document d set documentstatusid = :STATUS where d.id in (select id from sendlist where sendlist_o_id = :O_ID and dx_pack_id = :PACK_ID)

                    if (con != null && con.State != ConnectionState.Closed) con.Close();
                    con.Open();
                    tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = con;
                    m_cmd.Transaction = tran;
                    // �������� �� ������ ��� ����� DOCUMENT, ������� �������� ������� �������� ������� O_ID � � ������ PACK_ID
                    m_cmd.CommandText = "UPDATE DOCUMENT SET docstatusid = :STATUS WHERE id = (select first 1 id from sendlist where sendlist_o_id = :O_ID and dx_pack_id = :PACK_ID) ";


                    m_cmd.Parameters.Add(new OleDbParameter(":STATUS", nStatus));
                    m_cmd.Parameters.Add(new OleDbParameter(":O_ID", nO_ID));
                    m_cmd.Parameters.Add(new OleDbParameter(":PACK_ID", nPackID));

                    if (m_cmd.ExecuteNonQuery() == -1)
                    {
                        bUpdated = false;
                    }

                    tran.Commit();
                    con.Close();

                    if (!bUpdated)
                    {
                        Exception ex = new Exception("Error Updating document table sendlist sendlist_o_id = " + nO_ID.ToString());
                        throw ex;
                    }
                }
                else
                {
                    // ����� ����� ������ - ������ - ��� ���� ����� ������� ������ �� ������ ������

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                return false;
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }
            return true;

        }

        private bool SetDocumentStatus(Decimal nID, int nStatus, int nSecondStatus, bool bOpposit)
        {
            // ���� ����� ��������� != SECONDSTATUS - ������������ ��� ������� �������.
            if (bOpposit)
            {
                return SetDocumentStatus(nID, nStatus, nSecondStatus);
            }
            else
            {

                bool bUpdated = true;
                OleDbTransaction tran = null;
                Int32 nResult = 0;

                try
                {

                    if (nID != 0)
                    {
                        // UPDATE DOCUMENT SET docstatusid = nStatus WHERE id = nID and docstatusid != nOppositStatus

                        if (con != null && con.State != ConnectionState.Closed) con.Close();
                        con.Open();
                        tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                        m_cmd = new OleDbCommand();
                        m_cmd.Connection = con;
                        m_cmd.Transaction = tran;
                        m_cmd.CommandText = "UPDATE DOCUMENT SET docstatusid = :STATUS WHERE id = :ID and docstatusid = :SECONDSTATUS";

                        m_cmd.Parameters.Add(new OleDbParameter(":STATUS", nStatus));
                        m_cmd.Parameters.Add(new OleDbParameter(":ID", nID));
                        m_cmd.Parameters.Add(new OleDbParameter(":SECONDSTATUS", nSecondStatus));
                        nResult = m_cmd.ExecuteNonQuery();
                        
                        if (nResult == -1)
                        {
                            bUpdated = false;
                        }

                        tran.Commit();
                        con.Close();

                        if (!bUpdated)
                        {
                            Exception ex = new Exception("Error Updating document table id = " + nID.ToString());
                            throw ex;
                        }
                    }
                    else
                    {
                        // ����� ����� ������ - ������ - ��� ���� ����� ������� ������ �� ������ ������

                    }

                }
                catch (OleDbException ole_ex)
                {
                    foreach (OleDbError err in ole_ex.Errors)
                    {
                        MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                    }
                    if (tran != null)
                    {
                        tran.Rollback();
                    }
                    if (con != null)
                    {
                        con.Close();
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    if (tran != null)
                    {
                        tran.Rollback();
                    }
                    if (con != null)
                    {
                        con.Close();
                    }
                    MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                    return false;
                }
                if (nResult > 0)
                    return true;
                else return false;

            }

        }


        private bool SetDocumentStatus(Decimal nID, int nStatus, int nOppositStatus)
        {
            bool bUpdated = true;
            OleDbTransaction tran = null;
            Int32 nResult = 0;

            try
            {

                    if (nID != 0)
                    {
                        // UPDATE DOCUMENT SET docstatusid = nStatus WHERE id = nID and docstatusid != nOppositStatus

                        if (con != null && con.State != ConnectionState.Closed) con.Close();
                        con.Open();
                        tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                        m_cmd = new OleDbCommand();
                        m_cmd.Connection = con;
                        m_cmd.Transaction = tran;
                        m_cmd.CommandText = "UPDATE DOCUMENT SET docstatusid = :STATUS WHERE id = :ID and docstatusid != :OPPOSITSTATUS";

                        m_cmd.Parameters.Add(new OleDbParameter(":STATUS", nStatus));
                        m_cmd.Parameters.Add(new OleDbParameter(":ID", nID));
                        m_cmd.Parameters.Add(new OleDbParameter(":OPPOSITSTATUS", nOppositStatus));

                        nResult = m_cmd.ExecuteNonQuery();

                        if (nResult == -1)
                        {
                            bUpdated = false;
                        }

                        tran.Commit();
                        con.Close();

                        if (!bUpdated)
                        {
                            Exception ex = new Exception("Error Updating document table id = " + nID.ToString());
                            throw ex;
                        }
                }
                else
                {
                    // ����� ����� ������ - ������ - ��� ���� ����� ������� ������ �� ������ ������

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                return false;
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }

            if (nResult > 0)
                return true;
            else return false;
        }

        private bool SetDocumentStatus(Decimal nID, int nStatus)
        {
            bool bUpdated = true;
            OleDbTransaction tran = null;

            try
            {

                    if (nID != 0)
                    {
                        // UPDATE DOCUMENT SET docstatusid = nStatus WHERE id = nID 


                        if (con != null && con.State != ConnectionState.Closed) con.Close();
                        con.Open();
                        tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                        m_cmd = new OleDbCommand();
                        m_cmd.Connection = con;
                        m_cmd.Transaction = tran;
                        m_cmd.CommandText = "UPDATE DOCUMENT SET docstatusid = :STATUS WHERE id = :ID ";


                        m_cmd.Parameters.Add(new OleDbParameter(":STATUS", nStatus));
                        m_cmd.Parameters.Add(new OleDbParameter(":ID", nID));

                        if (m_cmd.ExecuteNonQuery() == -1)
                        {
                            bUpdated = false;
                        }

                        tran.Commit();
                        con.Close();

                        if (!bUpdated)
                        {
                            Exception ex = new Exception("Error Updating document table id = " + nID.ToString());
                            throw ex;
                        }
                }
                else
                {
                    // ����� ����� ������ - ������ - ��� ���� ����� ������� ������ �� ������ ������

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                return false;
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }
            return true;

        }

        private bool SetExtReqProcessed(Decimal nID, string txtAgreementCode, int nProcessed)
        {
            bool bUpdated = true;
            OleDbTransaction tran = null;

            try
            {

                if (nID != 0)
                {
                    // UPDATE ext_request SET processed = nProcessed WHERE req_id = nID and mvv_agreement_code = txtAgreementCode

                    if (con != null && con.State != ConnectionState.Closed) con.Close();
                    con.Open();
                    tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = con;
                    m_cmd.Transaction = tran;

                    m_cmd.CommandText = "UPDATE ext_request SET processed = :STATUS WHERE req_id = :ID and mvv_agreement_code = :AGREEMENT_CODE";


                    m_cmd.Parameters.Add(new OleDbParameter(":STATUS", nProcessed));
                    m_cmd.Parameters.Add(new OleDbParameter(":ID", nID));
                    m_cmd.Parameters.Add(new OleDbParameter(":AGREEMENT_CODE", txtAgreementCode));
                    

                    if (m_cmd.ExecuteNonQuery() == -1)
                    {
                        bUpdated = false;
                    }

                    tran.Commit();
                    con.Close();

                    if (!bUpdated)
                    {
                        Exception ex = new Exception("Error Updating ext_request table id = " + nID.ToString() + "mvv_agreement_code = " + txtAgreementCode);
                        throw ex;
                    }
                }
                else
                {
                    // ����� ����� ������ - ������ - ��� ���� ����� ������� ������ �� ������ ������

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                return false;
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }
            return true;

        }


        private bool SetExtReqProcessed(Decimal nID, int nProcessed)
        {
            bool bUpdated = true;
            OleDbTransaction tran = null;

            try
            {

                if (nID != 0)
                {
                    // UPDATE ext_request SET processed = nProcessed WHERE ext_request_id = nID 

                    if (con != null && con.State != ConnectionState.Closed) con.Close();
                    con.Open();
                    tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = con;
                    m_cmd.Transaction = tran;

                    m_cmd.CommandText = "UPDATE ext_request SET processed = :STATUS WHERE ext_request_id = :ID ";


                    m_cmd.Parameters.Add(new OleDbParameter(":STATUS", nProcessed));
                    m_cmd.Parameters.Add(new OleDbParameter(":ID", nID));

                    if (m_cmd.ExecuteNonQuery() == -1)
                    {
                        bUpdated = false;
                    }

                    tran.Commit();
                    con.Close();

                    if (!bUpdated)
                    {
                        Exception ex = new Exception("Error Updating ext_request table id = " + nID.ToString());
                        throw ex;
                    }
                }
                else
                {
                    // ����� ����� ������ - ������ - ��� ���� ����� ������� ������ �� ������ ������

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                return false;
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }
            return true;

        }


        private bool UpdateBankFizRequest(DataRow row)
        {
            Decimal nID = 0;
            string txtValue = "InfoChangeCredOrg"; // ������ ���������
            int nGod = 0;
            string txtGod, txtID;
            bool bUpdated = true;
            try
            {
                txtGod = Convert.ToString(row["GOD"]).Trim();
                if (!Int32.TryParse(txtGod, out nGod))
                {
                    nGod = 0;
                }

                txtID = Convert.ToString(row["ZAPROS"]).Trim();

                if (!Decimal.TryParse(txtID, out nID))
                {
                    nID = 0;
                }

                if (nGod != -1)
                {
                    if (nID != 0)
                    {

                        if (!SetOIpReqBank_ExternalSign(nID, txtValue))
                        {
                            bUpdated = false;
                        }


                        if (!bUpdated)
                        {
                            Exception ex = new Exception("Error Updating document table O_IP_REQ_BANK id = " + nID.ToString());
                            throw ex;
                        }
                    }
                }
                else
                {
                    // ����� ����� ������ - ������ - ��� ���� ����� ������� ������ �� ������ ������

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }
            return true;

        }

        private bool UpdateKredOrgRequest(DataRow row)
        {
            decimal nPackID = 0;
            decimal nID = 0;
            int iGod = 0;
            string txtGod, txtID, txtPackID;

            txtGod = Convert.ToString(row["GOD"]).Trim();
            txtID = Convert.ToString(row["ZAPROS"]).Trim();

            // ���� ����� �� ������ ���� �����, row["pack_id"], �� � �����������...
            // ������ ������, ��� ��� ����� �� ������ ���� sendlist � ��� ���� �����,
            // �� � ��� ��������� sendlist �� ����� ������� ID, ������� ���������� ������������ �� LegalList
            // � ��� ������, � ������� �������� ��� ����� sendlist
            // �� ���� �������� ������� �������� req.id � Legal_list
            // 

            if (!Int32.TryParse(txtGod, out iGod))
            {
                iGod = 0;
            }

            if (!Decimal.TryParse(txtID, out nID))
            {
                nID = 0;
            }
            //return UpdatePackRequest(iGod, nID, nPackID, 20);
            return UpdateRequestsThruLegalList(iGod, nID, 20);

        }

        private bool UpdateExtRequestThrowLegalList(DataRow row)
        {
            // �������� ������ �� ������
            decimal nID = 0;
            int iGod = 0;
            string txtGod, txtID, txtPackID, txtReq_id, txtAgreementCode;
            decimal nReq_id = 0;
            decimal nAgreement_code = 0;
            try{

                txtGod = Convert.ToString(row["GOD"]).Trim();

                // �������� req_id
                txtReq_id = Convert.ToString(row["zapros"]).Trim();
                if (!Decimal.TryParse(txtReq_id, out nReq_id))
                {
                    nReq_id = 0;
                }

                if (!Int32.TryParse(txtGod, out iGod))
                {
                    iGod = 0;
                }

                // �������� �� ������ ��� LegalList - �������� mvv_agreement_code
                foreach (string txtOrg_id in Legal_List)
                {
                    decimal nOrg_id = Convert.ToDecimal(txtOrg_id);
                    nAgreement_code =  GetAgr_by_Org(nOrg_id);
                    string txtAgrCode = GetAgreement_Code(nAgreement_code);
                    
                    // �������� ������ � ext_request �� 2-� ����������
                    
                    // ������ ��� ��������� - ����� ��� �� ������������� ��� �������������, ���� �� �������� �� ��� �����
                    SetExtReqProcessed(nReq_id, txtAgrCode, 1);
                    
                    //if (iGod == -1)
                    //{
                    //    SetExtReqProcessed(nReq_id, txtAgrCode, 0);
                    //}
                    //else
                    //{
                    //    SetExtReqProcessed(nReq_id, txtAgrCode, 1);
                    //}
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }
            return true;
                        
        }


        private bool UpdateExtRequestRow(DataRow row)
        {
            // �������� ������ �� ������
            decimal nID = 0;
            decimal nReqID = 0;
            int iGod = 0;
            string txtGod, txtID, txtPackID, txtReqID;

            txtGod = Convert.ToString(row["GOD"]).Trim();
            txtID = Convert.ToString(row["ext_request_id"]).Trim();
            txtReqID = Convert.ToString(row["zapros"]).Trim();

            // ���� ����� �� ������ ���� �����, row["pack_id"], �� � �����������...
            // ������ ������, ��� ��� ����� �� ������ ���� sendlist � ��� ���� �����,
            // �� � ��� ��������� sendlist �� ����� ������� ID, ������� ���������� ������������ �� LegalList
            // � ��� ������, � ������� �������� ��� ����� sendlist
            // �� ���� �������� ������� �������� req.id � Legal_list
            // 

            if (!Int32.TryParse(txtGod, out iGod))
            {
                iGod = 0;
            }

            if (!Decimal.TryParse(txtID, out nID))
            {
                nID = 0;
            }

            if (!Decimal.TryParse(txtReqID, out nReqID))
            {
                nReqID = 0;
            }

            if (iGod == -1)
            {
                // 15 - ��������� � �������
                SetDocumentStatus(nReqID, 15);
            }
            //else
            //{
            //    return SetExtReqProcessed(nID, 1);
            //}

            
            // ������ ��� ��������� ��� �����������, ���� �� ��������� ����� � ���� �������� ��� - �� �� ����� ��������
            return SetExtReqProcessed(nID, 1);
        }

        private bool UpdatePackRequest(DataRow row)
        {
            // �������� ������ �� ������
            decimal nPackID = 0;
            decimal nID = 0;
            int iGod = 0;
            string txtGod, txtID, txtPackID;

            txtGod = Convert.ToString(row["GOD"]).Trim();
            txtID = Convert.ToString(row["ZAPROS"]).Trim();
            txtPackID = Convert.ToString(row["pack_id"]).Trim();

            // ���� ����� �� ������ ���� �����, row["pack_id"], �� � �����������...
            // ������ ������, ��� ��� ����� �� ������ ���� sendlist � ��� ���� �����,
            // �� � ��� ��������� sendlist �� ����� ������� ID, ������� ���������� ������������ �� LegalList
            // � ��� ������, � ������� �������� ��� ����� sendlist
            // �� ���� �������� ������� �������� req.id � Legal_list
            // 

            if (!Int32.TryParse(txtGod, out iGod))
            {
                iGod = 0;
            }

            if (!Decimal.TryParse(txtID, out nID))
            {
                nID = 0;
            }

            if (!Decimal.TryParse(txtPackID, out nPackID))
            {
                nPackID = 0;
            }
            
            return UpdatePackRequest(iGod, nID, nPackID, 20);

        }

        private bool UpdateRequestsThruLegalList(int nGod, decimal nID, decimal nStatus)
        {

            bool bUpdated;
            bool bUpdatedPack;
            try
            {
                bUpdatedPack = true;
                bUpdated = true;
                if (nGod == -1)
                {
                    //  ���� �� ��� �������� ������, �� ����� � ������ �������� ��������� �� ������ 71 - ������ ����������� ��������
                    // ���� � ������ �� ��� ���� � ������� - �� ������ ������� 71
                    // ���� ��� ������ 70 - ��������� ��������.
                    // ��� ������ ��� ���� ������� 70 - � �� 71 - ���� ���� ���� 1 ����������
                    // �� ���� ���������� 71  � ���� ����� ���� 71 � ����� ������ ������� ������ � ������ - �� 70
                    nStatus = 71;
                }

                    foreach (string txtOrg_id in Legal_List)
                    {
                        bUpdatedPack = true;
                        bUpdated = true;
                        decimal nOrg_id = Convert.ToDecimal(txtOrg_id);
                        
                        // ������ ��� ��� sendlist, ��� nID � ������������ txtOrg_id
                        decimal nPackID = GetPackIdFromSendlistByOrgId(nID, nOrg_id);
                        
                        // ���� ������ � ������
                        if (nPackID > 0)
                            {
                                // ������ ������ ������ - ���������, ������ ���� ���� ����� � ������ ���� 1 ������ � ������� - ���� ������ �������� ������ - �� � OppositStatus
                                if (!SetDocumentStatus(nPackID, Convert.ToInt32(nStatus), 23, false))
                                {
                                    // ���� �������� �� ������� - �� �������� ��� ������ ������ ����� � ������ ���:
                                    // ���� ����� ����� �� ������, �� ��������, � ����� - �������� ��� ���
                                    if (!SetDocumentStatus(nPackID, 70, Convert.ToInt32(nStatus)))
                                    {
                                        bUpdatedPack = false;
                                    }
                                }

                                if (nID != 0)
                                {
                                    if (!SetSendlistDocumentStatus(nID, nPackID, Convert.ToInt32(nStatus)))
                                    {
                                        bUpdated = false;
                                    }
                                }
                            }
                    }
                    //if (!(bUpdated && bUpdatedPack))
                    if (!(bUpdated))
                    {
                        Exception ex = new Exception("Error Updating document table id = " + nID.ToString());
                        throw ex;
                    }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }
            return true;
        }


        private bool UpdatePackRequest(int nGod, decimal nID, decimal nPackID, decimal nStatus)
        {
            // �������� ������ �� ������, ��������� �������� ���������� ����� ���������� (� �� DATAROW)
            

            bool bUpdated = true;
            bool bUpdatedPack = true;
            try
            {
                if (nGod == -1)
                {
                    //  ���� �� ��� �������� ������, �� ����� � ������ �������� ��������� �� ������ 71 - ������ ����������� ��������
                    // ��� 70 - �������� ���������
                    nStatus = 71;
                }

                    // ���������� ������ ������
                    if(nPackID != 0){
                            // ���� ������� ����� �� ���� - ���� �� ��� ��������� ���-��
                            // ������ �� ����� ������� ����� ����� � �������� ������� ��� ��������� ���� ������ ����� �����, � �� �������� ���
                            // 23 - �������� ��������� ����������
                            if (!SetDocumentStatus(nPackID, Convert.ToInt32(nStatus), 23, false))
                            {
                                // ���� �������� �� ������� - �� �������� ��� ������ ������ ����� � ������ ���:
                                // ���� ����� ����� �� ������, �� ��������, � ����� - �������� ��� ���
                                if (!SetDocumentStatus(nPackID, 70, Convert.ToInt32(nStatus)))
                                {
                                    bUpdatedPack = false;
                                }
                            }

                            
                        // ���������� ������ �������-������ ������ �������� � ������
                        if (nID != 0)
                        {
                            if (!SetSendlistDocumentStatus(nID, nPackID, Convert.ToInt32(nStatus)))
                            {
                                bUpdated = false;
                            }
                        }
                        
                    }
                    //if (!(bUpdated && bUpdatedPack))
                    if (!(bUpdated))
                    {
                        Exception ex = new Exception("Error Updating document table id = " + nID.ToString());
                        throw ex;
                    }

            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }
            return true;
        }

        
        private bool UpdateRequest(DataRow row)
        {
            
            Decimal nID = 0;
            int nStatus = 10; // ������ ���������
            int nGod = 0;
            string txtGod, txtID;
            bool bUpdated = true;
                try
                {
                    txtGod = Convert.ToString(row["GOD"]).Trim();
                    if (!Int32.TryParse(txtGod, out nGod))
                    {
                        nGod = 0;
                    }

                    txtID = Convert.ToString(row["ZAPROS"]).Trim();
                    
                    if (!Decimal.TryParse(txtID, out nID))
                    {
                        nID = 0;
                    }

                    if (nGod != -1)
                    {
                        if (nID != 0)
                        {
                            
                            if (!SetDocumentStatus(nID,nStatus))
                            {
                                bUpdated = false;
                            } 

                            
                            if (!bUpdated)
                            {
                                Exception ex = new Exception("Error Updating document table id = " + nID.ToString());
                                throw ex;
                            }
                        }
                    }
                    else
                    {
                        // ����� ����� ������ - ������ - ��� ���� ����� ������� ������ �� ������ ������

                    }
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                    return false;
                }
                return true;
        }

        // ������� ������� ������� � �������� ����������� � DBF ����
        private bool InsertRowToDBF(DataRow row, decimal nOsp, decimal bOKON, decimal bZPRSPI, DateTime DatZapr1_param, DateTime DatZapr2_param, string tofind_name, bool bRegTable, ref decimal nErrorPackID)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);

            DateTime DatZapr = DateTime.Today;
            Decimal nLitzDolg = 2;
            DateTime dtDate, dtDatrozhd;
            Decimal nYear = 0;
            Double sum = 0;
            string txtDateBornD = "";
            string txtLitzDolg, txtNameDolg, txtZapros, txtNomspi, txtNOMIP, txtNPP, txtVIDVZISK, txtInnOrg, txtAddr, txtOsnokon;

            int iRewriteState = 2;
            
            int nAgreementID = Convert.ToInt32(row["AGREEMENT_ID"]); // ��� ����� ������ 30 (�����������) - ���� ����� �� ������ ��������
            // ��� ��� ���������� ������ ���� ��������� �� ������ ������ ����� �����
            string txtAgreementCode = "";
            string txtAgentCode = "";
            string txtAgentDeptCode = "";


            txtNomspi = "";
            int nNomspi = 0;
            bool bNoInnJur = false;
            bool bNoYearBorn = false;
            int nBirthYear = 0;

            if (DBFcon != null)
            {
                try
                {
                    dtDatrozhd = DateTime.MaxValue;
                    nLitzDolg = 1;// ��. ����
                    txtLitzDolg = Convert.ToString(row["LITZDOLG"]).Trim();
                    if (!Decimal.TryParse(txtLitzDolg, out nLitzDolg))
                    {
                        nLitzDolg = 0;
                    }

                    txtInnOrg = Convert.ToString(row["INNORG"]).Trim();
                    if (nLitzDolg == 1)
                    {
                        if (txtInnOrg != "" && txtInnOrg.Length == 10)
                        {
                            txtInnOrg = "00" + txtInnOrg; // � ��. ��� ����� ��������� 00, �� 12 ��������
                        }
                    }
                    txtInnOrg = cutEnd(txtInnOrg.Trim(), 12);

                    // ��������� vid_d � inn_d, ���� ��� ����� �� �� ���������
                    if (bRegTable && (nLitzDolg == 1))
                    {
                        
                        if (txtInnOrg.Length < 10)
                        {
                            bNoInnJur = true;
                        }
                        // TODO: ��������� �� ����� ��� ��� 10, 11, 12 ����
                    }

                    if (nLitzDolg == 2)// ���� ���. ����
                    {
                        // TODO: ������ ������ � �����
                        if (!Int32.TryParse(Convert.ToString(row["GOD"]), out nBirthYear))
                        {
                            nBirthYear = 0;
                        }

                        txtDateBornD = Convert.ToString(row["DATROZHD"]).Trim();
                        dtDatrozhd = DateTime.MaxValue;

                        // �������� ���� �������� � ��� ��������
                        if (DateTime.TryParse(txtDateBornD, out dtDate))
                        {
                            if (nBirthYear == 0) nBirthYear = dtDate.Year; // ���� �� ��� ���������� �.�., �� ����� �� ����
                            dtDatrozhd = dtDate;
                        }
                        else
                        {
                            dtDatrozhd = DateTime.MaxValue;
                            dtDate = DateTime.MaxValue;
                            if (nBirthYear == 0)
                            {
                                bNoYearBorn = true; // �������� � ����� ��������
                            }
                        }

                        
                    }

                    if (!bNoInnJur && !bNoYearBorn)// ���� ��� ������� �� �� ��� �� �� ���� ��������
                    {

                        m_cmd = new OleDbCommand();
                        m_cmd.Connection = DBFcon;

                        m_cmd.CommandText = "INSERT INTO " + tofind_name + " (NOMOSP, LITZDOLG, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FL_OKON, OSNOKON, DATROZHD) VALUES (";

                        m_cmd.CommandText += Convert.ToString(nOsp);

                        m_cmd.CommandText += ", " + Convert.ToString(nLitzDolg);

                        txtNameDolg = cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 50);
                        m_cmd.CommandText += ", '" + txtNameDolg + "'";

                        txtZapros = cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 25);
                        m_cmd.CommandText += ", '" + txtZapros + "'";

                        // TODO: ��� ���� ����� �� GOD
                        m_cmd.CommandText += ", " + Convert.ToString(nBirthYear);

                        txtNomspi = Convert.ToString(row["NOMSPI"]).Trim();
                        if (!Int32.TryParse(txtNomspi, out nNomspi))
                        {
                            nNomspi = 0;
                            txtNomspi = "0";
                        }

                        m_cmd.CommandText += ", " + Convert.ToString(nNomspi);

                        txtNOMIP = Convert.ToString(row["IPNO_NUM"]).Trim(); // ��� ����� �� - ���������� (����� ������� ��� ����������)
                        Decimal nNOMIP = 0;
                        if (!Decimal.TryParse(txtNOMIP, out nNOMIP))
                        {
                            nNOMIP = 0;
                        }

                        txtNPP = Convert.ToInt32(nNOMIP).ToString();

                        m_cmd.CommandText += ", " + txtNPP;

                        if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                        {
                            sum = 0;
                        }

                        m_cmd.CommandText += ", " + sum.ToString("F2").Replace(',', '.');

                        txtVIDVZISK = cutEnd(Convert.ToString(row["VIDVZISK"]).Trim(), 100);
                        m_cmd.CommandText += ", '" + txtVIDVZISK + "'";

                        m_cmd.CommandText += ", '" + txtInnOrg + "'";

                        if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))
                        {
                            DatZapr = DateTime.Today;
                        }

                        m_cmd.Parameters.Add(new OleDbParameter("DATZAPR", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR", System.Data.DataRowVersion.Original, DatZapr));
                        m_cmd.CommandText += ", ?";

                        txtAddr = cutEnd(Convert.ToString(row["ADDR"]).Trim(), 120);
                        m_cmd.CommandText += ", '" + txtAddr + "'";

                        m_cmd.CommandText += ", " + bZPRSPI.ToString();// FLZPRSPI

                        DatZapr1_param = DatZapr;
                        DatZapr2_param = DateTime.Today;

                        m_cmd.Parameters.Add(new OleDbParameter("DATZAPR1", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR1", System.Data.DataRowVersion.Original, DatZapr1_param));
                        m_cmd.CommandText += ", ?";

                        m_cmd.Parameters.Add(new OleDbParameter("DATZAPR2", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR2", System.Data.DataRowVersion.Original, DatZapr2_param));
                        m_cmd.CommandText += ", ?";

                        m_cmd.CommandText += ", " + bOKON.ToString();// Fl_OKON

                        //txtOsnokon = cutEnd(Convert.ToString(row["OSNOKON"]).Trim(), 250); - � ����� 2010 ���� ������ �������� ��������� ���������!!!
                        txtOsnokon = "";
                        m_cmd.CommandText += ", '" + txtOsnokon + "'";

                        m_cmd.Parameters.Add(new OleDbParameter("DATROZHD", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATROZHD", System.Data.DataRowVersion.Original, dtDatrozhd));
                        m_cmd.CommandText += ", ?";

                        m_cmd.CommandText += ')';
                        m_cmd.ExecuteNonQuery();
                        m_cmd.Dispose();

                        

                        //if (Convert.ToInt32(row["ID_DBTRCLS"]) == 95)
                        // ���� ��� �� (ID_DBTRCLS = 95) - �� �������� � DBF ��� � ������ �� ��. ����
                            if ((Convert.ToInt32(row["ID_DBTRCLS"]) == 95) && (txtInnOrg.Length >= 10) && (bOKON == 0))
                            {
                                m_cmd = new OleDbCommand();
                                m_cmd.Connection = DBFcon;

                                m_cmd.CommandText = "INSERT INTO " + tofind_name + " (NOMOSP, LITZDOLG, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FL_OKON, OSNOKON, DATROZHD) VALUES (";

                                m_cmd.CommandText += Convert.ToString(nOsp);

                                nLitzDolg = 1;// ��. ����
                                m_cmd.CommandText += ", " + Convert.ToString(nLitzDolg);

                                m_cmd.CommandText += ", '" + txtNameDolg + "'";

                                m_cmd.CommandText += ", '" + txtZapros + "'";

                                m_cmd.CommandText += ", " + Convert.ToString(nYear);

                                m_cmd.CommandText += ", " + nNomspi.ToString();

                                // ����� - ����� 5 ������... =((
                                // �������� ������������� ������ � ������ ���

                                m_cmd.CommandText += ", " + txtNPP;

                                m_cmd.CommandText += ", " + sum.ToString("F2").Replace(',', '.');

                                m_cmd.CommandText += ", '" + txtVIDVZISK + "'";

                                m_cmd.CommandText += ", '" + txtInnOrg + "'";

                                m_cmd.Parameters.Add(new OleDbParameter("DATZAPR", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR", System.Data.DataRowVersion.Original, DatZapr));
                                m_cmd.CommandText += ", ?";

                                m_cmd.CommandText += ", '" + txtAddr + "'";

                                m_cmd.CommandText += ", " + bZPRSPI.ToString();// FLZPRSPI

                                m_cmd.Parameters.Add(new OleDbParameter("DATZAPR1", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR1", System.Data.DataRowVersion.Original, DatZapr1_param));
                                m_cmd.CommandText += ", ?";

                                m_cmd.Parameters.Add(new OleDbParameter("DATZAPR2", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR2", System.Data.DataRowVersion.Original, DatZapr2_param));
                                m_cmd.CommandText += ", ?";

                                m_cmd.CommandText += ", " + bOKON.ToString();// Fl_OKON

                                m_cmd.CommandText += ", '" + txtOsnokon + "'";

                                m_cmd.Parameters.Add(new OleDbParameter("DATROZHD", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATROZHD", System.Data.DataRowVersion.Original, dtDatrozhd));
                                m_cmd.CommandText += ", ?";

                                m_cmd.CommandText += ')';
                                m_cmd.ExecuteNonQuery();
                                m_cmd.Dispose();

                            }
                        // else - // ��� ���� ����� ������ ������ ������� �� ������. ������ �� ��� ��������� (��� ���� ��������).
                        return true;

                    }
                    else // ���� ���� ������ �� ���� �������� ��� ���
                    {
                        string txtResponse = "";
                        if (bNoInnJur)
                        {
                            txtResponse += "������ �� ��� ��������� ������������, ��� ��� ����������� ��������� ���� ��� ��� �������� - ��. ����.";
                        }

                        if (bNoYearBorn)
                        {
                            txtResponse += "������ �� ��� ��������� ������������, ��� ��� � �������� - ���. ���� �� ��������� �� ��� �� ���� ��������.";
                        }

                        decimal nStatus = 15;
                        string txtID = Convert.ToString(row["ZAPROS"]).Trim();
                        Decimal nID = 0;
                        if(!Decimal.TryParse(txtID, out nID)){
                            nID = -1;
                        }

                        txtResponse += " ZAPROS = " + txtID + "\n";
                        // TODO: �������� ����� ������� �� nID
                        

                        if(nID > 0){
                            // ��������� ����� ��� ������� ����������� - �� ���� ��� ���� ������
                            foreach (string txtOrg_id in Legal_List)
                            {
                                // �� ���� ����������� ������ ����� ���������
                                decimal nOrg_id = Convert.ToDecimal(txtOrg_id);
                                // �������� Agreement_ID �� ������ �����������
                                decimal nAgr_id = GetAgr_by_Org(nOrg_id);
                                
                                // �������� nAgent_id, nAgent_dept_id
                                decimal nAgent_id = GetAgent_ID(nAgr_id);
                                decimal nAgent_dept_id = GetAgentDept_ID(nAgr_id);
                                
                                // ������� ������ ��� ����. ����������� � ���������� �� ������ � ������������ �������

                                // �������� ��������� - �����, agent_agreement, agent_dept_code, agent_code, enity_name
                                txtAgreementCode = GetAgreement_Code(nAgr_id);
                                txtAgentCode = GetAgent_Code(nAgr_id);
                                txtAgentDeptCode = GetAgentDept_Code(nAgr_id);

                                
                                // ���� ������ ������ ��� �� ������� - �� ������
                                if (nErrorPackID == 0)
                                {
                                    //nErrorPackID = ID_CreateDX_PACK_I(con, 70, nAgent_id, nAgent_dept_id, nAgr_id, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);
                                    //WritePackLog(con, nErrorPackID, "���� ����� ������ ��� �������� �������, ������� ������������� ������� ��� �������� ������������ �������� (��� ���� �������� ��� ���).����� ������ � ������ ����������� ����, ��� ����� ��������� ������������ ������ ������ ��������� ��� ����� ����� �� ������� �����.");
                                    // TODO: �������� ����� ��� � ��������
                                    nErrorPackID = CreateLLog(conGIBDD, 1, -1, txtAgreementCode, 0, "���� ����� ������ ��� �������� �������, ������� ������������� ������� ��� �������� ������������ �������� (��� ���� �������� ��� ���).\n");
                                }
                                
                                string txtEntityName = GetLegal_Name(nOrg_id);

                                // ��� ����������� ����� - ������ �������������

                                InsertResponseIntTable(con, nID, txtResponse, DateTime.Now, nStatus, nOrg_id, ref iRewriteState, nErrorPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName);

                                // ����� � ��� ��� ������ �� ��������
                                WriteLLog(conGIBDD, nErrorPackID, txtResponse);
                                // ������� ++ ��� ���������� �������� � ���� ������
                                AppendLLogCount(conGIBDD, nErrorPackID, 1);

                                // ������ ����� �� ������ ��������� ������, ��� �� ��� ��������...
                                // �� �� ����� ���������� ���� ���� �����, ��� ��� ��� ����� ��������� ������ �������

                                // ��� ����������� ����� - ������ �������������
                                //InsertZaprosTo_PK_OSP(con, nID, txtResponse, DateTime.Now, nStatus, nOrg_id, ref iRewriteState);
                            }
                        }

                        return false;

                    }
                        
                }
                catch (OleDbException ole_ex)
                {
                    foreach (OleDbError err in ole_ex.Errors)
                    {
                        MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    //if (DBFcon != null) DBFcon.Close();
                    MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                    return false;
                }
            }
            else return false;
        }

        private bool InsertRowToDBF_SBER(DataRow row, decimal bZPRSPI, DateTime DatZapr1_param, DateTime DatZapr2_param, string tofind_name)
        {
            DateTime DatZapr = DateTime.Today;
            DateTime dtDate;
            Decimal nYear = DateTime.Today.Year;
            string txtDateBornD = "";
            MatchCollection myMatchColl;
            string txtNomspi = "";
            int nNomspi = 0;

            if (DBFcon != null)
            {
                try
                {

                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = DBFcon;

                    m_cmd.CommandText = "INSERT INTO " + tofind_name + " (FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2) VALUES (";

                    m_cmd.CommandText += "'" + cutEnd(Convert.ToString(row["FIOVK"]).Trim().ToUpper(), 100) + "'";

                    m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40) + "'";

                    nYear = 0;
                    txtDateBornD = Convert.ToString(row["GOD"]).Trim();
                    if (txtDateBornD.Length == 4)
                    {
                        myMatchColl = Regex.Matches(txtDateBornD, @"\b[12]\d\d\d");
                        if (myMatchColl.Count > 0)
                        {
                            nYear = Convert.ToInt32(txtDateBornD);
                        }
                    }
                    else
                    {
                        if (txtDateBornD.Length == 10)
                        {
                            if (DateTime.TryParse(txtDateBornD, out dtDate))
                            {
                                nYear = dtDate.Year;
                            }
                        }
                    }

                    
                    m_cmd.CommandText += ", " + Convert.ToString(nYear);

                    txtNomspi = Convert.ToString(row["NOMSPI"]).Trim();
                    if(!Int32.TryParse(txtNomspi, out nNomspi)){
                        nNomspi = 0;
                    }
                    
                    m_cmd.CommandText += ", " + Convert.ToString(nNomspi);

                    // ����� - ����� 5 ������... =((
                    // �������� ������������� ������ � ������ ���

                    String txtNOMIP = Convert.ToString(row["NOMIP"]).Trim();
                    if (txtNOMIP.Trim() != "")
                    {
                        String[] strings = txtNOMIP.Split('/');
                        m_cmd.CommandText += ", " + Convert.ToString(strings[2]);
                    }
                    else
                    {
                        m_cmd.CommandText += ", 0";
                    }


                    if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))
                    {
                        DatZapr = DateTime.Today;
                    }

                    m_cmd.Parameters.Add(new OleDbParameter("DATZAPR", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR", System.Data.DataRowVersion.Original, DatZapr));
                    m_cmd.CommandText += ", ?";

                    m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["ADDR"]).Trim(), 120) + "'";

                    m_cmd.CommandText += ", " + bZPRSPI.ToString();// FLZPRSPI

                    m_cmd.Parameters.Add(new OleDbParameter("DATZAPR1", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR1", System.Data.DataRowVersion.Original, DatZapr1_param));
                    m_cmd.CommandText += ", ?";

                    m_cmd.Parameters.Add(new OleDbParameter("DATZAPR2", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATZAPR2", System.Data.DataRowVersion.Original, DatZapr2_param));
                    m_cmd.CommandText += ", ?";

                    m_cmd.CommandText += ')';
                    m_cmd.ExecuteNonQuery();
                    m_cmd.Dispose();
                }
                catch (OleDbException ole_ex)
                {
                    foreach (OleDbError err in ole_ex.Errors)
                    {
                        MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                    }
                    if (m_cmd != null) m_cmd.Dispose();
                    return false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                    if (m_cmd != null) m_cmd.Dispose();
                    return false;
                }
                return true;
            }
            else return false;
        }

        private static String cutEnd(string txtStr, int iLen)
        {
            if (txtStr.Length > iLen) txtStr = txtStr.Substring(0, iLen - 1);

            string[] StrColl = txtStr.Split(new char[] { '\n', '\t', '\b', '\r' });

            int StrCollLength = StrColl.Length;

            txtStr = "";

            for (int i = 0; i < StrCollLength; i++)
            {
                txtStr += StrColl[i];
            }

            return txtStr;
        }

        private static char monthCode(DateTime dtDate)
        {
            char[] monthCodes = new char[] {' ', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C' };
            return monthCodes[dtDate.Month];
        }

        private static char fileCode(int iNum)
        {
            char[] fileCodes = new char[] { '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };

            if (iNum > 34)
            {
                return '0';
            }
            else
            {
                return fileCodes[iNum];
            }
                
        }

        private void btnWriteDBF_Click(object sender, EventArgs e)
        {

            //btnMakeZapros.Enabled = false;
            btnWriteDBF.Enabled = false;

            if (bReadFromCopy)
            {
                con = new OleDbConnection(constr2);
            }

            con = new OleDbConnection(constrRDB);

            // ���������� ��� ������ �� ����������� � ��������
            // ��� ��������� ��������
            string txtUpdateSql = "";
            decimal nOrg_id = 0;



            // !��������������� ��� UPDATE - �.�. ��� ��� ����� ����� � ����� ��� ����������� ������
            
            //// ���������� ��������
            //nOrg_id = 86200999999005;
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 23 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 3 and req_type.sndl_contr_id = " + nOrg_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// ��� �������������� ��������
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 11 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 2 and req_type.sndl_contr_id = " + nOrg_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// ���������� ��� ��� � ������
            //foreach (string txtOrg_id in Legal_List)
            //{
            //    if (Decimal.TryParse(txtOrg_id, out nOrg_id))
            //    {
            //        txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 23 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 3 and req_type.sndl_contr_id = " + txtOrg_id + ")";
            //        UpdateSqlExecute(con, txtUpdateSql);

            //        // ��� �������������� ��������
            //        txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 11 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 2 and req_type.sndl_contr_id = " + txtOrg_id + ")";
            //        UpdateSqlExecute(con, txtUpdateSql);
            //    }
            //}

            //DT_doc = GetDataTableFromFB("SELECT DISTINCT UPPER(a.NAME_D) as FIOVK, a.NUM_IP as ZAPROS, a.VIDD_KEY as LITZDOLG, a.DATE_BORN_D as GOD, c.PRIMARY_SITE as NOMSPI, a.NUM_IP as NOMIP, a.SUM_ as SUMMA, a.WHY as VIDVZISK, a.INND as INNORG, b.DATE_DOC as DATZAPR, a.ADR_D as ADDR,a.TEXT_PP as OSNOKON, a.PK as FK_IP, a.PK_ID as FK_ID, a.uscode, a.FIO_SPI FROM IP a left join s_users c on (a.uscode=c.uscode) LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.DATE_IP_OUT is null and b.KOD = 1011 and (b.DATE_DOC >= '" + DatZapr1.ToShortDateString() + "' AND b.DATE_DOC <= '" + DatZapr2.ToShortDateString() + "') and a.ssd is null and a.ssv is null", "TOFIND");

            // �.�. ������ � ����� ������ ��� ����, �� ����� ������ �� ���� ��������
            // DT_doc_jur = GetDataTableFromFB("select 1 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 31 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 1 or (ip.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT  where ncc_parent_id = 1)))", "TOFIND");
            // DT_doc_jur = GetDataTableFromFB("select 1 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 31 and zapr_d.docstatusid = 2 and ip_d.docstatusid = 9 and (ip.ID_DBTRCLS = 1 or (ip.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 1)))", "TOFIND");
            //DT_doc_jur = GetDataTableFromFB("select pack.id as pack_id, 1 LITZDOLG, d_req.id ZAPROS, req.IPNO_NUM, req.div, req.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, d_req.doc_date DATZAPR,  req.ID_DBTR_ADR ADDR, req.ID_DBTR_BORN DATROZHD, req.ID_DBTRCLS,  req.DBTR_BORN_YEAR GOD, req.ID_DEBTSUM SUMMA, req.ID_DBTR_INN INNORG,  d_req.doc_number, req.ID_DEBTCLS_NAME VIDVZISK from dx_pack_o packo left join dx_pack pack on pack.id = packo.id join sendlist sl on pack.id = sl.dx_pack_id join o_ip req on sl.sendlist_o_id = req.id join document d_req on req.id = d_req.id join document ip_d on d_req.parent_id = ip_d.id join document dpack on pack.id = dpack.id join SPI on req.IP_EXEC_PRIST = spi.SUSER_ID where dpack.docstatusid = 23  and pack.agreement_id = 30 and packo.has_been_sent is null and d_req.docstatusid != 19 and d_req.docstatusid != 15 and (req.ID_DBTRCLS = 1 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 1)))", "TOFIND");
            DT_doc_jur = GetDataTableFromFB("select 30 agreement_id, ext_request_id,  pack_id,  1 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID where mvv_agreement_code = 30 and processed = 0 and (req.ID_DBTRCLS = 1 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 1)))", "TOFIND");

            // ��������� ��������� ����� ������� ������ �� ������ �� ����������? -
            // � ������������ ����������� - ��� ��������� ���������, �.�. � ��� ������ 1 ���� � ���������
            // ��! ���� ����� ������ ����������� ������� � ������������ - ��� ��� ������������ ��������� ����������� �� ������ agreement_id = 30 (���� �����������)

            // DT_doc_fiz = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 31 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_NSI_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            //DT_doc_fiz = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 31 and zapr_d.docstatusid = 2 and ip_d.docstatusid = 9 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            //DT_doc_fiz = GetDataTableFromFB("select pack.id as pack_id, 2 LITZDOLG, d_req.id ZAPROS, req.IPNO_NUM, req.div, req.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, d_req.doc_date DATZAPR,  req.ID_DBTR_ADR ADDR, req.ID_DBTR_BORN DATROZHD, req.ID_DBTRCLS,  req.DBTR_BORN_YEAR GOD, req.ID_DEBTSUM SUMMA, req.ID_DBTR_INN INNORG,  d_req.doc_number, req.ID_DEBTCLS_NAME VIDVZISK from dx_pack_o packo left join dx_pack pack on pack.id = packo.id join sendlist sl on pack.id = sl.dx_pack_id join o_ip req on sl.sendlist_o_id = req.id join document d_req on req.id = d_req.id join document ip_d on d_req.parent_id = ip_d.id join document dpack on pack.id = dpack.id join SPI on req.IP_EXEC_PRIST = spi.SUSER_ID where dpack.docstatusid = 23  and pack.agreement_id = 30 and packo.has_been_sent is null  and d_req.docstatusid != 19 and d_req.docstatusid != 15  and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            DT_doc_fiz = GetDataTableFromFB("select 30 agreement_id, ext_request_id,  pack_id,  2 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID where mvv_agreement_code = 30 and processed = 0 and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            
            
            // 30 - ������ �����������

            if (bReadFromCopy)
            {
                con = new OleDbConnection(constr1);
            }

            //btnMakeZapros.Enabled = true;
            //btnWriteDBF.Enabled = true;

            int iDocCnt = 0;
            if (DT_doc_jur != null) iDocCnt += DT_doc_jur.Rows.Count;
            if (DT_doc_fiz != null) iDocCnt += DT_doc_fiz.Rows.Count;

            lblReadRowsValue.Text = iDocCnt.ToString();
            
            
            Int64 cnt;
            if (bDateFolderAdd)
            {
                CreatePathWithDate(cred_org_path);
            }
            else
            {
                FolderExist(cred_org_path);
            }

            // ��� ���� true - bVFP
            // ����� ���-�� ������ ��� �� reg ������������� ��� ���� �� ���� ��� ������ � DBF
            cnt = WriteToDBF(true, fullpath, "tofind1.dbf", DatZapr1, DatZapr2, "tofind.dbf");
            //cnt = WriteToDBF(false, fullpath, "tofind1.dbf", DatZapr1, DatZapr2, "tofind.dbf");
            lblWriteRowsValue.Text = cnt.ToString();
            btnWriteDBF.Enabled = true;
        }

        private string makenewSberFileName()
        {
            // RDDMFFFF
            string txtRes = "R";
            txtRes += DateTime.Today.Day.ToString("D2"); // DD
            txtRes += monthCode(DateTime.Today); // M
            //txtRes += txtSberFilialCode; // FFFF
            return txtRes;
        }

        private string makenewSberFileExt(int iFileNum)
        {
            // .NXX
            string txtRes = ".";
            txtRes += fileCode(iFileNum); // N
            txtRes += Convert.ToInt32(GetOSP_Num()).ToString("D2"); // XX

            return txtRes;
        }



        private void Copy(String source, String destdir)
        {
            if (!Directory.Exists(destdir))
                Directory.CreateDirectory(destdir);

            String[] Files = Directory.GetFiles(destdir, DateTime.Today.ToShortDateString() + "*" + ".dbf");
            if (Files.Length == 0)
            {
                File.Copy(source, string.Format(@"{0}\{1}", destdir, DateTime.Today.ToShortDateString() + ".dbf"));
            }
            if (Files.Length == 1)
            {
                Array.Sort(Files);
                File.Copy(source, string.Format(@"{0}\{1}", destdir, DateTime.Today.ToShortDateString() + "." + "1.dbf"));
            }
            if (Files.Length == 2)
            {
                Array.Sort(Files);
                int num = Convert.ToInt32(Files[0].Substring(Files[0].Length - 5, 1)) + 1;
                File.Copy(source, string.Format(@"{0}\{1}", destdir, DateTime.Today.ToShortDateString() + "." + num + ".dbf"));
            }
        }

        private void btnMakeZapros_Click(object sender, EventArgs e)
        {
            int zapros = 0;
            
            // � ��� - ���������� progress-bar � ���� � ��������� �����
            prbWritingDBF.Value = 0;
            prbWritingDBF.Step = 1;

            if (DT_doc != null)
            {
                prbWritingDBF.Maximum = DT_doc.Rows.Count;
                // ��������� �������� - ���� �� ����������� �� ���
                zapros = InsertZapros_kred_org(DT_doc, DatZapr1, DatZapr2, true);
            }
            //lblZaprosMadevalue.Text = Convert.ToString(zapros);

        }

        private static ArrayList ReadPaths(string FromFilename)
        {
            ArrayList Filepaths = new ArrayList();
            using (StreamReader sr = new StreamReader(FromFilename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    Filepaths.Add(line);
                }
            }
            return Filepaths;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //btnMakeZapros.Enabled = false;
            //btnWriteDBF.Enabled = false;

            //btnWriteKtfomsDBF.Enabled = false;

            //btnWritePensDBF.Enabled = false;

            //btnWritePotdDBF.Enabled = false;

            cbxItems = new ArrayList();
            // �������� �������� �����������-���������
            Legal_Name_List = (String[])Legal_List.Clone();
            Legal_�onv_List = (String[])Legal_List.Clone();
            //int code = 0;
            Decimal code = 0;
            //for (int i = 1; i < Legal_Name_List.Length; i++)
            for (int i = 0; i < Legal_Name_List.Length; i++)
            {
                //code = Convert.ToInt32((Legal_List[i]).Trim());
                code = Convert.ToDecimal((Legal_List[i]).Trim());
                Legal_Name_List[i] = GetLegal_Name(code);
                Legal_�onv_List[i] = GetLegal_Conv(code);
                cbxItems.Add(new ComboItem(code, Legal_Name_List[i]));

            }

            cbxOrg.DataSource = cbxItems;
            cbxOrg.ValueMember = "Id";
            cbxOrg.DisplayMember = "Name";
            //cbxOrg.SelectedValue = Convert.ToInt32((Legal_List[0]).Trim());
            
            //string OptionsFilename = "options.txt";
            //ArrayList Options = ReadPaths(OptionsFilename);
            //string txtDateIntTableDplmnt = Convert.ToString(Options[0]);
            //if(!DateTime.TryParse(txtDateIntTableDplmnt, out dtIntTablesDeplmntDate))
            //{
            //    dtIntTablesDeplmntDate = DateTime.MaxValue;
            //}
            

        }


        private void prbWritingDBF_Click(object sender, EventArgs e)
        {

        }

        string ConvertDOS(string strdos)
        {
            string struni = "";

            char[] infc = new char[strdos.Length];
            infc = strdos.ToCharArray();
            int t;

            for (t = 0; t < infc.Length; t++)
            {
                switch (Convert.ToInt32(infc[t]))
                {
                    case 1026: infc[t] = '�'; break;
                    case 1027: infc[t] = '�'; break;
                    case 8218: infc[t] = '�'; break;
                    case 1107: infc[t] = '�'; break;
                    case 8222: infc[t] = '�'; break;
                    case 8230: infc[t] = '�'; break;
                    case 1088: infc[t] = '�'; break;
                    case 8224: infc[t] = '�'; break;
                    case 8225: infc[t] = '�'; break;
                    case 8364: infc[t] = '�'; break;
                    case 8240: infc[t] = '�'; break;
                    case 1033: infc[t] = '�'; break;
                    case 8249: infc[t] = '�'; break;
                    case 1034: infc[t] = '�'; break;
                    case 1036: infc[t] = '�'; break;
                    case 1035: infc[t] = '�'; break;
                    case 1039: infc[t] = '�'; break;
                    case 1106: infc[t] = '�'; break;
                    case 8216: infc[t] = '�'; break;
                    case 8217: infc[t] = '�'; break;
                    case 8220: infc[t] = '�'; break;
                    case 8221: infc[t] = '�'; break;
                    case 8226: infc[t] = '�'; break;
                    case 8211: infc[t] = '�'; break;
                    case 8212: infc[t] = '�'; break;
                    case 152: infc[t] = '�'; break;
                    case 8482: infc[t] = '�'; break;
                    case 1113: infc[t] = '�'; break;
                    case 8250: infc[t] = '�'; break;
                    case 1114: infc[t] = '�'; break;
                    case 1116: infc[t] = '�'; break;
                    case 1115: infc[t] = '�'; break;
                    case 1119: infc[t] = '�'; break;

                    case 160: infc[t] = '�'; break;
                    case 1038: infc[t] = '�'; break;
                    case 1118: infc[t] = '�'; break;
                    case 1032: infc[t] = '�'; break;
                    case 164: infc[t] = '�'; break;
                    case 1168: infc[t] = '�'; break;
                    case 1089: infc[t] = '�'; break;
                    case 166: infc[t] = '�'; break;
                    case 167: infc[t] = '�'; break;
                    case 1025: infc[t] = '�'; break;
                    case 169: infc[t] = '�'; break;
                    case 1028: infc[t] = '�'; break;
                    case 171: infc[t] = '�'; break;
                    case 172: infc[t] = '�'; break;
                    case 173: infc[t] = '�'; break;
                    case 174: infc[t] = '�'; break;
                    case 1031: infc[t] = '�'; break;
                    case 1072: infc[t] = '�'; break;
                    case 1073: infc[t] = '�'; break;
                    case 1074: infc[t] = '�'; break;
                    case 1075: infc[t] = '�'; break;
                    case 1076: infc[t] = '�'; break;
                    case 1077: infc[t] = '�'; break;
                    case 1078: infc[t] = '�'; break;
                    case 1079: infc[t] = '�'; break;
                    case 1080: infc[t] = '�'; break;
                    case 1081: infc[t] = '�'; break;
                    case 1082: infc[t] = '�'; break;
                    case 1083: infc[t] = '�'; break;
                    case 1084: infc[t] = '�'; break;
                    case 1085: infc[t] = '�'; break;
                    case 1086: infc[t] = '�'; break;
                    case 1087: infc[t] = '�'; break;

                    //case : infc[t]=''; break;case : infc[t]=''; break;case : infc[t]=''; break;

                }
                struni += infc[t];
            }

            //struni = Convert.ToString(infc);

            return struni;
        }


        private bool AppendZaprosIn_PK_OSP(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id)
        {

            OleDbCommand cmdInsMVV_RESPONSE;
            OleDbTransaction tran = null;

            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                // �������� MVV_RESPONSE

                cmdInsMVV_RESPONSE = new OleDbCommand();
                cmdInsMVV_RESPONSE.Connection = con;
                cmdInsMVV_RESPONSE.Transaction = tran;
                cmdInsMVV_RESPONSE.CommandText = "update MVV_RESPONSE SET ACT_DATA = :ACT_DATA, DATA_STR = DATA_STR || ' ' || '" + txtOtvet + "' WHERE ID = :ID";

                cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":ACT_DATA", dtDatOtv));
                cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));


                if (cmdInsMVV_RESPONSE.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error appending row in MVV_RESPONSE table id = " + nID.ToString());
                    throw ex;
                }

                tran.Commit();
                con.Close();

                //SetDocumentStatus(nID, 19);// ���������� ������ ������� ����� ��� �������


                return true;
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }
        
        // ���������� �������� ����������, ������� �� ����� ���� ����� ��� �� �����, �� �� �������� � insert �������
        private bool UpdateZaprosIn_PK_OSP(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id)
        {

            OleDbCommand cmdInsMVV_RESPONSE;
            OleDbTransaction tran = null;

            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                // �������� MVV_RESPONSE

                cmdInsMVV_RESPONSE = new OleDbCommand();
                cmdInsMVV_RESPONSE.Connection = con;
                cmdInsMVV_RESPONSE.Transaction = tran;
                cmdInsMVV_RESPONSE.CommandText = "update MVV_RESPONSE SET ACT_DATA = :ACT_DATA, DATA_STR = :DATA_STR WHERE ID = :ID";

                cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":ACT_DATA", dtDatOtv));
                cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":DATA_STR", txtOtvet));
                cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));


                if (cmdInsMVV_RESPONSE.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error updating row in MVV_RESPONSE table id = " + nID.ToString());
                    throw ex;
                }

                tran.Commit();
                con.Close();

                //SetDocumentStatus(nID, 19);// ���������� ������ ������� ����� ��� �������


                return true;
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }


        private bool WritePackLog(OleDbConnection connection, decimal nPackID, string txtText)
        {
            //string txtSql = "Update DX_PACK set PLAIN_LOG = PLAIN_LOG || '" + txtText + "' where ID = " + nPackID.ToString();
            string txtSql = "Update PACK_LOGS set LOG = LOG || '" + txtText + "' where PACK_ID = " + nPackID.ToString();
            return UpdateSqlExecute(connection, txtSql);
        }

        private bool UpdateSqlExecute(OleDbConnection con, string txtUpdSql)
        {

            OleDbCommand cmd;
            OleDbTransaction tran = null;

            try
            {
                if ((con == null) || (con.State.Equals(ConnectionState.Closed)))
                {
                    con.Open();
                }

                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                cmd = new OleDbCommand();
                cmd.Connection = con;
                cmd.Transaction = tran;
                cmd.CommandText = txtUpdSql;
                cmd.ExecuteNonQuery();
                tran.Commit();
                con.Close();

                return true;
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }


        // ������� ������ �� ������� I_ID
        private bool AlterIndxI_ID(OleDbConnection con, bool flActive)
        {

            OleDbCommand cmd;
            OleDbTransaction tran = null;

            try
            {
                if ((con == null) || (con.State.Equals(ConnectionState.Closed)))
                {
                    con.Open();
                }

                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                cmd = new OleDbCommand();
                cmd.Connection = con;
                cmd.Transaction = tran;

                if (flActive)
                {
                    cmd.CommandText = "alter index I_ID_IDX1 active";
                }
                else
                {
                    cmd.CommandText = "alter index I_ID_IDX1 inactive";
                }

                cmd.ExecuteNonQuery();

                //if (cmd.ExecuteNonQuery() == -1)
                //{
                //    Exception ex = new Exception("Error deleteting altering I_ID");
                //    throw ex;
                //}

                tran.Commit();
                con.Close();

                return true;
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }

        private bool DeleteUsedGibddPlat(OleDbConnection conG, bool flDeleteUsed, bool flDeleteOld)
        {

            OleDbCommand cmdD, cmdOldD;
            OleDbTransaction tran = null;

            try
            {
                if((conG == null) || (conG.State.Equals(ConnectionState.Closed))){
                    conG.Open();
                }

                tran = conG.BeginTransaction(IsolationLevel.ReadCommitted);

                // ���� ���� ������� �������� �����
                if (flDeleteUsed)
                {

                    // ������� ��� ��� FL_USE = 1

                    cmdD = new OleDbCommand();
                    cmdD.Connection = conG;
                    cmdD.Transaction = tran;
                    cmdD.CommandText = "DELETE FROM GIBDD_PLATEZH WHERE FL_USE = 1";

                    if (cmdD.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error deleteting rows from GIBDD_PLATEZH.");
                        throw ex;
                    }
                }
                
                // ���� ���� ������� ������ �����
                if (flDeleteOld)
                {
                    // �������� ����� ������� ���������� � ��� ��, ������� ���� ������ 1 ��� � 10 ���� ����� (10 ���� �� ���������� � ����)
                    cmdOldD = new OleDbCommand();
                    cmdOldD.Connection = conG;
                    cmdOldD.Transaction = tran;
                    //cmdOldD.CommandText = "DELETE FROM GIBDD_PLATEZH WHERE DATID < '" + DateTime.Today.AddYears(-1).AddDays(-10).ToShortDateString() + "'";
                    cmdOldD.CommandText = "DELETE FROM GIBDD_PLATEZH WHERE DATID < '" + DateTime.Today.AddYears(-2).AddDays(-10).ToShortDateString() + "'";

                    if (cmdOldD.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error deleteting rows from GIBDD_PLATEZH.");
                        throw ex;
                    }
                }

                tran.Commit();
                conG.Close();

                return true;
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (conG != null)
                {
                    conG.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (conG != null)
                {
                    conG.Close();
                }
            }
            return false;
        }

        private bool UpdateGibddPlatezh(OleDbConnection conG, string txtNumber, int iValue)
        {

            OleDbCommand cmdU;
            OleDbTransaction tran = null;

            try
            {
                conG.Open();
                tran = conG.BeginTransaction(IsolationLevel.ReadCommitted);

                // �������� MVV_DATA_RESPONSE

                cmdU = new OleDbCommand();
                cmdU.Connection = conG;
                cmdU.Transaction = tran;
                cmdU.CommandText = "update GIBDD_PLATEZH SET FL_USE = :FL_USE WHERE NUMBER = :NUMBER";

                cmdU.Parameters.Add(new OleDbParameter(":FL_USE", iValue));
                cmdU.Parameters.Add(new OleDbParameter(":NUMBER", txtNumber));


                if (cmdU.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error updating row in GIBDD_PLATEZH table NUMBER = " + txtNumber);
                    throw ex;
                }

                tran.Commit();
                conG.Close();

                return true;
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (conG != null)
                {
                    conG.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (conG != null)
                {
                    conG.Close();
                }
            }
            return false;
        }

        private bool UpdateZaprosIn_PK_OSP(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv)
        {

            OleDbCommand cmdInsMVV_DATA_RESPONSE;
            OleDbTransaction tran = null;

            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                // �������� MVV_DATA_RESPONSE

                cmdInsMVV_DATA_RESPONSE = new OleDbCommand();
                cmdInsMVV_DATA_RESPONSE.Connection = con;
                cmdInsMVV_DATA_RESPONSE.Transaction = tran;
                cmdInsMVV_DATA_RESPONSE.CommandText = "update MVV_DATA_RESPONSE SET ACT_DATA = :ACT_DATA, DATA_STR = :DATA_STR WHERE ID = :ID";

                cmdInsMVV_DATA_RESPONSE.Parameters.Add(new OleDbParameter(":ACT_DATA", dtDatOtv));
                cmdInsMVV_DATA_RESPONSE.Parameters.Add(new OleDbParameter(":DATA_STR", txtOtvet));
                cmdInsMVV_DATA_RESPONSE.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));
                

                if (cmdInsMVV_DATA_RESPONSE.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error updating row in MVV_DATA_RESPONSE table id = " + nID.ToString());
                    throw ex;
                }

                tran.Commit();
                con.Close();

                //SetDocumentStatus(nID, 19);// ���������� ������ ������� ����� ��� �������


                return true;
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }

        bool InsertPackDocs(decimal nPackID, decimal nDocID)
        {
            OleDbCommand cmd;
            OleDbTransaction tran = null;
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                cmd = new OleDbCommand();
                cmd.Connection = con;
                cmd.Transaction = tran;
                cmd.CommandText = "insert into PACK_DOCS (DX_PACK_ID, DOCUMENT_ID)";
                cmd.CommandText += " VALUES (:PACK_ID ,:DOC_ID)";

                cmd.Parameters.Add(new OleDbParameter(":PACK_ID", Convert.ToDecimal(nPackID)));
                cmd.Parameters.Add(new OleDbParameter(":DOC_ID", Convert.ToDecimal(nDocID)));

                if (cmd.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to pack_docs table pack_id = " + nPackID.ToString());
                    throw ex;
                }
                tran.Commit();
                con.Close();
                return true;
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }

            return false;
        }

        // ���. �������
        // �������� ����� �������� ������� ��� �������� DX_PACK_ID
        private bool InsertResponseIntTable(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id, ref int iRewriteState)
        {
            // string txtAgentCode, string txtAgentDeptCode, string txtAgentAgreementCode, string txtEntityName = ������ ���, �����-�� ���� ��� �����
            return InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, entt_id, ref iRewriteState, 0, " ", " ", " ", " ");

        }

        // ���. ������� - ��� �� ���� �������� � ��������
        // ��� � iRewriteState
        // ���������� �������� ���������� ������ � ��������������, ����� MVV_EXTERNAL_RESPONSE � out iRewriteState
        private bool InsertResponseIntTable(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id, ref int iRewriteState, decimal nDX_PACK_ID, string txtAgentCode, string txtAgentDeptCode, string txtAgentAgreementCode, string txtEntityName)
        {
            // ��� ����� entt_id - ������ ��� ��� ��� �� �� ����������� Legal
            // �������� EXT_INPUT_HEADER < - > EXT_RESPONSE (����� ���� � ������)
            OleDbCommand cmd, cmdEXT_INPUT_HEADER, cmdCheckAnsw, cmdEXT_RESPONSE, cmdPackDocs, cmdDocNumber;
            Decimal newID, prevID;
            OleDbTransaction tran = null;
            decimal nAgreementID = 0;
            decimal nAgent_dept_id = 0;
            decimal nAgent_id = 0;

            //iRewriteState = 
            //1 - ������� ����� - ����������� ������� � ������������ 
            //2 - ��������
            //3 - ������������
            //4 - ����������
            //20 - �������� ���
            //21 - ������������ ���
            //22 - ���������� ���, ������� �������

            try
            {
                // ��� - ������ �� ����� ID, ������ Code ������� � ���������� ������� ��������
                // TODO: ������ ���� ����� ����� ������� (������ 0 - ����� ���������� �����)
                // ���� �������� �� ����� ����� ��������, �� �������� ���������
                //if (nDX_PACK_ID > 0)
                //{
                //    DataTable dtParams = GetPackParams(con, nID, entt_id);
                //    if ((dtParams != null) && (dtParams.Rows.Count > 0))
                //    {
                //        nAgreementID = Convert.ToDecimal(dtParams.Rows[0]["agreement_id"]);
                //        nAgent_dept_id = Convert.ToDecimal(dtParams.Rows[0]["agent_dept_id"]);
                //        nAgent_id = Convert.ToDecimal(dtParams.Rows[0]["agent_id"]);
                //    }

                //}

                newID = 0;
                prevID = 0;
                
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                // ��������� ��� �� ���� ��������� ������ �� ���� ������
                // select id from MVV_EXTERNAL_RESPONSE ext left join MVV_RESPONSE resp on ext.id = resp.id left join DOCUMENT doc on ext.id = doc.id where doc.parent_id =:ID and resp.entity_id = :ENTITY_ID
                cmdCheckAnsw = new OleDbCommand("select first 1 ext.id from MVV_EXTERNAL_RESPONSE ext join MVV_RESPONSE resp on ext.id = resp.id join DOCUMENT doc on ext.id = doc.id where doc.parent_id =:ID and resp.entity_id = :ENTITY_ID", con, tran);
                cmdCheckAnsw.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));
                cmdCheckAnsw.Parameters.Add(new OleDbParameter(":ENTITY_ID", Convert.ToDecimal(entt_id)));
                prevID = Convert.ToDecimal(cmdCheckAnsw.ExecuteScalar());

                // ��� ���� ������������� iRewriteState

                // ���� ������ ������������ - ������ �� ��������� - ������ � ������� �� �����
                // - ����� ������� ����� ��� � ���������� ��� ����������� ��� ��� �� �����
                //if (prevID <= 0)
                //{

                    // �������� ����� ����
                    cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
                    newID = Convert.ToDecimal(cmd.ExecuteScalar());

                    // �������� DOCUMENT
                    cmdEXT_INPUT_HEADER = new OleDbCommand();
                    cmdEXT_INPUT_HEADER.Connection = con;
                    cmdEXT_INPUT_HEADER.Transaction = tran;
                    cmdEXT_INPUT_HEADER.CommandText = "insert into EXT_INPUT_HEADER (ID, METAOBJECTNAME, PROCEED, PACK_NUMBER, EXTERNAL_KEY, AGENT_CODE, AGENT_DEPT_CODE, AGENT_AGREEMENT_CODE, DATE_IMPORT)";
                    cmdEXT_INPUT_HEADER.CommandText += " VALUES (:ID ,'EXT_RESPONSE', 0, :PACK_NUMBER, :EXTERNAL_KEY, :AGENT_CODE, :AGENT_DEPT_CODE, :AGENT_AGREEMENT_CODE, :DATE_IMPORT)";

                    cmdEXT_INPUT_HEADER.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));


                    // 20120828 ������� ������� ���� �� 8 � �����
                    string txtExtPack = Convert.ToString(nDX_PACK_ID);
                    if (txtExtPack.Length > 8)
                    {
                        txtExtPack = txtExtPack.Substring(txtExtPack.Length - 8, 8);
                    }
                    decimal nExtPack = 0;
                    Decimal.TryParse(txtExtPack, out nExtPack);
                    
                    cmdEXT_INPUT_HEADER.Parameters.Add(new OleDbParameter(":PACK_NUMBER", nExtPack));
                    
                    // ���������� - ������� ����� ������ ����� ��������� � ���������� ���, ���� �������� ������, �� ����� ������ ���������
                    // cmdEXT_INPUT_HEADER.Parameters.Add(new OleDbParameter(":PACK_NUMBER", nDX_PACK_ID));

                    cmdEXT_INPUT_HEADER.Parameters.Add(new OleDbParameter(":EXTERNAL_KEY", Convert.ToString(newID)));
                    cmdEXT_INPUT_HEADER.Parameters.Add(new OleDbParameter(":AGENT_CODE", txtAgentCode));
                    cmdEXT_INPUT_HEADER.Parameters.Add(new OleDbParameter(":AGENT_DEPT_CODE", txtAgentDeptCode));
                    cmdEXT_INPUT_HEADER.Parameters.Add(new OleDbParameter(":AGENT_AGREEMENT_CODE", txtAgentAgreementCode));
                    cmdEXT_INPUT_HEADER.Parameters.Add(new OleDbParameter(":DATE_IMPORT", DateTime.Today));


                    if (cmdEXT_INPUT_HEADER.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to EXT_INPUT_HEADER table parent_id = " + nID.ToString());
                        throw ex;
                    }

                    // �������� MVV_RESPONSE
                    // � 14 ������ ���������� ��� ��������� EXAD_AGENT_ID, EXAD_DEPT_ID, OUTER_AGREEMENT_ID
                    // � � 68-� ������ ��� ����-�� ������� � ���������� ��� ���

                    cmdEXT_RESPONSE = new OleDbCommand();
                    cmdEXT_RESPONSE.Connection = con;
                    cmdEXT_RESPONSE.Transaction = tran;
                    cmdEXT_RESPONSE.CommandText = "insert into EXT_RESPONSE (ID, ENTITY_NAME, RESPONSE_DATE, REQUEST_NUM, REQUEST_ID, DATA_STR)"; //, EXAD_AGENT_ID, EXAD_DEPT_ID, OUTER_AGREEMENT_ID
                    cmdEXT_RESPONSE.CommandText += "  VALUES (:ID ,:ENTITY_NAME, :RESPONSE_DATE, :REQUEST_NUM, :REQUEST_ID, :DATA_STR)"; //, :EXAD_AGENT_ID, :EXAD_DEPT_ID, :OUTER_AGREEMENT_ID)";

                    cmdEXT_RESPONSE.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                    cmdEXT_RESPONSE.Parameters.Add(new OleDbParameter(":ENTITY_NAME", txtEntityName));
                    cmdEXT_RESPONSE.Parameters.Add(new OleDbParameter(":RESPONSE_DATE", DateTime.Today));

                    string txtReqDocNumber = "";
                    cmdDocNumber = new OleDbCommand("select DOC_NUMBER from document where id = " + nID.ToString() , con, tran);
                    txtReqDocNumber = Convert.ToString(cmdDocNumber.ExecuteScalar());

                    cmdEXT_RESPONSE.Parameters.Add(new OleDbParameter(":REQUEST_NUM", txtReqDocNumber));
                    cmdEXT_RESPONSE.Parameters.Add(new OleDbParameter(":REQUEST_ID", Convert.ToDecimal(nID)));

                    cmdEXT_RESPONSE.Parameters.Add(new OleDbParameter(":DATA_STR", txtOtvet));
                                       
                    //cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":EXAD_AGENT_ID", nAgent_id));
                    //cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":EXAD_DEPT_ID", nAgent_dept_id));
                    //cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":OUTER_AGREEMENT_ID", nAgreementID));

                    // TODO: OUTER_AGREEMENT_ID, OUTER_AGREEMENT_NAME - ����������

                    if (cmdEXT_RESPONSE.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to EXT_RESPONSE table id = " + nID.ToString());
                        throw ex;
                    }

                    tran.Commit();
                    con.Close();

                    // ��������� - ���� �� ��� ������, �.�. �������� ��� ������� ���� ����� ������� ������ ������������ ���. �������
                    // ������� ���� �� ������ - ���� ������� ������
                    SetDocumentStatus(nID, 19);// ���������� ������ ������� ����� ��� �������

                    //SetDocumentStatus(nID, 15);// ���������� ������ ���������� � �������
                    
                    return true;
                //}
                //// TODO: ������ ��� ������, ���� ��� ���� ����� �� ���� ������...
                //else
                //{
                //    tran.Rollback();
                //    con.Close();

                //    // ��������� ������ ���������� - � �� ����������� ������� ��������.
                //    // TODO: ������ - ������ ��������� �� ��� ������ - ������ ��� ���� �������� ��������
                //    // - ���� � ext_response.proceed = 1, �� �������������� ��� ������� ������ �� �����
                //    // � �� ������ update ������ ��, ����� ���� � proceed ������� 0
                //    // ��������, ������ ������� ���� ����� ��������..
                //    if (AppendResponseIntTable(con, prevID, txtOtvet, dtDatOtv, nStatus, entt_id))
                //    {
                //        iRewriteState = 1;
                //        return true;
                //    }
                //}

                
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            
            // ��������� ����������� - � �� ����� ������� ��� �� �������
            if (con != null && con.State != ConnectionState.Closed) con.Close();

            return false;
        }

        private bool AppendResponseIntTable(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id)
        {

            OleDbCommand cmdInsMVV_RESPONSE;
            OleDbTransaction tran = null;

            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                // �������� MVV_RESPONSE

                cmdInsMVV_RESPONSE = new OleDbCommand();
                cmdInsMVV_RESPONSE.Connection = con;
                cmdInsMVV_RESPONSE.Transaction = tran;
                cmdInsMVV_RESPONSE.CommandText = "update EXT_RESPONSE SET RESPONSE_DATE = :RESPONSE_DATE, DATA_STR = DATA_STR || ' ' || '" + txtOtvet + "' WHERE ID = :ID";

                cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":RESPONSE_DATE", dtDatOtv));
                cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));


                if (cmdInsMVV_RESPONSE.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error appending row in EXT_RESPONSE table id = " + nID.ToString());
                    throw ex;
                }

                tran.Commit();
                con.Close();

                //SetDocumentStatus(nID, 19);// ���������� ������ ������� ����� ��� �������


                return true;
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }


        // �������� ����� �������� ������� ��� �������� DX_PACK_ID
        private bool InsertZaprosTo_PK_OSP(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id, ref int iRewriteState)
        {
            
            return InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, entt_id, ref iRewriteState, 0);
        }

        // ��� � iRewriteState
        // ���������� �������� ���������� ������ � ��������������, ����� MVV_EXTERNAL_RESPONSE � out iRewriteState
        // ����������� � �������
        private bool InsertZaprosTo_PK_OSP(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id, ref int iRewriteState, decimal nDX_PACK_ID)
        {

            OleDbCommand cmd, cmdMVV_I, cmdCheckAnsw, cmdInsDoc, cmdInsMVV_RESPONSE, cmdInsMVV_EXTERNAL_RESPONSE, cmdPackDocs;
            Decimal newID, prevID;
            OleDbTransaction tran = null;
            decimal nAgreementID = 0;
            decimal nAgent_dept_id = 0;
            decimal nAgent_id = 0;

            //iRewriteState = 
            //1 - ������� ����� - ����������� ������� � ������������ 
            //2 - ��������
            //3 - ������������
            //4 - ����������
            //20 - �������� ���
            //21 - ������������ ���
            //22 - ���������� ���, ������� �������

            try
            {

                // ���� �������� �� ����� ����� ��������, �� �������� ���������
                if (nDX_PACK_ID > 0)
                {
                    DataTable dtParams = GetPackParams(con, nID, entt_id);
                    if ((dtParams != null) && (dtParams.Rows.Count > 0))
                    {
                        nAgreementID = Convert.ToDecimal(dtParams.Rows[0]["agreement_id"]);
                        nAgent_dept_id = Convert.ToDecimal(dtParams.Rows[0]["agent_dept_id"]);
                        nAgent_id = Convert.ToDecimal(dtParams.Rows[0]["agent_id"]);
                    }

                }

                newID = 0;
                prevID = 0;
                
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                // ��������� ��� �� ���� ��������� ������ �� ���� ������
                // select id from MVV_EXTERNAL_RESPONSE ext left join MVV_RESPONSE resp on ext.id = resp.id left join DOCUMENT doc on ext.id = doc.id where doc.parent_id =:ID and resp.entity_id = :ENTITY_ID
                cmdCheckAnsw = new OleDbCommand("select first 1 ext.id from MVV_EXTERNAL_RESPONSE ext join MVV_RESPONSE resp on ext.id = resp.id join DOCUMENT doc on ext.id = doc.id where doc.parent_id =:ID and resp.entity_id = :ENTITY_ID", con, tran);
                cmdCheckAnsw.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));
                cmdCheckAnsw.Parameters.Add(new OleDbParameter(":ENTITY_ID", Convert.ToDecimal(entt_id)));
                prevID = Convert.ToDecimal(cmdCheckAnsw.ExecuteScalar());

                // ��� ���� ������������� iRewriteState

                if (prevID <= 0)
                {

                    // �������� ����� ����
                    cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
                    newID = Convert.ToDecimal(cmd.ExecuteScalar());

                    // �������� DOCUMENT
                    cmdInsDoc = new OleDbCommand();
                    cmdInsDoc.Connection = con;
                    cmdInsDoc.Transaction = tran;
                    cmdInsDoc.CommandText = "insert into DOCUMENT (ID, METAOBJECTNAME, DOCSTATUSID, DOCUMENTCLASSID,PARENT_ID, CREATE_DATE, SUSER_ID, DOC_DATE)";
                    cmdInsDoc.CommandText += " VALUES (:ID ,'MVV_EXTERNAL_RESPONSE', :DOCSTATUSID, :DOCUMENTCLASSID, :PARENT_ID, :CREATE_DATE, 8992, :DOC_DATE)";

                    cmdInsDoc.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                    cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(nStatus)));
                    cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCUMENTCLASSID", Convert.ToDecimal(301)));
                    cmdInsDoc.Parameters.Add(new OleDbParameter(":PARENT_ID", Convert.ToDecimal(nID)));
                    cmdInsDoc.Parameters.Add(new OleDbParameter(":CREATE_DATE", DateTime.Today));
                    cmdInsDoc.Parameters.Add(new OleDbParameter(":DOC_DATE", DateTime.Today));
                    

                    if (cmdInsDoc.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to document table parent_id = " + nID.ToString());
                        throw ex;
                    }

                    // TODO: �������� MVV_I c 14 ������

                    cmdMVV_I = new OleDbCommand();
                    cmdMVV_I.Connection = con;
                    cmdMVV_I.Transaction = tran;
                    cmdMVV_I.CommandText = "insert into MVV_I (ID)";
                    cmdMVV_I.CommandText += " VALUES (:ID)";

                    cmdMVV_I.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));

                    if (cmdMVV_I.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to MVV_I table parent_id = " + nID.ToString());
                        throw ex;
                    }


                    // �������� MVV_RESPONSE
                    // � 14 ������ ���������� ��� ��������� EXAD_AGENT_ID, EXAD_DEPT_ID, OUTER_AGREEMENT_ID
                    // � � 68-� ������ ��� ����-�� ������� � ���������� ��� ���

                    cmdInsMVV_RESPONSE = new OleDbCommand();
                    cmdInsMVV_RESPONSE.Connection = con;
                    cmdInsMVV_RESPONSE.Transaction = tran;
                    cmdInsMVV_RESPONSE.CommandText = "insert into MVV_RESPONSE (ID, RECEIVED, EXPORTED, QUERY_ID, ACT_DATA, DATA_STR, ENTITY_ID)"; //, EXAD_AGENT_ID, EXAD_DEPT_ID, OUTER_AGREEMENT_ID
                    cmdInsMVV_RESPONSE.CommandText += "  VALUES (:ID ,:RECEIVED, :EXPORTED, :QUERY_ID, :ACT_DATA, :DATA_STR, :ENTITY_ID)"; //, :EXAD_AGENT_ID, :EXAD_DEPT_ID, :OUTER_AGREEMENT_ID)";
                    
                    cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                    cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":RECEIVED", DateTime.Today));
                    cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":EXPORTED", dtDatOtv));
                    cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":QUERY_ID", Convert.ToDecimal(nID)));
                    cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":ACT_DATA", dtDatOtv));
                    cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":DATA_STR", txtOtvet));
                    cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":ENTITY_ID", entt_id));
                    
                    //cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":EXAD_AGENT_ID", nAgent_id));
                    //cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":EXAD_DEPT_ID", nAgent_dept_id));
                    //cmdInsMVV_RESPONSE.Parameters.Add(new OleDbParameter(":OUTER_AGREEMENT_ID", nAgreementID));

                    // TODO: OUTER_AGREEMENT_ID, OUTER_AGREEMENT_NAME - ����������

                    if (cmdInsMVV_RESPONSE.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to MVV_RESPOSNSE table id = " + nID.ToString());
                        throw ex;

                    }

                    
                    // �������� MVV_EXTERNAL_RESPONSE, �������� �������� DX_PACK_ID


                    cmdInsMVV_EXTERNAL_RESPONSE = new OleDbCommand();
                    cmdInsMVV_EXTERNAL_RESPONSE.Connection = con;
                    cmdInsMVV_EXTERNAL_RESPONSE.Transaction = tran;
                    cmdInsMVV_EXTERNAL_RESPONSE.CommandText = "insert into MVV_EXTERNAL_RESPONSE (ID, DX_PACK_ID)";
                    cmdInsMVV_EXTERNAL_RESPONSE.CommandText += "  VALUES (:ID, :DX_PACK_ID)";

                    cmdInsMVV_EXTERNAL_RESPONSE.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                    cmdInsMVV_EXTERNAL_RESPONSE.Parameters.Add(new OleDbParameter(":DX_PACK_ID", Convert.ToDecimal(nDX_PACK_ID)));

                    if (cmdInsMVV_EXTERNAL_RESPONSE.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to MVV_EXTERNAL_RESPONSE table id = " + nID.ToString());
                        throw ex;

                    }

                    // ���� ��������� � �����, �� ����� � ������������� �������
                    if (nDX_PACK_ID > 0)
                    {

                        cmdPackDocs = new OleDbCommand();
                        cmdPackDocs.Connection = con;
                        cmdPackDocs.Transaction = tran;
                        cmdPackDocs.CommandText = "insert into PACK_DOCS (DX_PACK_ID, DOCUMENT_ID)";
                        cmdPackDocs.CommandText += " VALUES (:PACK_ID ,:DOC_ID)";

                        cmdPackDocs.Parameters.Add(new OleDbParameter(":PACK_ID", Convert.ToDecimal(nDX_PACK_ID)));
                        cmdPackDocs.Parameters.Add(new OleDbParameter(":DOC_ID", Convert.ToDecimal(newID)));

                        if (cmdPackDocs.ExecuteNonQuery() == -1)
                        {
                            Exception ex = new Exception("Error inserting new row to pack_docs table pack_id = " + nDX_PACK_ID.ToString());
                            throw ex;
                        }
                    }

                    tran.Commit();
                    con.Close();

                    SetDocumentStatus(nID, 19);// ���������� ������ ������� ����� ��� �������
                    //SetDocumentStatus(nID, 15);// ���������� ������ ���������� � �������
                    
                    return true;
                }
                else
                {
                    tran.Rollback();
                    con.Close();

                    // ��� ���� ��� ��� ������, �� ��������� ������ frmRewriteDialog.ShowForm(), 
                    // � ���� iRewriteState > 1, �� �����������
                    if (iRewriteState == 1)
                    {
                        frmRewriteDialog frmRwD = new frmRewriteDialog();
                        iRewriteState = frmRwD.ShowForm();
                    }

                    // �������� �����������, ��������, ����������
                    switch (iRewriteState)
                    {
                        //2 - ��������
                        case (2):
                            if (AppendZaprosIn_PK_OSP(con, prevID, txtOtvet, dtDatOtv, nStatus, entt_id))
                            {
                                iRewriteState = 1;
                                return true;
                            }
                            
                            break;

                        //3 - ������������
                        case (3):
                            
                            if (UpdateZaprosIn_PK_OSP(con, prevID, txtOtvet, dtDatOtv, nStatus, entt_id))
                            {
                                iRewriteState = 1;
                                return true;
                            }
                            break;
                        
                        // 4 - ����������
                        case (4):
                            iRewriteState = 1;
                            break;

                        //20 - �������� ���
                        case (20):
                            if (AppendZaprosIn_PK_OSP(con, prevID, txtOtvet, dtDatOtv, nStatus, entt_id))
                            {
                                return true;
                            }
                            break;
                        //21 - ������������ ���
                        case (21):
                            if (UpdateZaprosIn_PK_OSP(con, prevID, txtOtvet, dtDatOtv, nStatus, entt_id))
                            {
                                return true;
                            }
                            break;
                            //22 - ���������� ���
                        default:
                            ;
                            break;
                            
                    }

                    return false;
                }
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }

        // ������� �������� ����� DX_PACK_I
        private decimal ID_CreateDX_PACK_I(OleDbConnection con, decimal nStatus, decimal nAGENT_ID, decimal nAGENT_DEPT_ID, decimal nAGREEMENT_ID, string txtPLAIN_LOG, string txtAgent_code, string txtAgreement_code, string txtAgent_dept_code)
        {
            // 1- �����
            // 70 - ���������
            // 71 - ��������� � ��������
            decimal nID = 0;
            OleDbCommand cmd, cmdInsDoc, cmdDX_PACK, cmdDX_PACK_I, cmdPACK_LOGS;
            OleDbTransaction tran = null;

            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                // TODO: �������� DX_PACK_I

                // �������� ����� ����
                cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
                nID = Convert.ToDecimal(cmd.ExecuteScalar());



                // �������� DOCUMENT
                cmdInsDoc = new OleDbCommand();
                cmdInsDoc.Connection = con;
                cmdInsDoc.Transaction = tran;
                cmdInsDoc.CommandText = "insert into DOCUMENT (ID, METAOBJECTNAME, DOCSTATUSID, DOCUMENTCLASSID,CREATE_DATE, SUSER_ID)";
                cmdInsDoc.CommandText += " VALUES (:ID ,'DX_PACK_I', :DOCSTATUSID, :DOCUMENTCLASSID, :CREATE_DATE, 8992)";

                cmdInsDoc.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(nStatus)));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCUMENTCLASSID", Convert.ToDecimal(303)));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":CREATE_DATE", DateTime.Now));

                if (cmdInsDoc.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to document table parent_id = " + nID.ToString());
                    throw ex;
                }


                cmdDX_PACK = new OleDbCommand();
                cmdDX_PACK.Connection = con;
                cmdDX_PACK.Transaction = tran;
                cmdDX_PACK.CommandText = "insert into DX_PACK (ID, AGENT_ID, AGENT_DEPT_ID, AGREEMENT_ID, AGENT, AGENT_DEPT, AGREEMENT, AGENT_CODE, AGREEMENT_CODE, AGENT_DEPT_CODE)";
                cmdDX_PACK.CommandText += " VALUES (:ID, :AGENT_ID, :AGENT_DEPT_ID, :AGREEMENT_ID, (select CAPTION from MVV_AGENT WHERE ID = :AGENT_ID2), (select EXAD_CAPTION from MVV_AGENT_DEPT WHERE EXAD_ID = :AGENT_DEPT_ID2), (select NAME_AGREEMENT from MVV_AGENT_AGREEMENT WHERE ID = :AGREEMENT_ID2),  :AGENT_CODE, :AGREEMENT_CODE, :AGENT_DEPT_CODE)";
                cmdDX_PACK.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));
                cmdDX_PACK.Parameters.Add(new OleDbParameter(":AGENT_ID", Convert.ToDecimal(nAGENT_ID)));
                cmdDX_PACK.Parameters.Add(new OleDbParameter(":AGENT_DEPT_ID", Convert.ToDecimal(nAGENT_DEPT_ID)));
                cmdDX_PACK.Parameters.Add(new OleDbParameter(":AGREEMENT_ID", Convert.ToDecimal(nAGREEMENT_ID)));
                cmdDX_PACK.Parameters.Add(new OleDbParameter(":AGENT_ID2", Convert.ToDecimal(nAGENT_ID)));
                cmdDX_PACK.Parameters.Add(new OleDbParameter(":AGENT_DEPT_ID2", Convert.ToDecimal(nAGENT_DEPT_ID)));
                cmdDX_PACK.Parameters.Add(new OleDbParameter(":AGREEMENT_ID2", Convert.ToDecimal(nAGREEMENT_ID)));
                // ������� � 163 ������ 15 ������
                cmdDX_PACK.Parameters.Add(new OleDbParameter(":AGENT_CODE", Convert.ToString(txtAgent_code)));
                cmdDX_PACK.Parameters.Add(new OleDbParameter(":AGREEMENT_CODE", Convert.ToString(txtAgreement_code)));
                cmdDX_PACK.Parameters.Add(new OleDbParameter(":AGENT_DEPT_CODE", Convert.ToString(txtAgent_dept_code)));

                if (cmdDX_PACK.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to DX_PACK table parent_id = " + nID.ToString());
                    throw ex;
                }

                cmdPACK_LOGS = new OleDbCommand();
                cmdPACK_LOGS.Connection = con;
                cmdPACK_LOGS.Transaction = tran;
                cmdPACK_LOGS.CommandText = "insert into PACK_LOGS (PACK_ID, LOG)";
                cmdPACK_LOGS.CommandText += " VALUES (:ID, :PLAIN_LOG)";

                cmdPACK_LOGS.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));
                cmdPACK_LOGS.Parameters.Add(new OleDbParameter(":PLAIN_LOG", Convert.ToString(txtPLAIN_LOG)));

                if (cmdPACK_LOGS.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to PACK_LOGS table pack_id = " + nID.ToString());
                    throw ex;
                }

                cmdDX_PACK_I = new OleDbCommand();
                cmdDX_PACK_I.Connection = con;
                cmdDX_PACK_I.Transaction = tran;
                cmdDX_PACK_I.CommandText = "insert into DX_PACK_I (ID)";
                cmdDX_PACK_I.CommandText += " VALUES (:ID)";

                cmdDX_PACK_I.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));

                if (cmdDX_PACK_I.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to DX_PACK_I table parent_id = " + nID.ToString());
                    throw ex;
                }

                tran.Commit();
                con.Close();

            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            
            return nID;
        }

        private decimal ID_InsertOtherIP_DocTo_PK_OSP(OleDbConnection con, decimal nStatus, decimal nUserID, DateTime dtIdate, decimal nIP_ID, DateTime dtExtDocDate, string txtExtDocNum, string txtContent, decimal nContrID)
        {

            OleDbCommand cmdIP, cmd, cmdInsDoc, cmdInsI, cmdInsI_IP, cmdInsI_IP_OTHER;
            Decimal newID, prevID;
            OleDbTransaction tran = null;
            DataSet dsIP_params;
            DataTable dtIP_params;
            decimal nIPNO_NUM = 0;
            string txtIP_DocNumber = "";
            decimal nSUSER_ID;
            string txtSUSER = "";
            string txtID_DEBTCLS_NAME = "";
            string txtContrName = "";
            string txtContrAdr = "";

            try
            {
                txtContrName = GetLegal_Name(mvd_id);
                txtContrAdr = GetLegal_Adr(mvd_id);

                dsIP_params = new DataSet();
                dtIP_params = dsIP_params.Tables.Add("IP_params");
                newID = 0;
                prevID = 0;
                
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                // ������ ��������� �� �� ��

                cmdIP = new OleDbCommand();
                cmdIP.Connection = con;
                cmdIP.Transaction = tran;
                cmdIP.CommandText = "select d.docstatusid, d.doc_number, d_ip.ip_exec_prist, d_ip.ip_exec_prist_name, d_ip_d.id_docdate, d_ip_d.id_debttext, d_ip.ipno_num  from doc_ip_doc d_ip_d left join doc_ip d_ip on d_ip_d.id = d_ip.id left join document d on d_ip_d.id = d.id where d_ip_d.id = :IP_ID";
                cmdIP.Parameters.Add(new OleDbParameter(":IP_ID", Convert.ToDecimal(nIP_ID)));
                using (OleDbDataReader rdr = cmdIP.ExecuteReader(CommandBehavior.Default))
                {
                    dsIP_params.Load(rdr, LoadOption.OverwriteChanges, dtIP_params);
                    rdr.Close();
                }

                if ((dsIP_params != null) && (dsIP_params.Tables.Count > 0))
                {
                    txtIP_DocNumber = Convert.ToString(dsIP_params.Tables[0].Rows[0]["doc_number"]);
                    nIPNO_NUM = Convert.ToDecimal(dsIP_params.Tables[0].Rows[0]["ipno_num"]);
                    nSUSER_ID = Convert.ToDecimal(dsIP_params.Tables[0].Rows[0]["ip_exec_prist"]);
                    txtSUSER = Convert.ToString(dsIP_params.Tables[0].Rows[0]["ip_exec_prist_name"]);
                    txtID_DEBTCLS_NAME = Convert.ToString(dsIP_params.Tables[0].Rows[0]["id_debttext"]);
                    if (nSUSER_ID > 0)
                    {
                        nUserID = nSUSER_ID;
                    }
                }
                else
                {
                    return -1;
                }

                // �������� ����� ����
                cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
                newID = Convert.ToDecimal(cmd.ExecuteScalar());

                // �������� DOCUMENT
                cmdInsDoc = new OleDbCommand();
                cmdInsDoc.Connection = con;
                cmdInsDoc.Transaction = tran;
                cmdInsDoc.CommandText = "insert into DOCUMENT (ID, METAOBJECTNAME, DOCSTATUSID, DOCUMENTCLASSID, CREATE_DATE, SUSER_ID)";
                cmdInsDoc.CommandText += " VALUES (:ID, 'I_IP_OTHER', :DOCSTATUSID, :DOCUMENTCLASSID, :CREATE_DATE, :SUSER_ID)";

                cmdInsDoc.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));

                //cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(1)));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(nStatus)));

                cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCUMENTCLASSID", Convert.ToDecimal(11))); // ����� ���������������� ��� ������� I - �������� ��������
                //cmdInsDoc.Parameters.Add(new OleDbParameter(":PARENT_ID", Convert.ToDecimal(nID)));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":CREATE_DATE", DateTime.Now));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":SUSER_ID", Convert.ToDecimal(nUserID)));
                //cmdInsDoc.Parameters.Add(new OleDbParameter(":AMOUNT", Convert.ToDouble(nAmount)));


                if (cmdInsDoc.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to document table parent_id = " + newID.ToString());
                    throw ex;
                }

                // �������� I

                cmdInsI = new OleDbCommand();
                cmdInsI.Connection = con;
                cmdInsI.Transaction = tran;
                cmdInsI.CommandText = "insert into I (ID, PAGECNT, APPNCNT, SECURTYPE, APPBNPAGECNT, I_IDATE, EXTDOCDATE, EXTDOCNO, CONTR, CONTR_NAME, ADR)";
                cmdInsI.CommandText += "  VALUES (:ID, 1, 0, 2, 0, :I_IDATE, :EXTDOCDATE, :EXTDOCNO, :CONTR, :CONTR_NAME, :ADR)";
                cmdInsI.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                //cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", dtIdate));
                cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", DateTime.Today));
                cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCDATE", dtExtDocDate));
                cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCNO", txtExtDocNum));
                cmdInsI.Parameters.Add(new OleDbParameter(":CONTR", nContrID));
                cmdInsI.Parameters.Add(new OleDbParameter(":CONTR_NAME", txtContrName));
                cmdInsI.Parameters.Add(new OleDbParameter(":ADR", txtContrAdr));


                if (cmdInsI.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to I table id = " + newID.ToString());
                    throw ex;

                }


                // �������� I_IP


                cmdInsI_IP = new OleDbCommand();
                cmdInsI_IP.Connection = con;
                cmdInsI_IP.Transaction = tran;
                cmdInsI_IP.CommandText = "insert into I_IP (ID, IP_DOC_NUMBER,IP_ID, IPNO_NUM, ID_DEBTCLS_NAME, IP_EXEC_PRIST, IP_EXEC_PRIST_NAME)";
                cmdInsI_IP.CommandText += "  VALUES (:ID, :IP_DOC_NUMBER, :IP_ID, :IPNO_NUM, :ID_DEBTCLS_NAME, :IP_EXEC_PRIST, :IP_EXEC_PRIST_NAME)";

                cmdInsI_IP.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_DOC_NUMBER", Convert.ToString(txtIP_DocNumber)));
                cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_ID", Convert.ToDecimal(nIP_ID)));
                cmdInsI_IP.Parameters.Add(new OleDbParameter(":IPNO_NUM", Convert.ToDecimal(nIPNO_NUM)));
                cmdInsI_IP.Parameters.Add(new OleDbParameter(":ID_DEBTCLS_NAME", txtID_DEBTCLS_NAME));
                cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_EXEC_PRIST", Convert.ToDecimal(nSUSER_ID)));
                cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_EXEC_PRIST_NAME", Convert.ToString(txtSUSER)));

                if (cmdInsI_IP.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to I_IP table id = " + newID.ToString());
                    throw ex;
                }


                cmdInsI_IP_OTHER = new OleDbCommand();
                cmdInsI_IP_OTHER.Connection = con;
                cmdInsI_IP_OTHER.Transaction = tran;
                cmdInsI_IP_OTHER.CommandText = "insert into I_IP_OTHER (ID, INDOC_TYPE, INDOC_TYPE_NAME, I_IP_OTHER_CONTENT)";
                cmdInsI_IP_OTHER.CommandText += "  VALUES (:ID, :INDOC_TYPE, :INDOC_TYPE_NAME, :I_IP_OTHER_CONTENT)";
                cmdInsI_IP_OTHER.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                cmdInsI_IP_OTHER.Parameters.Add(new OleDbParameter(":INDOC_TYPE", Convert.ToInt32(37)));
                cmdInsI_IP_OTHER.Parameters.Add(new OleDbParameter(":INDOC_TYPE_NAME", Convert.ToString("���������������� ������")));
                cmdInsI_IP_OTHER.Parameters.Add(new OleDbParameter(":I_IP_OTHER_CONTENT", txtContent));
                
                if (cmdInsI_IP_OTHER.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to I_IP_OTHER  table id = " + newID.ToString());
                    throw ex;

                }

                tran.Commit();
                con.Close();

                return newID;

            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return -1;
        }
        
        // �������� ���������� �������� � ��������� ����� 
        private decimal ID_InsertPlatDocTo_PK_OSP(OleDbConnection con, decimal nStatus, decimal nUserID, double nAmount, DateTime dtIdate, decimal nIP_ID, DateTime dtExtDocDate, string txtExtDocNum, decimal nContrID, string txtFIO_D)
        {

            OleDbCommand cmdIP, cmd, cmdInsDoc, cmdInsI, cmdInsI_IP, cmdInsI_DEPOSIT, cmdInsI_OP_CS, cmdInsI_OP_CS_ENDDBT;
            Decimal newID, prevID;
            OleDbTransaction tran = null;
            DataSet dsIP_params;
            DataTable dtIP_params;
            decimal nIPNO_NUM = 0;
            string txtIP_DocNumber = "";
            decimal nSUSER_ID;
            string txtSUSER = "";
            string txtID_DEBTCLS_NAME = "";
            string txtContrName = "";
            string txtContrAdr = "";
            decimal id_dbtr = 0;
            
            try
            {
                // ����� ������ ������� - ��� ������� contr_id �� i_id � �������� ����

                dsIP_params = new DataSet();
                dtIP_params = dsIP_params.Tables.Add("IP_params");
                newID = 0;
                prevID = 0;
                id_dbtr = 0;

                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();

                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                // ������ ��������� �� �� ��

                cmdIP = new OleDbCommand();
                cmdIP.Connection = con;
                cmdIP.Transaction = tran;
                cmdIP.CommandText = "select d_ip_d.id_dbtr_name, d_ip_d.id_dbtr_adr, d.docstatusid, d.doc_number, d_ip.id_dbtr, d_ip.ip_exec_prist, d_ip.ip_exec_prist_name, d_ip_d.id_docdate, d_ip_d.id_debttext, d_ip.ipno_num  from doc_ip_doc d_ip_d left join doc_ip d_ip on d_ip_d.id = d_ip.id left join document d on d_ip_d.id = d.id where d_ip_d.id = :IP_ID";
                cmdIP.Parameters.Add(new OleDbParameter(":IP_ID", Convert.ToDecimal(nIP_ID)));
                using (OleDbDataReader rdr = cmdIP.ExecuteReader(CommandBehavior.Default))
                {
                    dsIP_params.Load(rdr, LoadOption.OverwriteChanges, dtIP_params);
                    rdr.Close();
                }

                if ((dsIP_params != null) && (dsIP_params.Tables.Count > 0))
                {
                    txtIP_DocNumber = Convert.ToString(dsIP_params.Tables[0].Rows[0]["doc_number"]);
                    nIPNO_NUM = Convert.ToDecimal(dsIP_params.Tables[0].Rows[0]["ipno_num"]);
                    nSUSER_ID = Convert.ToDecimal(dsIP_params.Tables[0].Rows[0]["ip_exec_prist"]);
                    txtSUSER = Convert.ToString(dsIP_params.Tables[0].Rows[0]["ip_exec_prist_name"]);
                    txtID_DEBTCLS_NAME = Convert.ToString(dsIP_params.Tables[0].Rows[0]["id_debttext"]);
                    txtID_DEBTCLS_NAME = Convert.ToString(dsIP_params.Tables[0].Rows[0]["id_debttext"]);
                    txtContrName = Convert.ToString(dsIP_params.Tables[0].Rows[0]["id_dbtr_name"]); // ������ ������������ ����� ��� �������
                    txtContrAdr = Convert.ToString(dsIP_params.Tables[0].Rows[0]["id_dbtr_adr"]);

                    if (nSUSER_ID > 0)
                    {
                        nUserID = nSUSER_ID;
                    }
                }
                else
                {
                    return -1;
                }

                // �������� ����� ����
                cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
                newID = Convert.ToDecimal(cmd.ExecuteScalar());

                // �������� DOCUMENT
                cmdInsDoc = new OleDbCommand();
                cmdInsDoc.Connection = con;
                cmdInsDoc.Transaction = tran;
                cmdInsDoc.CommandText = "insert into DOCUMENT (ID, METAOBJECTNAME, DOCSTATUSID, DOCUMENTCLASSID, CREATE_DATE, SUSER_ID, AMOUNT)";
                cmdInsDoc.CommandText += " VALUES (:ID, 'I_OP_CS_ENDDBT', :DOCSTATUSID, :DOCUMENTCLASSID, :CREATE_DATE, :SUSER_ID, :AMOUNT)";
                
                cmdInsDoc.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                
                //cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(1)));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(nStatus)));

                cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCUMENTCLASSID", Convert.ToDecimal(78)));
                //cmdInsDoc.Parameters.Add(new OleDbParameter(":PARENT_ID", Convert.ToDecimal(nID)));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":CREATE_DATE", DateTime.Now));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":SUSER_ID", Convert.ToDecimal(nUserID)));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":AMOUNT", Convert.ToDouble(nAmount)));
                

                if (cmdInsDoc.ExecuteNonQuery() == -1)
                {
                        Exception ex = new Exception("Error inserting new row to document table parent_id = " + newID.ToString());
                        throw ex;
                }

                // �������� I
                // - ����������� 	I.CONTR_NAME
                // - ����� ����������� I.ADR
                

                    cmdInsI = new OleDbCommand();
                    cmdInsI.Connection = con;
                    cmdInsI.Transaction = tran;
                    cmdInsI.CommandText = "insert into I (ID, PAGECNT, APPNCNT, SECURTYPE, APPBNPAGECNT, I_IDATE, EXTDOCDATE, EXTDOCNO, CONTR, CONTR_NAME, ADR)";
                    cmdInsI.CommandText += "  VALUES (:ID, 1, 0, 2, 0, :I_IDATE, :EXTDOCDATE, :EXTDOCNO, :CONTR, :CONTR_NAME, :ADR)";
                    cmdInsI.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                    //cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", dtIdate));
                    cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", DateTime.Today));
                    cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCDATE", dtExtDocDate));
                    cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCNO", txtExtDocNum));
                    cmdInsI.Parameters.Add(new OleDbParameter(":CONTR", nContrID));
                    cmdInsI.Parameters.Add(new OleDbParameter(":CONTR_NAME", txtContrName));
                    cmdInsI.Parameters.Add(new OleDbParameter(":ADR", txtContrAdr));
                
                    if (cmdInsI.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to I table id = " + newID.ToString());
                        throw ex;

                    }


                    // �������� I_IP


                    cmdInsI_IP = new OleDbCommand();
                    cmdInsI_IP.Connection = con;
                    cmdInsI_IP.Transaction = tran;
                    cmdInsI_IP.CommandText = "insert into I_IP (ID, IP_DOC_NUMBER,IP_ID, IPNO_NUM, ID_DEBTCLS_NAME, IP_EXEC_PRIST, IP_EXEC_PRIST_NAME)";
                    cmdInsI_IP.CommandText += "  VALUES (:ID, :IP_DOC_NUMBER, :IP_ID, :IPNO_NUM, :ID_DEBTCLS_NAME, :IP_EXEC_PRIST, :IP_EXEC_PRIST_NAME)";

                    cmdInsI_IP.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                    cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_DOC_NUMBER", Convert.ToString(txtIP_DocNumber)));
                    cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_ID", Convert.ToDecimal(nIP_ID)));
                    cmdInsI_IP.Parameters.Add(new OleDbParameter(":IPNO_NUM", Convert.ToDecimal(nIPNO_NUM)));
                    cmdInsI_IP.Parameters.Add(new OleDbParameter(":ID_DEBTCLS_NAME", txtID_DEBTCLS_NAME));
                    cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_EXEC_PRIST", Convert.ToDecimal(nSUSER_ID)));
                    cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_EXEC_PRIST_NAME", Convert.ToString(txtSUSER)));

                    if (cmdInsI_IP.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to I_IP table id = " + newID.ToString());
                        throw ex;
                    }


                    // �������� I_DEPOSIT 

                    cmdInsI_DEPOSIT  = new OleDbCommand();
                    cmdInsI_DEPOSIT.Connection = con;
                    cmdInsI_DEPOSIT.Transaction = tran;
                    cmdInsI_DEPOSIT.CommandText = "insert into I_DEPOSIT (ID)";
                    cmdInsI_DEPOSIT.CommandText += "  VALUES (:ID)";
                    cmdInsI_DEPOSIT.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));

                    if (cmdInsI_DEPOSIT.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to I_DEPOSIT  table id = " + newID.ToString());
                        throw ex;

                    }

                    // �������� I_OP_CS

                    cmdInsI_OP_CS  = new OleDbCommand();
                    cmdInsI_OP_CS.Connection = con;
                    cmdInsI_OP_CS.Transaction = tran;
                    cmdInsI_OP_CS.CommandText = "insert into I_OP_CS (ID, CHANGEDBT_REASON_ID, CHANGEDBT_REASON_DESCR, I_OP_CS_CHANGESUM)";
                    cmdInsI_OP_CS.CommandText += "  VALUES (:ID, 3, '������ ������ � �����', :I_OP_CS_CHANGESUM)";
                    cmdInsI_OP_CS.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                    cmdInsI_OP_CS.Parameters.Add(new OleDbParameter(":I_OP_CS_CHANGESUM", Convert.ToDouble(nAmount)));

                    if (cmdInsI_OP_CS.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to I_OP_CS  table id = " + newID.ToString());
                        throw ex;

                    }


                    // �������� I_OP_CS_ENDDBT

                    cmdInsI_OP_CS_ENDDBT  = new OleDbCommand();
                    cmdInsI_OP_CS_ENDDBT.Connection = con;
                    cmdInsI_OP_CS_ENDDBT.Transaction = tran;
                    cmdInsI_OP_CS_ENDDBT.CommandText = "insert into I_OP_CS_ENDDBT (ID)";
                    cmdInsI_OP_CS_ENDDBT.CommandText += "  VALUES (:ID)";
                    cmdInsI_OP_CS_ENDDBT.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));

                    if (cmdInsI_OP_CS_ENDDBT.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to I_OP_CS_ENDDBT  table id = " + newID.ToString());
                        throw ex;

                    }

                    tran.Commit();
                    con.Close();

                    return newID;
                
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return -1;
        }


        private decimal ID_Find_GIBDD_IP(DateTime dtDateID, string txtNomID)
        {
            decimal nID = -1;

            //� ������� ����������� �������, ���������� ���� � ����� �� ����� �� �� ������� � ����������

            // � ����� ������� ����� ������? join doc_ip_doc, doc_ip, document
            // where (d.docstatusid ! = -1) and (d.docstatusid ! = 7) and (d.docstatusid ! = 10) = !(������, �������, �������)
            // select d_ip_d.id from doc_ip_doc d_ip_d left join doc_ip d_ip on d_ip_d.id = d_ip.id left join document d on d_ip_d.id = d.id where (d.docstatusid != -1) and (d.docstatusid != 7) and (d.docstatusid != 10)  and d_ip_d.id_docno = '006971' and d_ip_d.id_docdate = '22.09.2010'
            try
            {
                
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand("select first 1 d_ip_d.id from doc_ip_doc d_ip_d left join doc_ip d_ip on d_ip_d.id = d_ip.id left join document d on d_ip_d.id = d.id where (d.docstatusid != -1) and (d.docstatusid != 7) and (d.docstatusid != 10)  and d_ip_d.id_docno = :DOC_NUM and d_ip_d.id_docdate = :DOC_DATE", con, tran);

                cmd.Parameters.Add(new OleDbParameter(":DOC_NUM", Convert.ToDateTime(txtNomID)));
                cmd.Parameters.Add(new OleDbParameter(":DOC_DATE", Convert.ToDateTime(dtDateID)));
                
                nID = Convert.ToDecimal(cmd.ExecuteScalar());
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
            
            return nID;
        }
        
        
        private void btnAnswer_Click(object sender, EventArgs e)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);
            decimal nParentID = 0;

            Decimal org;
            Decimal nStatus;
            int iRewriteState = 1; // ������� ����� ���������� ������� �� ������ (����������� �������� � ������������)
            decimal nAgreementID = 0;
            decimal nAgent_dept_id = 0;
            decimal nAgent_id = 0;
            decimal nDx_pack_id = 0;
            decimal nNewPackID = 0;

            string txtAgreementCode = "";
            string txtAgentCode = "";
            string txtAgentDeptCode = "";
            string txtEntityName = "";
            bool bNotIntTablesResp = false; // ���� ����� �� ������, ��������� ��� ������������ ������.

            if (cbxOrg.SelectedValue != null)
            {
                // ����� � ���� �� ��� �������� ������������, �� ID ������ � ������
                org = Convert.ToDecimal(cbxOrg.SelectedValue);

                openFileDialog1.Filter = "DBF �����(*.dbf)|*.dbf";
                DialogResult res = openFileDialog1.ShowDialog();
                if (res == DialogResult.OK)
                {
                    if (openFileDialog1.FileName != "")
                    {
                        ChangeByte(openFileDialog1.FileName, 0x65, 30);
                        nStatus = 0;
                        # region "NOFIND"
                        if (openFileDialog1.FileName.ToLower().Contains("nofind.dbf"))
                        {
                            try
                            {
                                DataSet ds = new DataSet();
                                DataTable tbl = ds.Tables.Add("NOFIND");
                                DBFcon = new OleDbConnection();
                                DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=RUSSIAN");
                                DBFcon.Open();
                                m_cmd = new OleDbCommand();
                                m_cmd.Connection = DBFcon;
                                m_cmd.CommandText = "SELECT * FROM NOFIND ORDER BY ZAPROS";// ����������� �� ���� ZAPROS ����� �������� � ����������� �������� �� 1 ������
                                using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                                {
                                    ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                                    rdr.Close();
                                }

                                DBFcon.Close();

                                Int32 iCnt = 0;
                                //OleDbTransaction tran;

                                prbWritingDBF.Value = 0;
                                prbWritingDBF.Maximum = tbl.Rows.Count;
                                prbWritingDBF.Step = 1;

                                //con.Open();
                                //tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                                DateTime dat1 = DateTime.Today;
                                DateTime dat2 = DateTime.Today;

                                if (tbl.Rows.Count > 0)
                                {
                                    if (!(DateTime.TryParse(Convert.ToString(tbl.Rows[0]["DATZPR1"]), out dat1)))
                                    {
                                        dat1 = DateTime.Today;
                                    }

                                    if (!(DateTime.TryParse(Convert.ToString(tbl.Rows[0]["DATZPR2"]), out dat2)))
                                    {
                                        dat2 = DateTime.Today;
                                    }
                                }

                                Decimal nID = 0;
                                String txtID = "";

                                nAgreementID = 0;
                                nAgent_dept_id = 0;
                                nAgent_id = 0;
                                nDx_pack_id = 0;
                                nNewPackID = 0;

                                txtAgreementCode = "";
                                txtAgentCode = "";
                                txtAgentDeptCode = "";
                                txtEntityName = "";


                                // ���� ���� � �������� �� ������, �� �� ������� ������ ����������
                                // ��b ������ �� ������� �� 14 ������ ��� ������ (��� �������)
                                if (tbl.Rows.Count > 0)
                                {
                                    decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["ZAPROS"]);

                                    if (FindSendlist(nFirstID, org)) // ��������� �������� org - ���������� �� ������ ��������, �������� ���� ���������� �����
                                    {
                                        // ������ ��� ����� ������
                                        // �������� ���������: ����������, ����������, �������������
                                        DataTable dtParams = GetPackParams(con, nFirstID, org);
                                        if ((dtParams != null) && (dtParams.Rows.Count > 0))
                                        {
                                            if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agreement_id"]), out nAgreementID))
                                            {
                                                nAgreementID = 0;
                                            }

                                            if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agent_dept_id"]), out nAgent_dept_id))
                                            {
                                                nAgent_dept_id = 0;
                                            }

                                            if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agent_id"]), out nAgent_id))
                                            {
                                                nAgent_id = 0;
                                            }

                                            if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["dx_pack_id"]), out nDx_pack_id))
                                            {
                                                nDx_pack_id = 0;
                                            }

                                            //nAgreementID = Convert.ToDecimal(dtParams.Rows[0]["agreement_id"]);
                                            //nAgent_dept_id = Convert.ToDecimal(dtParams.Rows[0]["agent_dept_id"]);
                                            //nAgent_id = Convert.ToDecimal(dtParams.Rows[0]["agent_id"]);
                                            //nDx_pack_id = Convert.ToDecimal(dtParams.Rows[0]["dx_pack_id"]);
                                        }
                                    }

                                    if (nAgreementID == 0)
                                    {
                                        nAgreementID = GetAgr_by_Org(org); // ����� ����������
                                        nAgent_id = GetAgent_ID(nAgreementID);
                                        nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                    }

                                    txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                    txtEntityName = GetLegal_Name(org);

                                    // ����� ������� ����� �������� �����
                                    //nNewPackID = ID_CreateDX_PACK_I(con, 1, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);

                                    // TODO: �������� ����� ������ �������, � �������� ������ �����
                                    frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD, "FL_NOFIND");
                                    nParentID = LogList.ShowForm();

                                    // ���� �� ���� ������� ���������� �������� ������
                                    if (nParentID != -1)
                                    {
                                        // 1 - �����
                                        // 4 - ����� �������������
                                        nNewPackID = CreateLLog(conGIBDD, 1, 4, txtAgreementCode, nParentID, "����� ������� �� " + txtEntityName + ".");



                                        // �������� � ��� ������ ���� � ������ ���������
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                        //WritePackLog(con, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");


                                        foreach (DataRow row in tbl.Rows)
                                        {
                                            //m_cmd = new OleDbCommand();
                                            //m_cmd.Connection = con;
                                            //m_cmd.Transaction = tran;
                                            // ��� ���-�� ���� ������ �������� ������� �������� ������ ������
                                            // ���������� �� ���� (ZAPROS)  ����� �������� � ���� DOCUMENT.ID
                                            nStatus = 7; // ��� ������
                                            txtID = Convert.ToString(row["ZAPROS"]);
                                            if (!Decimal.TryParse(txtID, out nID))
                                            {
                                                nID = 0;
                                            }
                                            if (FindZapros(nID))
                                            {
                                                // ������� �������� ��������� � ���� ��������� ������ ������
                                                try
                                                {
                                                    string txtDatOtv = "";
                                                    DateTime dtDatOtv;

                                                    txtDatOtv = Convert.ToString(row["DATOTV"]);
                                                    if (!DateTime.TryParse(txtDatOtv, out dtDatOtv))
                                                    {
                                                        dtDatOtv = DateTime.MaxValue;
                                                    }

                                                    string txtDatZap = "";
                                                    DateTime dtDatZap;

                                                    // ��������� ���� �������
                                                    txtDatZap = Convert.ToString(row["DATZPR2"]);
                                                    if (!DateTime.TryParse(txtDatZap, out dtDatZap))
                                                    {
                                                        dtDatZap = DateTime.MaxValue;
                                                    }

                                                    bNotIntTablesResp = false; // ������ ��� ������ ����� �� ������������ ������
                                                    //if (dtDatZap < dtIntTablesDeplmntDate)
                                                    //{
                                                    //    bNotIntTablesResp = true;
                                                    //}
                                                    //else
                                                    //{
                                                    //    bNotIntTablesResp = false;
                                                    //}


                                                    string txtOtvet;

                                                    // ��� ���� � �������� ��� � ��� ��������, �����
                                                    string txtResLine = Convert.ToString(row["FIO"]).TrimEnd();
                                                    if (row["GODR"] != System.DBNull.Value)
                                                    {
                                                        txtResLine += " (" + Convert.ToInt32(row["GODR"]).ToString() + " �.�.)";
                                                    }
                                                    txtResLine += " " + Convert.ToString(row["ADRES"]).TrimEnd();

                                                    txtOtvet = "� ������������ � " + PKOSP_GetOrgConvention(org);
                                                    txtOtvet += " ������� �����: ";

                                                    txtOtvet += "����� �� " + GetLegal_Name(org) + ". ��� ������ � �������� " + txtResLine + ". ���� ������: " + dtDatOtv.ToShortDateString();

                                                    // iRewriteState
                                                    // 1 - ������� ����� - ����������� ������� � ������������ 
                                                    // 2 - �������� ���
                                                    // 3 - ������������ ���4 - ���������� ���, ������� �������
                                                    if (bNotIntTablesResp)
                                                    {
                                                        if (InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID))
                                                        {
                                                            iCnt++;
                                                            WritePackLog(con, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");
                                                            prbWritingDBF.PerformStep();
                                                            prbWritingDBF.Refresh();
                                                            System.Windows.Forms.Application.DoEvents();
                                                        }
                                                        else
                                                        {
                                                            // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                            WritePackLog(con, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                            nStatus = 15; // ������
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                        {
                                                            iCnt++;
                                                            // WritePackLog(con, nNewPackID, "��������� ����� # " + nID.ToString() + "\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");

                                                            prbWritingDBF.PerformStep();
                                                            prbWritingDBF.Refresh();
                                                            System.Windows.Forms.Application.DoEvents();
                                                        }
                                                        else
                                                        {
                                                            // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                            //WritePackLog(con, nNewPackID, "������! ����� # " + nID.ToString() + " ���������� �� �������.\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                            nStatus = 15; // ������
                                                        }
                                                    }

                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                                                    if (nNewPackID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "������! �������� ������ ������� ��������� ����������.\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "�������� �������� = " + iCnt.ToString() + "\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");

                                                        if (nID > 0)
                                                        {
                                                            WriteLLog(conGIBDD, nNewPackID, "ID ������� = " + nID.ToString() + "\n");
                                                        }
                                                    }

                                                }
                                            }
                                            else
                                            {
                                                // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                if (nNewPackID > 0)
                                                {
                                                    WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ ID = " + nID.ToString() + " �� ������� ��������� �.�. �� ��������� ������-��������.\n");
                                                }
                                            }
                                        }
                                        //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                        //WritePackLog(con, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                        // ���������� ���������� ������������ ��������
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                        // �������� ������ ����-������
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                        // �������� ������ ����-��������
                                        // ����� ������ ������, �.�. ������ ��� ����� �� ������
                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - �������� �����
                                        // �������� ���� ��� ��������� NOFIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_NOFIND");


                                        // ���������������� �.�. ������ �������� ������ �� ���������
                                        //// ���� ��� ��, �� ����� �������� ������ ������
                                        //if (nNewPackID > 0)
                                        //{
                                        //    SetDocumentStatus(nNewPackID, 70);
                                        //}

                                    }
                                }
                                else
                                {
                                    if (nAgreementID == 0)
                                    {
                                        nAgreementID = GetAgr_by_Org(org); // ����� ����������
                                        nAgent_id = GetAgent_ID(nAgreementID);
                                        nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                    }


                                    txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                    txtEntityName = GetLegal_Name(org);

                                    // ������� ��� ������������� ������� � �������� ���� ��� 0 � ������ �������
                                    frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD, "FL_NOFIND");
                                    nParentID = LogList.ShowForm();

                                    // ���� �� ���� ������� ���������� �������� ������
                                    if (nParentID != -1)
                                    {
                                        // 1 - �����
                                        // 4 - ����� �������������
                                        nNewPackID = CreateLLog(conGIBDD, 1, 4, txtAgreementCode, nParentID, "����� ������� �� " + txtEntityName + ".");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");
                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                        // ���������� ���������� ������������ ��������
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                        // �������� ������ ����-������
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                        // �������� ������ ����-��������
                                        // ����� ������ ������, �.�. ������ ��� ����� �� ������
                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - �������� �����
                                        // �������� ���� ��� ��������� NOFIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_NOFIND");
                                    }
                                }
   
                                    
                                MessageBox.Show("���������� �������: " + iCnt.ToString() + ".\n������ ����� ����������� ������ �������.", "���������", MessageBoxButtons.OK);
#region "REESTR"
                                //**********������������**�������**nofind************
                                //��� ���������� �� ���������. 

                                //������ ���� ���������
                                DataTable dtspi = ds.Tables.Add("SPI");

                                DBFcon.Open();
                                m_cmd = new OleDbCommand();
                                m_cmd.Connection = DBFcon;
                                m_cmd.CommandText = "SELECT DISTINCT NOMSPI FROM NOFIND";

                                using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                                {
                                    ds.Load(rdr, LoadOption.OverwriteChanges, dtspi);
                                    rdr.Close();
                                }

                                DBFcon.Close();

                                #region "HTML print"
                                // ���� ���������� ��� HTML
                                prbWritingDBF.Value = 0;
                                prbWritingDBF.Maximum = dtspi.Rows.Count;
                                prbWritingDBF.Step = 1;
                                Int32 spi = 0;

                                ReportMaker report = new ReportMaker();
                                report.StartReport();
                                foreach (DataRow drspi in dtspi.Rows)
                                {
                                    bool fl_no_answer = true;
                                    report.AddToReport("<h3>");
                                    report.AddToReport("������ ������� �� ������� ��-� � ������� ���. ������� �� ��������� �����������");
                                    report.AddToReport(GetLegal_Name(org) + " �� " + Convert.ToDateTime(tbl.Rows[0]["DATOTV"]).ToShortDateString() + "<br>");
                                    //report.AddToReport("�� ������ � " + Convert.ToDateTime(tbl.Rows[0]["DATZPR1"]).ToShortDateString() + " �� " + Convert.ToDateTime(tbl.Rows[0]["DATZPR2"]).ToShortDateString() + "<br>");
                                    report.AddToReport("��� ������ � ������� ������ � ���������<br>");

                                    spi = Convert.ToInt32(drspi["NOMSPI"]);
                            
                                    report.AddToReport("��-�: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "<br>");
                                    report.AddToReport("</h3>");

                                    foreach (DataRow row in tbl.Rows)
                                    {
                                        report.AddToReport("<p>");
                                        if (spi == Convert.ToInt32(row["NOMSPI"]))
                                        {
                                            report.AddToReport(GetIPNum(Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIO"]).TrimEnd() + "");
                                      
                                        }
                                        report.AddToReport("</p>");
                                    }


                                    report.SplitNewPage();
                                    prbWritingDBF.PerformStep();
                                    prbWritingDBF.Refresh();
                                    System.Windows.Forms.Application.DoEvents();

                                }
                                report.EndReport();
                                report.ShowReport();

                                #endregion

                                #region "OLD PRINT"
                                //if (OooIsInstall)
                                //{
                                //    //OOo start
                                //    OOo_Writer OOo_cld = new OOo_Writer();
                                //    OOo_cld.OOo_Cred(GetLegal_Name(org), "nofind", ds, con, this);
                                //}
                                //else
                                //{
                                //    //      ������ ��� �����
                                //    Microsoft.Office.Interop.Word.ApplicationClass app = new ApplicationClass();

                                //    object s1 = "";
                                //    object fl = false;
                                //    object t = WdNewDocumentType.wdNewBlankDocument;
                                //    object fl2 = true;

                                //    Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref s1, ref fl, ref t, ref fl2);

                                //    Paragraph par;

                                //    int spi;
                                //    int sch_line;
                                //    int fl_fst = 1;

                                //    prbWritingDBF.Value = 0;
                                //    prbWritingDBF.Maximum = dtspi.Rows.Count;
                                //    prbWritingDBF.Step = 1;

                                //    foreach (DataRow drspi in dtspi.Rows)
                                //    {
                                //        sch_line = 0;
                                //        if (fl_fst == 1)
                                //        {
                                //            sch_line = 1;
                                //            fl_fst = 0;
                                //            par = doc.Paragraphs[1];
                                //        }
                                //        else
                                //        {
                                //            object oMissing = System.Reflection.Missing.Value;
                                //            par = doc.Paragraphs.Add(ref oMissing);
                                //            object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                                //            par.Range.InsertBreak(ref oPageBreak);
                                //        }

                                //        par.Range.Font.Name = "Courier";
                                //        par.Range.Font.Size = 8;
                                //        float a = par.Range.PageSetup.RightMargin;
                                //        float b = par.Range.PageSetup.LeftMargin;
                                //        float c = par.Range.PageSetup.TopMargin;
                                //        float d = par.Range.PageSetup.BottomMargin;

                                //        par.Range.PageSetup.RightMargin = 30;
                                //        par.Range.PageSetup.LeftMargin = 30;
                                //        par.Range.PageSetup.TopMargin = 20;
                                //        par.Range.PageSetup.BottomMargin = 20;

                                //        par.Range.Text += "������ ������� �� ������� ��-� � ������� ���. ������� �� �����";
                                //        par.Range.Text += GetLegal_Name(org) + " �� " + Convert.ToDateTime(tbl.Rows[0]["DATOTV"]).ToShortDateString() + "\n";
                                //        par.Range.Text += "�� ������ � " + Convert.ToDateTime(tbl.Rows[0]["DATZPR1"]).ToShortDateString() + " �� " + Convert.ToDateTime(tbl.Rows[0]["DATZPR2"]).ToShortDateString() + "\n";
                                //        par.Range.Text += "��� ������ � ������� ������ � ���������\n";

                                //        spi = Convert.ToInt32(drspi["NOMSPI"]);

                                //        //par.Range.Text += "��-�: " + GetSpiName3(Convert.ToInt32(drspi["NOMSPI"])) + "\n";
                                //        par.Range.Text += "��-�: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";
                                //        sch_line += 10;

                                //        foreach (DataRow row in tbl.Rows)
                                //        {
                                //            if (spi == Convert.ToInt32(row["NOMSPI"]))
                                //            {
                                //                par.Range.Text += GetIPNum(Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIO"]).TrimEnd() + "\n";
                                //                sch_line += 2;
                                //            }
                                //        }
                                //    }

                                //    app.Visible = true;
                                //    //*************************************************
                                //}
                                #endregion
# endregion
                            }
                            catch (OleDbException ole_ex)
                            {
                                foreach (OleDbError err in ole_ex.Errors)
                                {
                                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                                }
                                //return false;
                            }
                            catch (Exception ex)
                            {
                                //if (DBFcon != null) DBFcon.Close();
                                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                                //return false;
                            }
                            //return true;
                        }
                        # endregion
                        #region "E_TOFIND"
                        else if(openFileDialog1.FileName.ToLower().Contains("e_tofind.dbf"))
                        {
                            // ��� ���� � ��������
                            try
                            {
                                DataSet ds = new DataSet();
                                DataTable tbl = ds.Tables.Add("e_tofind");
                                DBFcon = new OleDbConnection();
                                DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=RUSSIAN");
                                //DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=MACHINE");
                                //DBFcon.

                                DBFcon.Open();
                                m_cmd = new OleDbCommand();
                                m_cmd.Connection = DBFcon;
                                // ������� ������� �� ������� tofind
                                m_cmd.CommandText = "select  distinct  NOMOSP, LITZDOLG, FIOVK, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FL_OKON, OSNOKON, OSNOKON from E_TOFIND ORDER BY ZAPROS";// ����������� �� ���� ZAPROS ����� �������� � ����������� �������� �� 1 ������ - ��� �� ���� �� ������, �� ��� ����� ��� ��� � ������
                                //m_cmd.CommandText = "SELECT * FROM FIND";// ����������� �� ���� ZAPROS ����� �������� � ����������� �������� �� 1 ������

                                using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                                {
                                    ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                                    rdr.Close();
                                }

                                DBFcon.Close();

                                Int32 iCnt = 0;
                                OleDbTransaction tran;

                                prbWritingDBF.Value = 0;
                                prbWritingDBF.Maximum = tbl.Rows.Count;
                                prbWritingDBF.Step = 1;

                                
                                Int32 i;
                                Decimal nID = 0;
                                String txtID = "";
                                nStatus = 15; // ������

                                nAgreementID = 0;
                                nAgent_dept_id = 0;
                                nAgent_id = 0;
                                nDx_pack_id = 0;
                                nNewPackID = 0;

                                txtAgreementCode = "";
                                txtAgentCode = "";
                                txtAgentDeptCode = "";
                                txtEntityName = "";


                                // ���� ���� � �������� �� ������, �� �� ������� ������ ����������
                                // ��b ������ �� ������� �� 14 ������ ��� ������ (��� �������)
                                if (tbl.Rows.Count > 0)
                                {
                                    decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["ZAPROS"]);
                                    if (FindSendlist(nFirstID, org)) // ��������� �������� org - ���������� �� ������ ��������, �������� ���� ���������� �����
                                    {
                                        // ������ ��� ����� ������
                                        // �������� ���������: ����������, ����������, �������������
                                        DataTable dtParams = GetPackParams(con, nFirstID, org);
                                        if ((dtParams != null) && (dtParams.Rows.Count > 0))
                                        {
                                            nAgreementID = Convert.ToDecimal(dtParams.Rows[0]["agreement_id"]);
                                            nAgent_dept_id = Convert.ToDecimal(dtParams.Rows[0]["agent_dept_id"]);
                                            nAgent_id = Convert.ToDecimal(dtParams.Rows[0]["agent_id"]);
                                            nDx_pack_id = Convert.ToDecimal(dtParams.Rows[0]["dx_pack_id"]);
                                        }
                                    }

                                    if (nAgreementID == 0)
                                    {
                                        nAgreementID = GetAgr_by_Org(org); // ����� ����������
                                        nAgent_id = GetAgent_ID(nAgreementID);
                                        nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                    }


                                    txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                    txtEntityName = GetLegal_Name(org);



                                    // ����� ������� ����� �������� �����
                                    // nNewPackID = ID_CreateDX_PACK_I(con, 1, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);
                                    // TODO: �������� ����� ������ �������, � �������� ������ �����
                                    frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD, "FL_E_TOFIND");
                                    nParentID = LogList.ShowForm();

                                    if (nParentID != -1)
                                    {
                                        // 1 - �����
                                        // 5 - �������� � ��������� �������
                                        nNewPackID = CreateLLog(conGIBDD, 1, 5, txtAgreementCode, nParentID, "����� ������� �� " + txtEntityName + ".");

                                        // �������� � ��� ������ ���� � ������ ���������
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                        //WritePackLog(con, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");

                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");


                                        for (i = 0; i < tbl.Rows.Count; i++)
                                        {
                                            DataRow row = tbl.Rows[i];
                                            txtID = Convert.ToString(row["ZAPROS"]);
                                            if (!Decimal.TryParse(txtID, out nID))
                                            {
                                                nID = 0;
                                            }
                                            if (FindZapros(nID))
                                            {
                                                // ������� �������� ��������� � ���� ��������� ������ ������
                                                try
                                                {
                                                    string txtDatOtv = "";
                                                    DateTime dtDatOtv;

                                                    txtDatOtv = Convert.ToString(row["DATZAPR2"]); // ��� ���� �������� ������� - ����� ��� � ����� ����� ������ � ���������� �������
                                                    if (!DateTime.TryParse(txtDatOtv, out dtDatOtv))
                                                    {
                                                        dtDatOtv = DateTime.MaxValue;
                                                    }

                                                    string txtDatZap = "";
                                                    DateTime dtDatZap;


                                                    // ��������� ���� �������
                                                    txtDatZap = Convert.ToString(row["DATZAPR"]);
                                                    if (!DateTime.TryParse(txtDatZap, out dtDatZap))
                                                    {
                                                        dtDatZap = DateTime.MaxValue;
                                                    }

                                                    bNotIntTablesResp = false; // ������ ��� ������ ����� �� ������������ ������
                                                    //if (dtDatZap < dtIntTablesDeplmntDate)
                                                    //{
                                                    //    bNotIntTablesResp = true;
                                                    //}
                                                    //else
                                                    //{
                                                    //    bNotIntTablesResp = false;
                                                    //}

                                                    string txtOtvet;
                                                    txtOtvet = "� ������������ � " + PKOSP_GetOrgConvention(org);
                                                    txtOtvet += " ������� �����: ";

                                                    txtOtvet += "����� �� " + GetLegal_Name(org);
                                                    txtOtvet += ". ������ �� ������ � ���������. ������ � ������ �������. ���������� ��������� ������������ ���������� ��. ��� ���. ���: ��� �������� (������), ���� �������� ��������. ��� ��. ���: ���, ������������ ��������.\n.";

                                                    //if (InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, org))
                                                    if (bNotIntTablesResp)
                                                    {
                                                        if (InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID))
                                                        {
                                                            iCnt++;
                                                            WritePackLog(con, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");
                                                            prbWritingDBF.PerformStep();
                                                            prbWritingDBF.Refresh();
                                                            System.Windows.Forms.Application.DoEvents();
                                                        }
                                                        else
                                                        {
                                                            // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                            WritePackLog(con, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                        {
                                                            iCnt++;
                                                            // WritePackLog(con, nNewPackID, "��������� ����� # " + nID.ToString() + "\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");

                                                            prbWritingDBF.PerformStep();
                                                            prbWritingDBF.Refresh();
                                                            System.Windows.Forms.Application.DoEvents();
                                                        }
                                                        else
                                                        {
                                                            // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                            // WritePackLog(con, nNewPackID, "������! ����� # " + nID.ToString() + " ���������� �� �������.\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                        }
                                                    }

                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                                                    if (nNewPackID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "������! �������� ������ ������� ��������� ����������.\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "�������� �������� = " + iCnt.ToString() + "\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");

                                                        if (nID > 0)
                                                        {
                                                            WriteLLog(conGIBDD, nNewPackID, "ID ������� = " + nID.ToString() + "\n");
                                                        }
                                                    }

                                                }
                                            }
                                            else
                                            {
                                                // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                if (nNewPackID > 0)
                                                {
                                                    WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ ID = " + nID.ToString() + " �� ������� ��������� �.�. �� ��������� ������-��������.\n");
                                                }
                                            }

                                        }
                                        //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                        //WritePackLog(con, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                        // ���������� ���������� ������������ ��������
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);


                                        // �������� ������ ����-������
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);


                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - �������� �����

                                        // �������� ���� ��� ��������� E_TOFIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_E_TOFIND");

                                        // ���������������� �.�. ������ �������� ������ �� ���������
                                        //// ���� ��� ��, �� ����� �������� ������ ������
                                        //if (nNewPackID > 0)
                                        //{
                                        //    SetDocumentStatus(nNewPackID, 70);
                                        //}
                                    }
                                }
                                else
                                {
                                    if (nAgreementID == 0)
                                    {
                                        nAgreementID = GetAgr_by_Org(org); // ����� ����������
                                        nAgent_id = GetAgent_ID(nAgreementID);
                                        nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                    }


                                    txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                    txtEntityName = GetLegal_Name(org);

                                    // ������� ��� ������������� ������� � �������� ���� ��� 0 � ������ �������
                                    frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD, "FL_E_TOFIND");
                                    nParentID = LogList.ShowForm();

                                    // ���� �� ���� ������� ���������� �������� ������
                                    if (nParentID != -1)
                                    {
                                        // 1 - �����
                                        // 5 - ����� �� ����� � ���������
                                        nNewPackID = CreateLLog(conGIBDD, 1, 5, txtAgreementCode, nParentID, "����� ������� �� " + txtEntityName + ".");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");
                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                        // ���������� ���������� ������������ ��������
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                        // �������� ������ ����-������
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                        // �������� ������ ����-��������
                                        // ����� ������ ������, �.�. ������ ��� ����� �� ������
                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - �������� �����
                                        // �������� ���� ��� ��������� E_TOFIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_E_TOFIND");
                                    }
                                }

                                MessageBox.Show("���������� �������: " + iCnt.ToString() + ".\n������ ����� ����������� ������ �������.", "���������", MessageBoxButtons.OK);

                                //**********������������**�������**find************
                                //���� ��������� ������ � ���� + ����������� ���������� ����� 
                                //��� ���������� �� ���������. 

                                //������ ���� ���������
                                DataTable dtspi = ds.Tables.Add("SPI");

                                DBFcon.Open();
                                m_cmd = new OleDbCommand();
                                m_cmd.Connection = DBFcon;
                                m_cmd.CommandText = "SELECT DISTINCT NOMSPI FROM E_TOFIND";

                                using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                                {
                                    ds.Load(rdr, LoadOption.OverwriteChanges, dtspi);
                                    rdr.Close();
                                }

                                DBFcon.Close();

                                #region "HTML print"
                                // ���� ���������� ��� HTML
                                prbWritingDBF.Value = 0;
                                prbWritingDBF.Maximum = dtspi.Rows.Count;
                                prbWritingDBF.Step = 1;
                                Int32 spi = 0;

                                ReportMaker report = new ReportMaker();
                                report.StartReport();
                                foreach (DataRow drspi in dtspi.Rows)
                                {
                                    bool fl_no_answer = true;
                                    report.AddToReport("<h3>");
                                    report.AddToReport("������ �� �������� � ��������� �������� ��-� � ������� ���. ������� �� �����<br>");
                                    report.AddToReport("" + GetLegal_Name(org) + " �� " + DateTime.Today.ToShortDateString() + "<br>");
                                    //report.AddToReport("�� ������ � " + Convert.ToDateTime(ds.Tables["E_TOFIND"].Rows[0]["DATZAPR1"]).ToShortDateString() + " �� " + Convert.ToDateTime(ds.Tables["E_TOFIND"].Rows[0]["DATZAPR1"]).ToShortDateString() + "<br>");

                                    spi = Convert.ToInt32(drspi["NOMSPI"]);

                                    report.AddToReport("��-�: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "<br>");
                                                                        
                                    report.AddToReport("</h3>");

                                    foreach (DataRow row in tbl.Rows)
                                    {
                                        report.AddToReport("<p>");
                                        if (spi == Convert.ToInt32(row["NOMSPI"]))
                                        {
                                            string txtResLine = GetIPNum(con, Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIOVK"]).TrimEnd();
                                            if (Convert.ToInt32(row["LITZDOLG"]) == 2)
                                            {
                                                txtResLine += " (" + Convert.ToInt32(row["GOD"]).ToString() + " �.�.)";
                                                //txtResLine += " (" + Convert.ToInt32(row["GODR"]).ToString() + " �.�.)";
                                            }
                                            report.AddToReport(txtResLine + "");
                                        }
                                        report.AddToReport("</p>");
                                    }


                                    report.SplitNewPage();
                                    prbWritingDBF.PerformStep();
                                    prbWritingDBF.Refresh();
                                    System.Windows.Forms.Application.DoEvents();

                                }
                                report.EndReport();
                                report.ShowReport();

                                #endregion

                                #region "OLD PRINT"
                                //if (OooIsInstall)
                                //{
                                //    //OOo start
                                //    OOo_Writer OOo_cld = new OOo_Writer();
                                //    OOo_cld.OOo_Cred(GetLegal_Name(org), "e_tofind", ds, con, this);
                                //}
                                //else
                                //{
                                //    // ��� �����

                                //    Microsoft.Office.Interop.Word.ApplicationClass app = new ApplicationClass();

                                //    object s1 = "";
                                //    object fl = false;
                                //    object t = WdNewDocumentType.wdNewBlankDocument;
                                //    object fl2 = true;

                                //    Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref s1, ref fl, ref t, ref fl2);

                                //    Paragraph par;

                                //    int spi;
                                //    int sch_line;
                                //    int fl_fst = 1;

                                //    string nline = "";

                                //    prbWritingDBF.Value = 0;
                                //    prbWritingDBF.Maximum = dtspi.Rows.Count;
                                //    prbWritingDBF.Step = 1;

                                //    foreach (DataRow drspi in dtspi.Rows)
                                //    {
                                //        sch_line = 0;
                                //        if (fl_fst == 1)
                                //        {
                                //            sch_line = 1;
                                //            fl_fst = 0;
                                //            par = doc.Paragraphs[1];
                                //        }
                                //        else
                                //        {
                                //            object oMissing = System.Reflection.Missing.Value;
                                //            par = doc.Paragraphs.Add(ref oMissing);
                                //            object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                                //            par.Range.InsertBreak(ref oPageBreak);
                                //        }

                                //        par.Range.Font.Name = "Courier";
                                //        par.Range.Font.Size = 8;
                                //        float a = par.Range.PageSetup.RightMargin;
                                //        float b = par.Range.PageSetup.LeftMargin;
                                //        float c = par.Range.PageSetup.TopMargin;
                                //        float d = par.Range.PageSetup.BottomMargin;

                                //        par.Range.PageSetup.RightMargin = 30;
                                //        par.Range.PageSetup.LeftMargin = 30;
                                //        par.Range.PageSetup.TopMargin = 20;
                                //        par.Range.PageSetup.BottomMargin = 20;

                                //        par.Range.Text += "������ �� �������� � ��������� �������� ��-� � ������� ���. ������� �� �����\n";
                                //        par.Range.Text += GetLegal_Name(org) + " �� " + DateTime.Today.ToShortDateString() + "\n";
                                //        par.Range.Text += "�� ������ � " + Convert.ToDateTime(tbl.Rows[0]["DATZAPR1"]).ToShortDateString() + " �� " + Convert.ToDateTime(tbl.Rows[0]["DATZAPR2"]).ToShortDateString() + "\n";

                                //        spi = Convert.ToInt32(drspi["NOMSPI"]);

                                //        sch_line = 0;
                                //        if (fl_fst == 1)
                                //        {
                                //            sch_line = 1;
                                //            fl_fst = 0;
                                //        }
                                //        par.Range.Text += PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";

                                //        //par.Range.Text += "����� ��             �������                            �����                      ����       �������\n";
                                //        par.Range.Text += "����� ��             �������      ��� ��������\n";

                                //        sch_line += 10;

                                //        foreach (DataRow row in tbl.Rows)
                                //        {
                                //            if (spi == Convert.ToInt32(row["NOMSPI"]))
                                //            {
                                //                string txtResLine = GetIPNum(Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIOVK"]).TrimEnd();
                                //                if (Convert.ToInt32(row["LITZDOLG"]) == 2)
                                //                {
                                //                    txtResLine += " (" + Convert.ToInt32(row["GOD"]).ToString() + " �.�.)";
                                //                }

                                //                par.Range.Text += txtResLine + "\n";

                                //            }
                                //        }

                                //        prbWritingDBF.PerformStep();
                                //        prbWritingDBF.Refresh();
                                //        System.Windows.Forms.Application.DoEvents();
                                //    }

                                //    app.Visible = true;
                                //    //*************************************************
                                //}
                                #endregion
                            }
                            catch (OleDbException ole_ex)
                            {
                                foreach (OleDbError err in ole_ex.Errors)
                                {
                                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                                }
                                //return false;
                            }
                            catch (Exception ex)
                            {
                                //if (DBFcon != null) DBFcon.Close();
                                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                                //return false;
                            }
                        }
                        # endregion
                        # region "FIND"
                        else if (openFileDialog1.FileName.ToLower().Contains("find.dbf"))
                        {
                            try
                            {
                                DataSet ds = new DataSet();
                                DataTable tbl = ds.Tables.Add("FIND");
                                DBFcon = new OleDbConnection();
                                DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=RUSSIAN");
                                //DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=MACHINE");
                                //DBFcon.

                                DBFcon.Open();
                                m_cmd = new OleDbCommand();
                                m_cmd.Connection = DBFcon;
                                //m_cmd.CommandText = "SELECT FIL, PRIZ, LITZDOLG, FIO, ADRES, NOMLS, PRIZS, OSTAT, RSCHET, OSTSCH, NOMOSP, ZAPROS FROM FIND ORDER BY ZAPROS";// ����������� �� ���� ZAPROS ����� �������� � ����������� �������� �� 1 ������
                                m_cmd.CommandText = "select distinct  fil, priz, litzdolg, fio, godr, adres, nomls, prizs, ostat, rschet, ostsch, nomosp, zapros, nomspi, nomip, datotv, flzprspi, datzpr1, datzpr2 from FIND ORDER BY ZAPROS";// ����������� �� ���� ZAPROS ����� �������� � ����������� �������� �� 1 ������
                                //m_cmd.CommandText = "SELECT * FROM FIND";// ����������� �� ���� ZAPROS ����� �������� � ����������� �������� �� 1 ������

                                using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                                {
                                    ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                                    rdr.Close();
                                }

                                DBFcon.Close();

                                Int32 iCnt = 0;
                                //OleDbTransaction tran;

                                prbWritingDBF.Value = 0;
                                prbWritingDBF.Maximum = tbl.Rows.Count;
                                prbWritingDBF.Step = 1;

                                //con.Open();
                                //tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                                Int32 i;
                                bool bMoreTanOne = false;
                                string txtCommonAnswer = "";
                                Decimal nID = 0;
                                String txtID = "";
                                Decimal newID = 0;
                                nStatus = 20; // ����� �������

                                // ������� ��������� txtPreamb
                                string txtPreamb = "� ������������ � " + PKOSP_GetOrgConvention(org);
                                txtPreamb += " ������� �����: ";
                                // txtPreamb += "����� �� " + GetLegal_Name(org);

                                // ������� ����� ������� txtCommonRespText
                                string txtCommonRespText = "";

                                nAgreementID = 0;
                                nAgent_dept_id = 0;
                                nAgent_id = 0;
                                nDx_pack_id = 0;
                                nNewPackID = 0;

                                txtAgreementCode = "";
                                txtAgentCode = "";
                                txtAgentDeptCode = "";
                                txtEntityName = "";

                                // ���� ���� � �������� �� ������, �� �� ������� ������ ����������
                                // ��b ������ �� ������� �� 14 ������ ��� ������ (��� �������)
                                if (tbl.Rows.Count > 0)
                                {
                                    decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["ZAPROS"]);
                                    if (FindSendlist(nFirstID, org)) // ��������� �������� org - ���������� �� ������ ��������, �������� ���� ���������� �����
                                    {
                                        // ������ ��� ����� ������
                                        // �������� ���������: ����������, ����������, �������������
                                        DataTable dtParams = GetPackParams(con, nFirstID, org);
                                        if ((dtParams != null) && (dtParams.Rows.Count > 0))
                                        {
                                            nAgreementID = Convert.ToDecimal(dtParams.Rows[0]["agreement_id"]);
                                            nAgent_dept_id = Convert.ToDecimal(dtParams.Rows[0]["agent_dept_id"]);
                                            nAgent_id = Convert.ToDecimal(dtParams.Rows[0]["agent_id"]);
                                            nDx_pack_id = Convert.ToDecimal(dtParams.Rows[0]["dx_pack_id"]);
                                        }
                                    }

                                    // ����� ������� ����� �������� �����
                                    if (nAgreementID == 0)
                                    {
                                        nAgreementID = GetAgr_by_Org(org); // ����� ����������
                                        nAgent_id = GetAgent_ID(nAgreementID);
                                        nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                    }

                                    txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                    txtEntityName = GetLegal_Name(org);

                                    //nNewPackID = ID_CreateDX_PACK_I(con, 1, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);
                                    frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD, "FL_FIND");
                                    nParentID = LogList.ShowForm();

                                    if (nParentID != -1)
                                    {
                                        // 1 - �����
                                        // 3 - ����� �������������
                                        nNewPackID = CreateLLog(conGIBDD, 1, 3, txtAgreementCode, nParentID, "����� ������� �� " + txtEntityName + ".");


                                        // �������� � ��� ������ ���� � ������ ���������
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                        //WritePackLog(con, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");


                                        for (i = 0; i < tbl.Rows.Count; i++)
                                        {
                                            DataRow row = tbl.Rows[i];

                                            // ������� ����� ������� txtCommonRespText
                                            // ������� ������� ������� txtCurrRowRespText
                                            // ���� ��� �� ��������� ������� � ����. ������ �������� ����������� ������, �� 
                                            // txtCommonRespText += txtCurrRowRespText;
                                            // ����� - txtTotalRespText = txtPreamb + txtResponseHeader + txtCommonRespText;
                                            // �������� txtTotalRespText
                                            // � ���������� �������� UPDATE ��� ������ ������������� �������.



                                            txtID = Convert.ToString(row["ZAPROS"]);
                                            if (!Decimal.TryParse(txtID, out nID))
                                            {
                                                nID = 0;
                                            }

                                            // ��������� - ��� ���� �� ����� ������ ����� FindZapros(nID)
                                            // ���������, ��� ������ ��� ����� ������ - FindSendlist(nID)

                                            if (FindZapros(nID))
                                            {
                                                // ������� �������� ��������� � ���� ��������� ������ ������
                                                try
                                                {
                                                    string txtDatOtv = "";
                                                    DateTime dtDatOtv;

                                                    txtDatOtv = Convert.ToString(row["DATOTV"]);
                                                    if (!DateTime.TryParse(txtDatOtv, out dtDatOtv))
                                                    {
                                                        dtDatOtv = DateTime.MaxValue;
                                                    }

                                                    string txtDatZap = "";
                                                    DateTime dtDatZap;


                                                    // ��������� ���� �������
                                                    txtDatZap = Convert.ToString(row["DATZPR2"]);
                                                    if (!DateTime.TryParse(txtDatZap, out dtDatZap))
                                                    {
                                                        dtDatZap = DateTime.MaxValue;
                                                    }


                                                    bNotIntTablesResp = false; // ������ ��� ������ ����� �� ������������ ������
                                                    //if (dtDatZap < dtIntTablesDeplmntDate)
                                                    //{
                                                    //    bNotIntTablesResp = true;
                                                    //}
                                                    //else
                                                    //{
                                                    //    bNotIntTablesResp = false;
                                                    //}

                                                    // ������� ������� ������� txtCurrRowRespText
                                                    string txtCurrRowRespText = "";

                                                    // ��� ���� � �������� ��� � ��� ��������, �����
                                                    txtCurrRowRespText += Convert.ToString(row["FIO"]).TrimEnd();
                                                    if (Convert.ToInt32(row["LITZDOLG"]) == 2)
                                                    {
                                                        txtCurrRowRespText += " (" + Convert.ToInt32(row["GODR"]).ToString() + " �.�.)";
                                                    }
                                                    txtCurrRowRespText += " " + Convert.ToString(row["ADRES"]).TrimEnd() + " ";

                                                    string priz = Convert.ToString(row["PRIZ"]).TrimEnd();
                                                    if (priz.Length > 0) txtCurrRowRespText += Convert.ToString(row["PRIZ"]).TrimEnd();

                                                    if ((row.Table.Columns.Contains("NOMLS")) && (row.Table.Columns.Contains("OSTAT")) && (Convert.ToString(row["NOMLS"]).TrimEnd() != ""))
                                                    {
                                                        string txtLs = Convert.ToString(row["NOMLS"]).TrimEnd();
                                                        txtCurrRowRespText += "�/�: " + txtLs + " ������� = " + Convert.ToDecimal(row["OSTAT"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtLs);
                                                    }

                                                    if ((row.Table.Columns.Contains("RSCHET")) && (row.Table.Columns.Contains("OSTSCH")) && (Convert.ToString(row["RSCHET"]).TrimEnd() != ""))
                                                    {
                                                        string txtRs = Convert.ToString(row["RSCHET"]).TrimEnd();
                                                        txtCurrRowRespText += "�/�: " + txtRs + " ������� = " + Convert.ToDecimal(row["OSTSCH"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtRs);
                                                    }

                                                    // ���� ��� �� ��������� ������� � ����. ������ �������� ����������� ������, �� 
                                                    if (i < tbl.Rows.Count - 1)
                                                    {
                                                        if (Convert.ToString(tbl.Rows[i + 1]["ZAPROS"]).TrimEnd() == Convert.ToString(row["ZAPROS"]).TrimEnd())
                                                        {
                                                            bMoreTanOne = true;
                                                            txtCommonRespText += txtCurrRowRespText + ", ";
                                                        }
                                                        else
                                                        {
                                                            bMoreTanOne = false;

                                                            txtCommonRespText += txtCurrRowRespText;

                                                            txtCommonRespText = txtPreamb + " " + txtCommonRespText;

                                                            // �������� � ���� - �������� Rewrite State
                                                            if (bNotIntTablesResp)
                                                            {
                                                                if (InsertZaprosTo_PK_OSP(con, nID, txtCommonRespText, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID))
                                                                {
                                                                    iCnt++;
                                                                    prbWritingDBF.PerformStep();
                                                                    prbWritingDBF.Refresh();
                                                                    System.Windows.Forms.Application.DoEvents();
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (InsertResponseIntTable(con, nID, txtCommonRespText, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                                {
                                                                    iCnt++;
                                                                    //WritePackLog(con, nNewPackID, "��������� ����� # " + nID.ToString() + "\n");
                                                                    WriteLLog(conGIBDD, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");
                                                                    prbWritingDBF.PerformStep();
                                                                    prbWritingDBF.Refresh();
                                                                    System.Windows.Forms.Application.DoEvents();
                                                                }
                                                                else
                                                                {
                                                                    //WritePackLog(con, nNewPackID, "������! ����� # " + nID.ToString() + " ���������� �� �������.\n");
                                                                    WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                                }
                                                            }
                                                            // �������� ������� � �������� ���������
                                                            txtCommonRespText = "";
                                                        }

                                                    }
                                                    else
                                                    {
                                                        // ������ ��� ��������� �������.
                                                        bMoreTanOne = false;

                                                        txtCommonRespText += txtCurrRowRespText;

                                                        txtCommonRespText = txtPreamb + " " + txtCommonRespText;

                                                        // �������� � ���� - �������� Rewrite State
                                                        if (bNotIntTablesResp)
                                                        {
                                                            if (InsertZaprosTo_PK_OSP(con, nID, txtCommonRespText, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID))
                                                            {
                                                                iCnt++;
                                                                // WritePackLog(con, nNewPackID, "��������� ����� # " + nID.ToString() + "\n");
                                                                WriteLLog(conGIBDD, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");
                                                                prbWritingDBF.PerformStep();
                                                                prbWritingDBF.Refresh();
                                                                System.Windows.Forms.Application.DoEvents();
                                                            }
                                                            else
                                                            {
                                                                // WritePackLog(con, nNewPackID,
                                                                WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (InsertResponseIntTable(con, nID, txtCommonRespText, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                            {
                                                                iCnt++;
                                                                // WritePackLog(con, nNewPackID, "��������� ����� # " + nID.ToString() + "\n");
                                                                WriteLLog(conGIBDD, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");
                                                                prbWritingDBF.PerformStep();
                                                                prbWritingDBF.Refresh();
                                                                System.Windows.Forms.Application.DoEvents();
                                                            }
                                                            else
                                                            {
                                                                // WritePackLog(con, nNewPackID, "������! ����� # " + nID.ToString() + " ���������� �� �������.\n");
                                                                WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                            }
                                                        }
                                                        // �������� ������� � �������� ���������
                                                        txtCommonRespText = "";
                                                    }

                                                    // ������ ���� ��������� ���� -������-�� ��� ���� :D - ���� �� ����� ����� �����������

                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                                                    if (nNewPackID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "������! �������� ������ ������� ��������� ����������.\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "�������� �������� = " + iCnt.ToString() + "\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");

                                                        if (nID > 0)
                                                        {
                                                            WriteLLog(conGIBDD, nNewPackID, "ID ������� = " + nID.ToString() + "\n");
                                                        }
                                                    }

                                                }
                                            }
                                            else
                                            {
                                                // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                if (nNewPackID > 0)
                                                {
                                                    WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ ID = " + nID.ToString() + " �� ������� ��������� �.�. �� ��������� ������-��������.\n");
                                                }
                                            }

                                        }

                                        //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                        //WritePackLog(con, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                        // ���������� ���������� ������������ ��������
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);


                                        // �������� ������ ����-������
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - �������� �����

                                        // �������� ���� ��� ��������� FIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_FIND");

                                        //TODO: ����� ����� ���� ��� � ���� �������� - ��� ��������

                                        // ���������������� �.�. ������ �������� ������ �� ���������   
                                        //// ���� ��� ��, �� ����� �������� ������ ������
                                        //if(nNewPackID >0){
                                        //    SetDocumentStatus(nNewPackID, 70);
                                        //}
                                    }
                                }
                                else
                                {

                                    if (nAgreementID == 0)
                                    {
                                        nAgreementID = GetAgr_by_Org(org); // ����� ����������
                                        nAgent_id = GetAgent_ID(nAgreementID);
                                        nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                    }


                                    txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                    txtEntityName = GetLegal_Name(org);

                                    // ������� ��� ������������� ������� � �������� ���� ��� 0 � ������ �������
                                    frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD, "FL_FIND");
                                    nParentID = LogList.ShowForm();

                                    // ���� �� ���� ������� ���������� �������� ������
                                    if (nParentID != -1)
                                    {
                                        // 1 - �����
                                        // 3 - ����� �������������
                                        nNewPackID = CreateLLog(conGIBDD, 1, 3, txtAgreementCode, nParentID, "����� ������� �� " + txtEntityName + ".");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");
                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                        // ���������� ���������� ������������ ��������
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                        // �������� ������ ����-������
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                        // �������� ������ ����-��������
                                        // ����� ������ ������, �.�. ������ ��� ����� �� ������
                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - �������� �����
                                        // �������� ���� ��� ��������� FIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_FIND");
                                    }


                                }

                                //tran.Commit();
                                //con.Close();
                                MessageBox.Show("���������� �������: " + iCnt.ToString() + ".\n������ ����� ����������� ������ �������.", "���������", MessageBoxButtons.OK);

                                //**********������������**�������**find************
                                //���� ��������� ������ � ���� + ����������� ���������� ����� 
                                //��� ���������� �� ���������. 

                                //������ ���� ���������
                                DataTable dtspi = ds.Tables.Add("SPI");

                                DBFcon.Open();
                                m_cmd = new OleDbCommand();
                                m_cmd.Connection = DBFcon;
                                m_cmd.CommandText = "SELECT DISTINCT NOMSPI FROM FIND";

                                using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                                {
                                    ds.Load(rdr, LoadOption.OverwriteChanges, dtspi);
                                    rdr.Close();
                                }

                                DBFcon.Close();

                                #region "HTML print"
                                // ���� ���������� ��� HTML
                                prbWritingDBF.Value = 0;
                                prbWritingDBF.Maximum = dtspi.Rows.Count;
                                prbWritingDBF.Step = 1;
                                Int32 spi = 0;

                                ReportMaker report = new ReportMaker();
                                report.StartReport();
                                foreach (DataRow drspi in dtspi.Rows)
                                {
                                    bool fl_no_answer = true;
                                    report.AddToReport("<h3>");
                                    report.AddToReport("������ ������� �� ������� ��-� � ������� ���. ������� �� �����<br>");
                                    report.AddToReport("" + GetLegal_Name(org) + " �� " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATOTV"]).ToShortDateString() + "<br>");
                                    report.AddToReport("�� ������ � " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATZPR1"]).ToShortDateString() + " �� " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATZPR2"]).ToShortDateString() + "<br>");

                                    spi = Convert.ToInt32(drspi["NOMSPI"]);

                                    report.AddToReport("��-�: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "<br>");
                                    report.AddToReport("</h3>");
                                    
                                                                        
                                    

                                    foreach (DataRow row in tbl.Rows)
                                    {
                                        report.AddToReport("<p>");
                                        if (spi == Convert.ToInt32(row["NOMSPI"]))
                                        {
                                            string txtResponse = "";
                                            if ((row.Table.Columns.Contains("NOMLS")) && (row.Table.Columns.Contains("OSTAT")) && (Convert.ToString(row["NOMLS"]).TrimEnd() != ""))
                                            {
                                                //txtResponse += "�/�: " + Convert.ToString(row["NOMLS"]).TrimEnd() + " ������� = " + Money_ToStr(Convert.ToDecimal(row["OSTAT"])).TrimEnd();
                                                string txtLs = Convert.ToString(row["NOMLS"]).TrimEnd();
                                                txtResponse += "�/�: " + txtLs + " ������� = " + Convert.ToDecimal(row["OSTAT"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtLs);
                                            }

                                            if ((row.Table.Columns.Contains("RSCHET")) && (row.Table.Columns.Contains("OSTSCH")) && (Convert.ToString(row["RSCHET"]).TrimEnd() != ""))
                                            {
                                                //txtResponse += "; �/�: " + Convert.ToString(row["RSCHET"]).TrimEnd() + " ������� = " + Money_ToStr(Convert.ToDecimal(row["OSTSCH"])).TrimEnd();
                                                string txtRs = Convert.ToString(row["RSCHET"]).TrimEnd();
                                                txtResponse += "�/�: " + txtRs + " ������� = " + Convert.ToDecimal(row["OSTSCH"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtRs);
                                            }

                                            string txtResLine = GetIPNum(con, Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIO"]).TrimEnd();
                                            if (Convert.ToInt32(row["LITZDOLG"]) == 2)
                                            {

                                                txtResLine += " (" + Convert.ToInt32(row["GODR"]).ToString() + " �.�.)";
                                                //txtResLine += " (" + Convert.ToString(row["GODR"]).TrimEnd('0').TrimEnd(',') + " �.�.)";
                                                //txtResLine += " (" + Convert.ToInt32(row["GODR"]).ToString() + " �.�.)";
                                            }
                                            txtResLine += " " + Convert.ToString(row["ADRES"]).TrimEnd() + " " + txtResponse + " " + Convert.ToString(row["PRIZ"]).TrimEnd();


                                            report.AddToReport(txtResLine + "<br>");
                                        }
                                        report.AddToReport("</p>");
                                    }


                                    report.SplitNewPage();
                                    prbWritingDBF.PerformStep();
                                    prbWritingDBF.Refresh();
                                    System.Windows.Forms.Application.DoEvents();

                                }
                                report.EndReport();
                                report.ShowReport();

                                #endregion

                                #region "OLD PRINT"
                                //if (OooIsInstall)
                                //{
                                //    //OOo start
                                //    OOo_Writer OOo_cld = new OOo_Writer();
                                //    OOo_cld.OOo_Cred(GetLegal_Name(org), "find", ds, con, this);
                                //}
                                //else
                                //{

                                //    // ��� �����
                                //    # region "WORD FIND"

                                //    //Microsoft.Office.Interop.Word.ApplicationClass app = new ApplicationClass();

                                //    //object s1 = "";
                                //    //object fl = false;
                                //    //object t = WdNewDocumentType.wdNewBlankDocument;
                                //    //object fl2 = true;

                                //    //Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref s1, ref fl, ref t, ref fl2);

                                //    //Paragraph par;

                                //    //int spi;
                                //    //int sch_line;
                                //    //int fl_fst = 1;

                                //    //string nline = "";

                                //    //prbWritingDBF.Value = 0;
                                //    //prbWritingDBF.Maximum = dtspi.Rows.Count;
                                //    //prbWritingDBF.Step = 1;

                                //    //foreach (DataRow drspi in dtspi.Rows)
                                //    //{
                                //    //    sch_line = 0;
                                //    //    if (fl_fst == 1)
                                //    //    {
                                //    //        sch_line = 1;
                                //    //        fl_fst = 0;
                                //    //        par = doc.Paragraphs[1];
                                //    //    }
                                //    //    else
                                //    //    {
                                //    //        object oMissing = System.Reflection.Missing.Value;
                                //    //        par = doc.Paragraphs.Add(ref oMissing);
                                //    //        object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                                //    //        par.Range.InsertBreak(ref oPageBreak);
                                //    //    }

                                //    //    par.Range.Font.Name = "Courier";
                                //    //    par.Range.Font.Size = 8;
                                //    //    float a = par.Range.PageSetup.RightMargin;
                                //    //    float b = par.Range.PageSetup.LeftMargin;
                                //    //    float c = par.Range.PageSetup.TopMargin;
                                //    //    float d = par.Range.PageSetup.BottomMargin;

                                //    //    par.Range.PageSetup.RightMargin = 30;
                                //    //    par.Range.PageSetup.LeftMargin = 30;
                                //    //    par.Range.PageSetup.TopMargin = 20;
                                //    //    par.Range.PageSetup.BottomMargin = 20;

                                //    //    par.Range.Text += "������ ������� �� ������� ��-� � ������� �������� ������� �� �����\n";
                                //    //    par.Range.Text += GetLegal_Name(org) + " �� " + Convert.ToDateTime(tbl.Rows[0]["DATOTV"]).ToShortDateString() + "\n";
                                //    //    par.Range.Text += "�� ������ � " + Convert.ToDateTime(tbl.Rows[0]["DATZPR1"]).ToShortDateString() + " �� " + Convert.ToDateTime(tbl.Rows[0]["DATZPR2"]).ToShortDateString() + "\n";

                                //    //    spi = Convert.ToInt32(drspi["NOMSPI"]);

                                //    //    sch_line = 0;
                                //    //    if (fl_fst == 1)
                                //    //    {
                                //    //        sch_line = 1;
                                //    //        fl_fst = 0;
                                //    //    }
                                //    //    par.Range.Text += PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";

                                //    //    //par.Range.Text += "����� ��             �������                            �����                      ����       �������\n";
                                //    //    par.Range.Text += "����� ��             �������      ��� ��������          �����                      ����       �������\n";

                                //    //    sch_line += 10;

                                //    //    foreach (DataRow row in tbl.Rows)
                                //    //    {
                                //    //        if (spi == Convert.ToInt32(row["NOMSPI"]))
                                //    //        {
                                //    //            string txtResponse = "";
                                //    //            if ((row.Table.Columns.Contains("NOMLS")) && (row.Table.Columns.Contains("OSTAT")) && (Convert.ToString(row["NOMLS"]).TrimEnd() != ""))
                                //    //            {
                                //    //                //txtResponse += "�/�: " + Convert.ToString(row["NOMLS"]).TrimEnd() + " ������� = " + Money_ToStr(Convert.ToDecimal(row["OSTAT"])).TrimEnd();
                                //    //                string txtLs = Convert.ToString(row["NOMLS"]).TrimEnd();
                                //    //                txtResponse += "�/�: " + txtLs + " ������� = " + Convert.ToDecimal(row["OSTAT"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtLs);
                                //    //            }

                                //    //            if ((row.Table.Columns.Contains("RSCHET")) && (row.Table.Columns.Contains("OSTSCH")) && (Convert.ToString(row["RSCHET"]).TrimEnd() != ""))
                                //    //            {
                                //    //                //txtResponse += "; �/�: " + Convert.ToString(row["RSCHET"]).TrimEnd() + " ������� = " + Money_ToStr(Convert.ToDecimal(row["OSTSCH"])).TrimEnd();
                                //    //                string txtRs = Convert.ToString(row["RSCHET"]).TrimEnd();
                                //    //                txtResponse += "�/�: " + txtRs + " ������� = " + Convert.ToDecimal(row["OSTSCH"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtRs);
                                //    //            }

                                //    //            string txtResLine = GetIPNum(Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIO"]).TrimEnd();
                                //    //            if (Convert.ToInt32(row["LITZDOLG"]) == 2)
                                //    //            {
                                //    //                txtResLine += " (" + Convert.ToInt32(row["GODR"]).ToString() + " �.�.)";
                                //    //            }
                                //    //            txtResLine += " " + Convert.ToString(row["ADRES"]).TrimEnd() + " " + txtResponse + " " + Convert.ToString(row["PRIZ"]).TrimEnd();


                                //    //            par.Range.Text += txtResLine + "\n";

                                //    //        }
                                //    //    }

                                //    //    prbWritingDBF.PerformStep();
                                //    //    prbWritingDBF.Refresh();
                                //    //    System.Windows.Forms.Application.DoEvents();
                                //    //}

                                //    //app.Visible = true;
                                //    ////*************************************************

                                //    # endregion
                                //}
                                #endregion
                            }
                            catch (OleDbException ole_ex)
                            {
                                foreach (OleDbError err in ole_ex.Errors)
                                {
                                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                                }
                                if (nNewPackID > 0)
                                {
                                    WriteLLog(conGIBDD, nNewPackID, "������! ������ �� ������� ��������� ��� ������.\n");
                                    // �������� ������ ����-������
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ����� �������� � �������
                                }
                                //return false;
                            }
                            catch (Exception ex)
                            {
                                //if (DBFcon != null) DBFcon.Close();
                                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                                if (nNewPackID > 0)
                                {
                                    WriteLLog(conGIBDD, nNewPackID, "������! ������ �� ������� ��������� ��� ������.\n");
                                    // �������� ������ ����-������
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ����� �������� � �������
                                }
                                //return false;
                            }
                            //return true;
                        }
                        # endregion
                    }

                }
            }
            else MessageBox.Show("������ ����������. �������� ����������� �� ������.", "��������!", MessageBoxButtons.OK);

        }

        private string getValuteByCod(string ls)
        {
            string txtRes = "";
            if (ls.Length >= 8)
            {
                string txtCod = ls.Trim().Substring(5, 3);
                switch (txtCod)
                {
                    case "810":
                        txtRes = "���.";
                        break;

                    case "840":
                        txtRes = "����.";
                        break;

                    case "978":
                        txtRes = "����";
                        break;

                    case "826":
                        txtRes = "���� �����.";
                        break;

                    case "392":
                        txtRes = "��. ����";
                        break;

                    case "756":
                        txtRes = "�����. �����";
                        break;

                    default:
                        txtRes = "������ � ����� " + txtCod;
                        break;
                }
            }
            
            return txtRes;
        }

        private void lblDat1_Click(object sender, EventArgs e)
        {

        }

        private void lblDat2_Click(object sender, EventArgs e)
        {

        }

       
        private void btnWriteKtfomsPksp_Click(object sender, EventArgs e)
        {
            btnWriteKtfomsDBF.Enabled = false;

            if (bReadFromCopy)
            {
                con = new OleDbConnection(constr2);
            }
            con = new OleDbConnection(constrRDB);

            // ��� ��������������� - ��������� UPDATE ����� ������� ��������
            //string txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 23 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 3 and req_type.sndl_contr_id = " + ktfoms_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// ��� �������������� ��������
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 11 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 2 and req_type.sndl_contr_id = " + ktfoms_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            // ��������� ������ �� ���������
            DT_ktfoms_reg = null;

            // DT_ktfoms_doc = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 11 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_NSI_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            // DT_ktfoms_doc = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 11 and zapr_d.docstatusid = 2 and ip_d.docstatusid = 9 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            
            // ������� ��� 14 ������
            // �������� �������� pack.agreement_id = 30 - ����� ����������
            //DT_ktfoms_doc = GetDataTableFromFB("select  pack.id as pack_id, 2 LITZDOLG, d_req.id ZAPROS, req.IPNO_NUM, req.div, req.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, d_req.doc_date DATZAPR,  req.ID_DBTR_ADR ADDR, req.ID_DBTR_BORN DATROZHD, req.ID_DBTRCLS,  req.DBTR_BORN_YEAR GOD, req.ID_DEBTSUM SUMMA, req.ID_DBTR_INN INNORG,  d_req.doc_number, req.ID_DEBTCLS_NAME VIDVZISK from dx_pack_o packo left join dx_pack pack on pack.id = packo.id join sendlist sl on pack.id = sl.dx_pack_id join o_ip req on sl.sendlist_o_id = req.id join document d_req on req.id = d_req.id join document ip_d on d_req.parent_id = ip_d.id join document dpack on pack.id = dpack.id join SPI on req.IP_EXEC_PRIST = spi.SUSER_ID where dpack.docstatusid = 23  and pack.agreement_id = 120 and packo.has_been_sent is null and d_req.docstatusid != 19 and d_req.docstatusid != 15 and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            DT_ktfoms_doc = GetDataTableFromFB("select 120 agreement_id, ext_request_id,  pack_id,  2 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID where mvv_agreement_code = 120 and processed = 0 and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");

            //if (bReadFromCopy)
            //{
            //    con = new OleDbConnection(constr1);
            //}

            btnWriteKtfomsDBF.Enabled = true;

            int iDocCnt = 0;
            if (DT_ktfoms_doc != null) iDocCnt = DT_ktfoms_doc.Rows.Count;
            lblReadKtfomsRowsValue.Text = iDocCnt.ToString();
            
            if (bDateFolderAdd)
            {
                CreatePathWithDate(ktfoms_path);
            }
            else
            {
                FolderExist(ktfoms_path);
            }

            // �������� ��� ������� �� osp_xxx.dbf
            string tablename = "osp_" + GetOSP_Num().ToString().PadLeft(2, '0') + ".dbf";

            Int64 cnt = WriteKrfomsToDBF(false, fullpath, tablename);
            //WriteToDBF_SBER(false, fullpath, "tofind.dbf");

            //lblWriteRowsValue.Text = cnt.ToString();
            lblKtfomsDbfValue.Text = cnt.ToString();
        }

        private int RowsInString(string txtString)
        {
            int length = txtString.Length;
            int secondSpace = 0;
            int firstSpace = 0;
            string sub = "";
            int charCounter = 0;
            int sch_line = 1; // ������� �����, ��� ������� ����, ���� ���� ��� ������
            while (length > 112)
            {
                if (charCounter + 112 < txtString.Length)
                {
                    secondSpace = txtString.IndexOf(' ', charCounter + 112);
                }
                else
                {
                    secondSpace = txtString.Length - 1;
                }
                if (secondSpace != -1)
                {
                    sub = txtString.Substring(0, secondSpace);
                }
                else
                {
                    sub = txtString;
                }
                firstSpace = sub.LastIndexOf(' ');
                sch_line++; // ���� ��� ������� ������ ��-�� �����
                length = length - firstSpace - 1 + charCounter;
                charCounter = firstSpace + 1;

            }
            
            return sch_line;
        }

        private void b_ansktfoms_Click(object sender, EventArgs e)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);
            decimal nParentID = 0;
            decimal nNewPackID = 0;

            
            openFileDialog1.Filter = "DBF �����(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();
            int iRewriteState = 1; // ������� ����� ���������� ������� �� ������ (����������� �������� � ������������)
            bool bNotIntTablesResp = false; // ���� ����� �� ������, ��������� ��� ������������ ������.
            if (res == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {
                    ChangeByte(openFileDialog1.FileName, 0x65, 30);
                    if (openFileDialog1.FileName.ToLower().Contains("osp"))
                    {
                        string tablename = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - 4);
                        tablename = tablename.Substring(tablename.LastIndexOf("\\") + 1);

                        try
                        {
                            DataSet ds = new DataSet();
                            DataTable tbl = ds.Tables.Add(tablename);
                            DBFcon = new OleDbConnection();
                            DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=RUSSIAN");
                            DBFcon.Open();
                            m_cmd = new OleDbCommand();
                            m_cmd.Connection = DBFcon;
                            m_cmd.CommandText = "SELECT * FROM " + tablename + " ORDER BY NOMSPI";
                            using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                            {
                                ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                                rdr.Close();
                            }

                            DBFcon.Close();

                            Int32 iCnt = 0;
                            //OleDbTransaction tran;

                            prbWritingDBF.Value = 0;
                            prbWritingDBF.Maximum = tbl.Rows.Count;
                            prbWritingDBF.Step = 1;

                            //con.Open();
                            //tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                            DateTime dat1 = DateTime.Today;
                            DateTime dat2 = DateTime.Today;
                            /*
                            if (tbl.Rows.Count > 0)
                            {
                                if (!(DateTime.TryParse(Convert.ToString(tbl.Rows[0]["DATZAPR1"]), out dat1)))
                                {
                                    dat1 = DateTime.Today;
                                }

                                if (!(DateTime.TryParse(Convert.ToString(tbl.Rows[0]["DATZAPR2"]), out dat2)))
                                {
                                    dat2 = DateTime.Today;
                                }
                            }
                            */
                            //dat1 = DatZapr1_ktfoms.Date;
                            //dat2 = DatZapr2_ktfoms.Date;

                            Decimal nID = 0;
                            String txtID = "";

                            string txtNum_IP = "";
                            string priz = "";
                            string txtResponse = "";

                            //--
                            decimal nAgreementID = 0;
                            decimal nAgent_dept_id = 0;
                            decimal nAgent_id = 0;
                            decimal nDx_pack_id = 0; //  ��������� - ����� �� �����
                            string txtAgreementCode = "";
                            string txtAgentCode = "";
                            string txtAgentDeptCode = "";


                            // ���� ���� � �������� �� ������, �� �� ������� ������ ����������
                            // ��b ������ �� ������� �� 14 ������ ��� ������ (��� �������)
                            if (tbl.Rows.Count > 0)
                            {
                                decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["ZAPROS"]);
                                if (FindSendlist(nFirstID, ktfoms_id))
                                {
                                    // ������ ��� ����� ������
                                    // �������� ���������: ����������, ����������, �������������
                                    DataTable dtParams = GetPackParams(con, nFirstID, ktfoms_id);
                                    if ((dtParams != null) && (dtParams.Rows.Count > 0))
                                    {

                                        if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agreement_id"]), out nAgreementID))
                                        {
                                            nAgreementID = 0;
                                        }

                                        if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agent_dept_id"]), out nAgent_dept_id))
                                        {
                                            nAgent_dept_id = 0;
                                        }

                                        if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agent_id"]), out nAgent_id))
                                        {
                                            nAgent_id = 0;
                                        }

                                        if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["dx_pack_id"]), out nDx_pack_id))
                                        {
                                            nDx_pack_id = 0;
                                        }

                                    }
                                }

                                // ���� ���������� �� �������, �� ������������� ��� ������
                                if (nAgreementID == 0)
                                {
                                    //GetAgr_by_Org - ���� �������� �������
                                    nAgreementID = 120;
                                    nAgent_id = GetAgent_ID(nAgreementID);
                                    nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                }

                                txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                string txtEntityName = GetLegal_Name(ktfoms_id);

                                // ����� ������� ����� �������� �����
                                //nNewPackID = ID_CreateDX_PACK_I(con, 1, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);

                                // TODO: �������� ����� ������ �������, � �������� ������ �����
                                frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD);
                                nParentID = LogList.ShowForm();

                                if (nParentID != -1)
                                {
                                    // 1 - �����
                                    // 2 - ����� �������
                                    nNewPackID = CreateLLog(conGIBDD, 1, 2, txtAgreementCode, nParentID, "����� ������� �� " + txtEntityName + ".");

                                    // �������� � ��� ������ ���� � ������ ���������
                                    WritePackLog(con, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                    WritePackLog(con, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");

                                    foreach (DataRow row in tbl.Rows)
                                    {
                                        txtID = Convert.ToString(row["ZAPROS"]);
                                        if (!Decimal.TryParse(txtID, out nID))
                                        {
                                            nID = 0;
                                        }

                                        if (FindZapros(nID))
                                        {
                                            // ������� �������� ��������� � ���� ��������� ������ ������
                                            try
                                            {
                                                string txtDatZap = "";
                                                DateTime dtDatOtv, dtDatZap;


                                                // ��������� ���� �������
                                                txtDatZap = Convert.ToString(row["DATZAPR"]);
                                                if (!DateTime.TryParse(txtDatZap, out dtDatZap))
                                                {
                                                    dtDatZap = DateTime.MaxValue;
                                                }

                                                bNotIntTablesResp = false; // ������ ��� ������ ����� �� ������������ ������
                                                //if (dtDatZap < dtIntTablesDeplmntDate)
                                                //{
                                                //    bNotIntTablesResp = true;
                                                //}
                                                //else
                                                //{
                                                //    bNotIntTablesResp = false;
                                                //}

                                                dtDatOtv = DateTime.Now;

                                                string txtOtvet;
                                                txtResponse = "";
                                                priz = Convert.ToString(row["PRIZ"]).TrimEnd();
                                                decimal nStatus = 0;

                                                // TODO: ����� ����� 
                                                txtResponse = "� ������������ � " + PKOSP_GetOrgConvention(ktfoms_id);
                                                txtResponse += " ������� �����: ";

                                                if (priz.ToUpper().Equals("T"))
                                                {
                                                    txtResponse += "������ ������������ ����� ������: " + Convert.ToString(row["NAMELONG"]).TrimEnd() + ".\n";
                                                    txtResponse += "��� ������������: " + Convert.ToString(row["FIO_BOSS"]).TrimEnd() + ".\n";
                                                    txtResponse += "������� ������������: " + Convert.ToString(row["TEL_BOSS"]).TrimEnd() + ".\n";
                                                    txtResponse += "����� ������������: " + Convert.ToString(row["ADR_PR"]).TrimEnd() + ".\n";
                                                    txtResponse += "����� ��������: " + Convert.ToString(row["ADRES"]).TrimEnd() + ".\n";
                                                    txtResponse += "��� �������� �����������: " + Convert.ToString(row["TYPE_DOG"]).TrimEnd() + ".\n";
                                                    txtResponse += "����� �������� �����������: " + Convert.ToString(row["N_DOG"]).TrimEnd() + ".\n";
                                                    nStatus = 20;

                                                }
                                                else
                                                {
                                                    //txtResponse = "��� ������ � �������� ��� ������� �� ������ � " + dat1.ToShortDateString() + " �� " + dat2.ToShortDateString();
                                                    txtResponse += "��� ������ � �������� �� ������� �� " + Convert.ToDateTime(row["DATZAPR"]).ToShortDateString();
                                                    nStatus = 7;
                                                }

                                                txtOtvet = txtResponse;

                                                // �������� � �������� ��������� ID �����������
                                                if (bNotIntTablesResp)
                                                {
                                                    if (InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, ktfoms_id, ref iRewriteState, nNewPackID))
                                                    {
                                                        iCnt++;
                                                        WritePackLog(con, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");
                                                        prbWritingDBF.PerformStep();
                                                        prbWritingDBF.Refresh();
                                                        System.Windows.Forms.Application.DoEvents();
                                                    }
                                                    else
                                                    {
                                                        // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                        WritePackLog(con, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                    }

                                                }
                                                else
                                                {
                                                    if (InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, ktfoms_id, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                    {
                                                        iCnt++;
                                                        //WritePackLog(con, nNewPackID, "��������� ����� # " + nID.ToString() + "\n");
                                                        WritePackLog(con, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");
                                                        prbWritingDBF.PerformStep();
                                                        prbWritingDBF.Refresh();
                                                        System.Windows.Forms.Application.DoEvents();
                                                    }
                                                    else
                                                    {
                                                        // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                        //WritePackLog(con, nNewPackID, "������! ����� # " + nID.ToString() + " ���������� �� �������.\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                                                if (nNewPackID > 0)
                                                {
                                                    WriteLLog(conGIBDD, nNewPackID, "������! �������� ������ ������� ��������� ����������.\n");
                                                    WriteLLog(conGIBDD, nNewPackID, "�������� �������� = " + iCnt.ToString() + "\n");
                                                    WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");
                                                    if (nID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "ID ������� = " + nID.ToString() + "\n");
                                                    }
                                                }

                                            }
                                        }
                                        else
                                        {
                                            // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                            // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                            if (nNewPackID > 0)
                                            {
                                                WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ � " + nID.ToString() + " �� ������� ��������� �.�. �� ��������� ������-��������.\n");
                                            }
                                        }

                                    }
                                    //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                    //WritePackLog(con, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                    //WritePackLog(con, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");
                                    WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                    WritePackLog(con, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                    WritePackLog(con, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                    // ���������� ���������� ������������ ��������
                                    UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                    // �������� ������ ����-������
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                    // �������� ������ ����-��������
                                    UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - �������� �����

                                    // ��� ������ ������ �� ��� � ������� ��� ������ �� ����
                                    //// ���� ��� ��, �� ����� �������� ������ ������
                                    //if (nNewPackID > 0)
                                    //{
                                    //    SetDocumentStatus(nNewPackID, 70);
                                    //}
                                }
                            }

                            MessageBox.Show("���������� �������: " + iCnt.ToString() + ".\n ������ ����� ����������� ������ �������.", "���������", MessageBoxButtons.OK);   
                            

                            //**********������������**�������**ktfoms************
                            //���� ��������� ������ � ���� + ����������� ���������� ����� 
                            //��� ���������� �� ���������. 

                            //������ ���� ���������
                            DataTable dtspi = ds.Tables.Add("SPI");

                            DBFcon.Open();
                            m_cmd = new OleDbCommand();
                            m_cmd.Connection = DBFcon;
                            m_cmd.CommandText = "SELECT DISTINCT NOMSPI FROM " + tablename + " WHERE PRIZ = 'T'";

                            using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                            {
                                ds.Load(rdr, LoadOption.OverwriteChanges, dtspi);
                                rdr.Close();
                            }

                            DBFcon.Close();

                            #region "HTML print"
                            // ���� ���������� ��� HTML
                            prbWritingDBF.Value = 0;
                            prbWritingDBF.Maximum = dtspi.Rows.Count;
                            prbWritingDBF.Step = 1;
                            Int32 spi = 0;

                            ReportMaker report = new ReportMaker();
                            report.StartReport();
                            foreach (DataRow drspi in dtspi.Rows)
                            {
                                bool fl_no_answer = true;
                                report.AddToReport("<h3>");
                                report.AddToReport( "������ ������� �� ������� ��-� � ������<br>");
                                report.AddToReport("����� �� ������ �� " + Convert.ToDateTime(tbl.Rows[0]["DATZAPR"]).ToShortDateString() + "<br>");
                                spi = Convert.ToInt32(drspi["NOMSPI"]);
                                report.AddToReport("��-�: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "<br>");
                                                                
                                report.AddToReport("</h3>");

                                foreach (DataRow row in tbl.Rows)
                                {
                                    report.AddToReport("<p>");
                                    priz = Convert.ToString(row["PRIZ"]).Trim();
                                    if (priz == "T")
                                    {
                                        if (spi == Convert.ToInt32(row["NOMSPI"]))
                                        {
                                            txtResponse = "";
                                            if (row.Table.Columns.Contains("NAME"))
                                            {
                                                txtResponse = GetIPNum(con, Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["NAME"]).TrimEnd() + " " + Convert.ToString(row["FNAME"]).TrimEnd() + " " + Convert.ToString(row["SNAME"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd();
                                            }
                                            else
                                            {
                                                txtResponse = GetIPNum(con, Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FAM"]).TrimEnd() + " " + Convert.ToString(row["IM"]).TrimEnd() + " " + Convert.ToString(row["OT"]).TrimEnd() + " " + Convert.ToDateTime(row["DD_R"]).ToShortDateString().TrimEnd();
                                            }

                                            report.AddToReport( txtResponse);

                                            if ((Convert.ToString(row["TYPE_DOG"]).Trim()) == "")
                                                txtResponse = "��� ������";
                                            else
                                                txtResponse = Convert.ToString(row["TYPE_DOG"]).TrimEnd() + ", " + Convert.ToString(row["NAMELONG"]).TrimEnd();

                                            if ((Convert.ToString(row["FIO_BOSS"]).Trim()) != "")
                                                txtResponse += ", ��� ������������: " + Convert.ToString(row["FIO_BOSS"]).TrimEnd();

                                            if ((Convert.ToString(row["TEL_BOSS"]).Trim()) != "")
                                                txtResponse += ", ������� ������������: " + Convert.ToString(row["TEL_BOSS"]).TrimEnd();

                                            if ((Convert.ToString(row["ADR_PR"]).Trim()) != "")
                                                txtResponse += ", ����� ������������: " + Convert.ToString(row["ADR_PR"]).TrimEnd();

                                            report.AddToReport( txtResponse);
                                           
                                            fl_no_answer = false;


                                        }
                                    }
                                    report.AddToReport("</p>");
                                }

                              
                                report.SplitNewPage();
                                prbWritingDBF.PerformStep();
                                prbWritingDBF.Refresh();
                                System.Windows.Forms.Application.DoEvents();

                            }
                            report.EndReport();
                            report.ShowReport();

                            #endregion

                            #region "OLD PRINT"
                            //if (OooIsInstall)
                            //{
                            //    //OOo start
                            //    OOo_Writer OOo_cld = new OOo_Writer();
                            //    OOo_cld.OOo_Ktfoms(tablename, ds, con, this);
                            //}
                            //else
                            //{

                            //    Microsoft.Office.Interop.Word.ApplicationClass app = new ApplicationClass();

                            //    object s1 = "";
                            //    object fl = false;
                            //    object t = WdNewDocumentType.wdNewBlankDocument;
                            //    object fl2 = true;

                            //    Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref s1, ref fl, ref t, ref fl2);
                            //    Paragraph par;

                            //    int spi;
                            //    int sch_line = 0;
                            //    int fl_fst = 1;

                            //    prbWritingDBF.Value = 0;
                            //    prbWritingDBF.Maximum = dtspi.Rows.Count;
                            //    prbWritingDBF.Step = 1;

                            //    foreach (DataRow drspi in dtspi.Rows)
                            //    {
                            //        sch_line = 0;
                            //        if (fl_fst == 1)
                            //        {
                            //            sch_line = 1;
                            //            fl_fst = 0;
                            //            par = doc.Paragraphs[1];
                            //        }
                            //        else
                            //        {
                            //            object oMissing = System.Reflection.Missing.Value;
                            //            par = doc.Paragraphs.Add(ref oMissing);
                            //            object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                            //            par.Range.InsertBreak(ref oPageBreak);
                            //        }

                            //        par.Range.Font.Name = "Courier";
                            //        par.Range.Font.Size = 8;
                            //        float a = par.Range.PageSetup.RightMargin;
                            //        float b = par.Range.PageSetup.LeftMargin;
                            //        float c = par.Range.PageSetup.TopMargin;
                            //        float d = par.Range.PageSetup.BottomMargin;

                            //        par.Range.PageSetup.RightMargin = 30;
                            //        par.Range.PageSetup.LeftMargin = 30;
                            //        par.Range.PageSetup.TopMargin = 20;
                            //        par.Range.PageSetup.BottomMargin = 20;

                            //        par.Range.Text += "������ ������� �� ������� ��-� � ������";
                            //        //par.Range.Text += "������ �� " + DateTime.Today.ToShortDateString() + "\n";
                            //        par.Range.Text += "����� �� ������ �� " + Convert.ToDateTime(tbl.Rows[0]["DATZAPR"]).ToShortDateString() + "\n";

                            //        // ����� �.�. dat1 � dat2 �������� ��������� � �������� ����� ������� � ������ �� �����
                            //        // par.Range.Text += "�� ������ � " + dat1.ToShortDateString() + " �� " + dat2.ToShortDateString() + "\n";

                            //        spi = Convert.ToInt32(drspi["NOMSPI"]);
                            //        par.Range.Text += "��-�: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";
                            //        sch_line += 5;

                            //        foreach (DataRow row in tbl.Rows)
                            //        {
                            //            priz = Convert.ToString(row["PRIZ"]).Trim();
                            //            if (priz == "T")
                            //            {
                            //                if (spi == Convert.ToInt32(row["NOMSPI"]))
                            //                {
                            //                    txtResponse = "";
                            //                    if (row.Table.Columns.Contains("NAME"))
                            //                    {
                            //                        txtResponse = Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["NAME"]).TrimEnd() + " " + Convert.ToString(row["FNAME"]).TrimEnd() + " " + Convert.ToString(row["SNAME"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd();
                            //                    }
                            //                    else
                            //                    {
                            //                        txtResponse = Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FAM"]).TrimEnd() + " " + Convert.ToString(row["IM"]).TrimEnd() + " " + Convert.ToString(row["OT"]).TrimEnd() + " " + Convert.ToDateTime(row["DD_R"]).ToShortDateString().TrimEnd();
                            //                    }

                            //                    par.Range.Text += txtResponse;

                            //                    if ((Convert.ToString(row["TYPE_DOG"]).Trim()) == "")
                            //                        txtResponse = "��� ������";
                            //                    else
                            //                        txtResponse = Convert.ToString(row["TYPE_DOG"]).TrimEnd() + ", " + Convert.ToString(row["NAMELONG"]).TrimEnd();

                            //                    if ((Convert.ToString(row["FIO_BOSS"]).Trim()) != "")
                            //                        txtResponse += ", ��� ������������: " + Convert.ToString(row["FIO_BOSS"]).TrimEnd();

                            //                    if ((Convert.ToString(row["TEL_BOSS"]).Trim()) != "")
                            //                        txtResponse += ", ������� ������������: " + Convert.ToString(row["TEL_BOSS"]).TrimEnd();

                            //                    if ((Convert.ToString(row["ADR_PR"]).Trim()) != "")
                            //                        txtResponse += ", ����� ������������: " + Convert.ToString(row["ADR_PR"]).TrimEnd();

                            //                    par.Range.Text += txtResponse;

                            //                    par.Range.Text += "";
                            //                    sch_line++;


                            //                }
                            //            }
                            //        }
                            //        prbWritingDBF.PerformStep();
                            //        prbWritingDBF.Refresh();
                            //        System.Windows.Forms.Application.DoEvents();
                            //    }

                            //    app.Visible = true;
                            //    //*************************************************
                            //}
#endregion
                        }
                        catch (OleDbException ole_ex)
                        {
                            foreach (OleDbError err in ole_ex.Errors)
                            {
                                MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                            }
                            if (nNewPackID > 0)
                            {
                                WriteLLog(conGIBDD, nNewPackID, "������! ������ �� ������� ��������� ��� ������.\n");
                                // �������� ������ ����-������
                                UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ����� �������� � �������
                            }
                            //return false;
                        }
                        catch (Exception ex)
                        {
                            //if (DBFcon != null) DBFcon.Close();
                            MessageBox.Show("������ ��� ������ � �������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                            WriteLLog(conGIBDD, nNewPackID, "������! ������ �� ������� ��������� ��� ������.\n");
                            // �������� ������ ����-������
                            UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ����� �������� � �������
                            //return false;
                        }
                        //return true;
                    }
                }
            }
        }




        private void b_anspens_Click(object sender, EventArgs e)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);
            decimal nNewPackID = 0;


            openFileDialog1.Filter = "DBF �����(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();
            int iRewriteState = 1; // ������� ����� ���������� ������� �� ������ (����������� �������� � ������������)
            DataTable tbl = null;
            decimal nStatus = 0;
            DataSet ds = null;
            bool bVFP_DBASE_local = false;
            OleDbConnection DbaseCon;
            bool bEx = false;
            string txtFileDir;
            
            bool bNotIntTablesResp = false; // ���� ����� �� ������, ��������� ��� ������������ ������.

            if (res == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {
                    ChangeByte(openFileDialog1.FileName, 0x65, 30);
                    if (openFileDialog1.FileName.ToLower().Contains("upf"))
                    {
                        string tablename = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - 4);
                        tablename = tablename.Substring(tablename.LastIndexOf("\\") + 1);
                        txtFileDir = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - tablename.Length - 4);


                        try
                        {
                            ds = new DataSet();
                            tbl = ds.Tables.Add(tablename);
                            nStatus = 0;
                            DBFcon = new OleDbConnection();
                            DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=RUSSIAN");
                            DBFcon.Open();
                            
                            m_cmd = new OleDbCommand();
                            m_cmd.Connection = DBFcon;
                            m_cmd.CommandText = "SELECT * FROM " + tablename;
                            using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                            {
                                ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                                rdr.Close();
                            }

                            DBFcon.Close();

                        }
                        catch (OleDbException ole_ex)
                        {
                            foreach (OleDbError err in ole_ex.Errors)
                            {
                                MessageBox.Show("������ ��� ������ � �������. ����� ����������� ��������� ������� ���������� ����. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                            }
                            bVFP_DBASE_local = true; // ����������� ���������� ����� DBASE
                        }

                        if (bVFP_DBASE_local)
                            {
                                try
                                {
                                    // ���� ��� ����� ������ 8 �������� - �� ���������� � ��������� �������
                                    string txtShortFileName = tablename;
                                    if (tablename.Length > 8)
                                    {
                                        txtShortFileName = tablename.Substring(0, 8) + ".dbf";
                                        File.Copy(openFileDialog1.FileName, txtFileDir + txtShortFileName);
                                    }
                                    else
                                    {
                                        txtShortFileName += ".dbf";
                                    }

                                    
                                    ds = new DataSet();
                                    tbl = ds.Tables.Add(tablename);
                                    DbaseCon = new OleDbConnection();
                                    DbaseCon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", txtFileDir);
                                    DbaseCon.Open();
                                    m_cmd = new OleDbCommand();
                                    m_cmd.Connection = DbaseCon;
                                    m_cmd.CommandText = "SELECT * FROM " + txtShortFileName;
                                    using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                                    {
                                        ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                                        rdr.Close();
                                    }

                                    DbaseCon.Close();
                                }
                                catch (OleDbException ole_ex)
                                {
                                    foreach (OleDbError err in ole_ex.Errors)
                                    {
                                        MessageBox.Show("������ ��� ������ � �������. ���� ���������� �� �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                                    }
                                    bEx = true;
                                }
                                bVFP_DBASE_local = false;
                            }

                        # region "��������� �����"
                        try{
                            if (!bEx)
                            {
                            Int32 iCnt = 0;
                            OleDbTransaction tran;

                            prbWritingDBF.Value = 0;
                            
                            if (tbl != null)
                                prbWritingDBF.Maximum = tbl.Rows.Count;
                            else prbWritingDBF.Maximum = 0;

                            prbWritingDBF.Step = 1;

                            //con.Open();
                            //tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                            DateTime dat1 = DateTime.Today;
                            DateTime dat2 = DateTime.Today;

                            dat1 = DatZapr1_pens.Date;
                            dat2 = DatZapr2_pens.Date;
                            
                            Decimal nID = 0;
                            String txtID = "";
                            Double sum;
                            DateTime dtDatZap;

                            decimal nAgreementID = 0;
                            decimal nAgent_dept_id = 0;
                            decimal nAgent_id = 0;
                            decimal nDx_pack_id = 0;

                            string txtAgreementCode = "";
                            string txtAgentCode = "";
                            string txtAgentDeptCode = "";

                            decimal nParentID = 0;

                            // ���� ���� � �������� �� ������, �� �� ������� ������ ����������
                            // ��b ������ �� ������� �� 14 ������ ��� ������ (��� �������)
                            if (tbl.Rows.Count > 0)
                            {
                                // ����� ��� �� ������
                                decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["NOMZAP"]);
                                if (FindSendlist(nFirstID, pens_id)) // ��������� �������� org - ���������� �� ������ ��������, �������� ���� ���������� �����
                                {
                                    // ������ ��� ����� ������
                                    // �������� ���������: ����������, ����������, �������������
                                    // TODO: ���� ����� ������ - ��������� �������� �� Agreement_Code ��� Agreement_ID
                                    DataTable dtParams = GetPackParams(con, nFirstID, pens_id);
                                    if ((dtParams != null) && (dtParams.Rows.Count > 0))
                                    {
                                        if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agreement_id"]), out nAgreementID))
                                        {
                                            nAgreementID = 0;
                                        }

                                        if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agent_dept_id"]), out nAgent_dept_id))
                                        {
                                            nAgent_dept_id = 0;
                                        }

                                        if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agent_id"]), out nAgent_id))
                                        {
                                            nAgent_id = 0;
                                        }

                                        if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["dx_pack_id"]), out nDx_pack_id))
                                        {
                                            nDx_pack_id = 0;
                                        }
                                        //nAgreementID = Convert.ToDecimal(dtParams.Rows[0]["agreement_id"]);
                                        //nAgent_dept_id = Convert.ToDecimal(dtParams.Rows[0]["agent_dept_id"]);
                                        //nAgent_id = Convert.ToDecimal(dtParams.Rows[0]["agent_id"]);
                                        //nDx_pack_id = Convert.ToDecimal(dtParams.Rows[0]["dx_pack_id"]);
                                    }
                                }

                                // ��� �������� ������� ��������� ���������� - �� Agreement_ID
                                if (nAgreementID == 0)
                                {
                                    //GetAgr_by_Org - ���� �������� �������
                                    nAgreementID = 100;
                                    nAgent_id = GetAgent_ID(nAgreementID);
                                    nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                }


                                txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                string txtEntityName = GetLegal_Name(pens_id);

                                // ������ �� ����� ��������� ����� - ������ ������� ���� ���
                                // ����� ������� ����� �������� �����
                                //nNewPackID = ID_CreateDX_PACK_I(con, 1, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);

                                // TODO: �������� ����� ������ �������, � �������� ������ �����
                                // ����� - ������� � ������������ ������ �� ����
                                // ������� ��������� - txtAgreementCode
                                frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD);
                                nParentID = LogList.ShowForm();


                                if (nParentID != -1)
                                {
                                    // 1 - �����
                                    // 2 - ����� �������
                                    nNewPackID = CreateLLog(conGIBDD, 1, 2, txtAgreementCode, nParentID, "����� ������� �� " + txtEntityName + ".");

                                    // �������� � ��� ������ ���� � ������ ���������

                                    //WritePackLog(con, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                    //WritePackLog(con, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");

                                    WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                    WriteLLog(conGIBDD, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");


                                    foreach (DataRow row in tbl.Rows)
                                    {
                                        txtID = Convert.ToString(row["NOMZAP"]);
                                        if (!Decimal.TryParse(txtID, out nID))
                                        {
                                            nID = 0;
                                        }
                                        // ���� ��� ������
                                        if (FindZapros(nID))
                                        {
                                            // ������� �������� ��������� � ���� ��������� ������ ������
                                            try
                                            {

                                                string txtDatZap = "";
                                                DateTime dtDatOtv;



                                                // ��� DATOTV � �����, ����� ������ ������� ���� ������
                                                // txtDatOtv = Convert.ToString(row["DATOTV"]);
                                                // if (!DateTime.TryParse(txtDatOtv, out dtDatOtv))
                                                // {
                                                //    dtDatOtv = DateTime.MaxValue;
                                                // }
                                                dtDatOtv = DateTime.Now;

                                                string txtOtvet;
                                                string txtResponse = "";
                                                int priz = 0;

                                                if (!(int.TryParse(Convert.ToString(row["PRIZ"]), out priz)))
                                                {
                                                    priz = 2;
                                                }

                                                if (!DateTime.TryParse(Convert.ToString(row["DATZAP"]), out dtDatZap))
                                                {
                                                    dtDatZap = DateTime.MaxValue;
                                                }

                                                txtResponse = "� ������������ � " + PKOSP_GetOrgConvention(pens_id);
                                                txtResponse += " ������� �����: ";

                                                if (priz == 1)
                                                {
                                                    txtResponse += "������� �������� ����������� ������.\n";
                                                    txtResponse += "�����: " + Convert.ToString(row["ADRES"]).TrimEnd() + "\n";
                                                    txtResponse += "C���� ������, �� ������� ����� �������� ���������: " + Convert.ToString(row["SUMMA"]).TrimEnd() + ". " + Convert.ToString(row["KOMMENT"]).TrimEnd() + "\n";
                                                    nStatus = 20; // ����� �������

                                                }
                                                else
                                                {
                                                    nStatus = 7; // ��� ������
                                                    if (priz == 0)
                                                    {
                                                        txtResponse += "��� ������ � �������� �� ������� �� " + dtDatZap.ToShortDateString();
                                                    }
                                                    else
                                                    {
                                                        txtResponse += "��� ������ � �������� �� ������� �� " + dtDatZap.ToShortDateString() + " " + Convert.ToString(row["SUMMA"]).TrimEnd();
                                                    }
                                                }
                                                txtOtvet = txtResponse;

                                                // TODO: ��� ���� �������� ����� ������� - ����� ������������ ������� =)
                                                // �������� ��������� - �����, agent_agreement, agent_dept_code, agent_code, enity_name

                                                bNotIntTablesResp = false; // ������ ��� ������ ����� �� ������������ ������
                                                //if (dtDatZap < dtIntTablesDeplmntDate)
                                                //{
                                                //    bNotIntTablesResp = true;
                                                //}
                                                //else
                                                //{
                                                //    bNotIntTablesResp = false;
                                                //}

                                                if (bNotIntTablesResp)
                                                {
                                                    if (InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, pens_id, ref iRewriteState, nNewPackID))
                                                    {
                                                        iCnt++;
                                                        WritePackLog(con, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");
                                                        prbWritingDBF.PerformStep();
                                                        prbWritingDBF.Refresh();
                                                        System.Windows.Forms.Application.DoEvents();
                                                    }
                                                    else
                                                    {
                                                        // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                        WritePackLog(con, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                        nStatus = 15; // ������
                                                    }
                                                }
                                                else
                                                {

                                                    if (InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, pens_id, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                    {
                                                        iCnt++;
                                                        //WritePackLog(con, nNewPackID, "��������� ����� # " + nID.ToString() + "\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");
                                                        prbWritingDBF.PerformStep();
                                                        prbWritingDBF.Refresh();
                                                        System.Windows.Forms.Application.DoEvents();
                                                    }
                                                    else
                                                    {
                                                        // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                        // WritePackLog(con, nNewPackID, "������! ����� # " + nID.ToString() + " ���������� �� �������.\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                        nStatus = 15; // ������
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                                                if (nNewPackID > 0)
                                                {
                                                    WriteLLog(conGIBDD, nNewPackID, "������! �������� ������ ������� ��������� ����������.\n");
                                                    WriteLLog(conGIBDD, nNewPackID, "�������� �������� = " + iCnt.ToString() + "\n");
                                                    WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");
                                                    if (nID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "ID ������� = " + nID.ToString() + "\n");
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                            if (nNewPackID > 0)
                                            {
                                                WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ ID = " + nID.ToString() + " �� ������� ��������� �.�. �� ��������� ������-��������.\n");
                                            }
                                        }
                                    }

                                    //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                    //WritePackLog(con, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                    //WritePackLog(con, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                    WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                    WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                    WriteLLog(conGIBDD, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                    // ���������� ���������� ������������ ��������
                                    UpdateLLogCount(conGIBDD, nNewPackID, iCnt);


                                    // �������� ������ ����-������
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                    // �������� ������ ����-��������
                                    // ��� ��� � ��� - ������ ����� 2, �������:
                                    //  - ���� ������� ������ = 2 (���������), �� ���������� ������ 12 (���������� ����� �������)
                                    //  - ���� ������� ������ = 12, �� ���������� ������ 10 (�������� �����)
                                    
                                    // ���� ����� ������ ��� �������� � �������� - �� ������ ������ �� ����
                                    if (nParentID > 0)
                                    {
                                        decimal nOldStatus = GetLLogStatus(conGIBDD, nParentID);
                                        if (nOldStatus == 2) UpdateLLogParentStatus(conGIBDD, nNewPackID, 12); // 12 (���������� ����� �������)
                                        else if (nOldStatus == 12) UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - �������� �����
                                    }

                                    // ��� ���� ��������� ��� ����� ������ ���:
                                    // UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - �������� �����

                                    // ����� �.�. ������ �������� ������ �� ���������
                                    //// ���� ��� ��, �� ����� �������� ������ ������
                                    //if (nNewPackID > 0)
                                    //{
                                    //    SetDocumentStatus(nNewPackID, 70);
                                    //}
                                }
                            }

                            MessageBox.Show("���������� �������: " + iCnt.ToString() + ".\n ������ ����� ����������� ������ �������.", "���������", MessageBoxButtons.OK);

                            //**********������������**�������**pens************
                            //���� ��������� ������ � ���� + ����������� ���������� ����� 
                            //��� ���������� �� ���������. 

                            //������ ���� ��������� � �������������� �������� ������ �� ������� tbl
                            DataTable dtspi = null;
                            if(ds != null){
                                dtspi = ds.Tables.Add("SPI");
                            }

                            string[] cols = new string[] { "NOMSPI" };
                            DataTable dtPriz1 = tbl.Clone();
                            dtPriz1.Clear();

                            DataRow[] Priz1Rows = tbl.Select("priz = '1'", "NOMSPI");

                            foreach (DataRow rowpriz1 in Priz1Rows)
                            {
                                dtPriz1.ImportRow(rowpriz1);
                            }
                            dtspi = SelectDistinct(dtPriz1, cols);

                            #region "HTML print"
                            // ���� ���������� ��� HTML
                            prbWritingDBF.Value = 0;
                            prbWritingDBF.Maximum = dtspi.Rows.Count;
                            prbWritingDBF.Step = 1;
                            Int32 spi = 0;

                            ReportMaker report = new ReportMaker();
                            report.StartReport();
                            foreach (DataRow drspi in dtspi.Rows)
                            {
                                bool fl_no_answer = true;
                                report.AddToReport("<h3>");
                                report.AddToReport("������ ������� �� ������� ��-� � ������������������� ������ � ���<br />");
                                report.AddToReport("������ �� ��� �� " + DateTime.Today.ToShortDateString() + "<br />");
                                //report.AddToReport("�� ������ � " + dat1.ToShortDateString() + " �� " + dat2.ToShortDateString() + "<br />");
                                spi = Convert.ToInt32(drspi["NOMSPI"]);
                                report.AddToReport("��-�: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "<br />");
                                report.AddToReport("</h3>");

                                foreach (DataRow row in tbl.Rows)
                                {
                                    report.AddToReport("<p>");                               
                                    if (spi == Convert.ToInt32(row["NOMSPI"]))
                                    {
                                        int priz = 0;
                                        if (!(int.TryParse(Convert.ToString(row["PRIZ"]), out priz)))
                                        {
                                            priz = 2;
                                        }
                                        if (priz == 1)
                                        {
                                            report.AddToReport( Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["NAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["FNAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["SNAMEDOL"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd());
                                            report.AddToReport( Convert.ToString(row["ADRES"]).TrimEnd() + "");
                                            report.AddToReport( "������� �������� ����������� ������. C���� ������, �� ������� ����� �������� ���������: " + Convert.ToString(row["SUMMA"]).TrimEnd() + "<br>");
                                            fl_no_answer = false;
                                        }
                                       
                                            
                                    }
                                    report.AddToReport("</p>");
                                }
                                
                                // ���� ������ �������������� � ������� ���, �� ��� � �����
                                if (fl_no_answer)
                                {
                                    report.AddToReport( "��� ������������� ������� �� �������� � ������� ������ � ���������.");                       
                                }
                                report.SplitNewPage();
                                prbWritingDBF.PerformStep();
                                prbWritingDBF.Refresh();
                                System.Windows.Forms.Application.DoEvents();

                            }
                            report.EndReport();
                            report.ShowReport();

                            #endregion

                            #region "old_print"
                            //if (OooIsInstall)
                            //{
                            //    //OOo start
                            //    OOo_Writer OOo_cld = new OOo_Writer();
                            //    OOo_cld.OOo_Pens(tablename, ds, con, this);
                            //}
                            //else
                            //{
                            //    //      ������ ��� �����

                            //    Microsoft.Office.Interop.Word.ApplicationClass app = new ApplicationClass();

                            //    object s1 = "";
                            //    object fl = false;
                            //    object t = WdNewDocumentType.wdNewBlankDocument;
                            //    object fl2 = true;

                            //    Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref s1, ref fl, ref t, ref fl2);

                            //    Paragraph par;

                            //    int spi;
                            //    int sch_line;
                            //    int fl_fst = 1;

                            //    string nline = "";

                            //    prbWritingDBF.Value = 0;
                            //    prbWritingDBF.Maximum = dtspi.Rows.Count;
                            //    prbWritingDBF.Step = 1;

                            //    foreach (DataRow drspi in dtspi.Rows)
                            //    {

                            //        sch_line = 0;
                            //        if (fl_fst == 1)
                            //        {
                            //            sch_line = 1;
                            //            fl_fst = 0;
                            //            par = doc.Paragraphs[1];
                            //        }
                            //        else
                            //        {
                            //            object oMissing = System.Reflection.Missing.Value;
                            //            par = doc.Paragraphs.Add(ref oMissing);
                            //            object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                            //            par.Range.InsertBreak(ref oPageBreak);
                            //        }

                            //        par.Range.Font.Name = "Courier";
                            //        par.Range.Font.Size = 8;
                            //        float a = par.Range.PageSetup.RightMargin;
                            //        float b = par.Range.PageSetup.LeftMargin;
                            //        float c = par.Range.PageSetup.TopMargin;
                            //        float d = par.Range.PageSetup.BottomMargin;

                            //        par.Range.PageSetup.RightMargin = 30;
                            //        par.Range.PageSetup.LeftMargin = 30;
                            //        par.Range.PageSetup.TopMargin = 20;
                            //        par.Range.PageSetup.BottomMargin = 20;

                            //        par.Range.Text += "������ ������� �� �������� ��-� � ��� � ������� ������\n";
                            //        par.Range.Text += "�� ������ � ��� �� " + DateTime.Today.ToShortDateString() + "\n";
                            //        //par.Range.Text += "�� ������ � " + dat1.ToShortDateString() + " �� " + dat2.ToShortDateString() + "\n";

                            //        spi = Convert.ToInt32(drspi["NOMSPI"]);

                            //        par.Range.Text += "��-�: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";

                            //        sch_line += 6;

                            //        foreach (DataRow row in tbl.Rows)
                            //        {
                            //            if (spi == Convert.ToInt32(row["NOMSPI"]))
                            //            {
                            //                //int priz = Convert.ToInt32(row["PRIZ"]);
                            //                int priz = 0;

                            //                if (!(int.TryParse(Convert.ToString(row["PRIZ"]), out priz)))
                            //                {
                            //                    priz = 2;
                            //                }

                            //                if (priz == 1)
                            //                {
                            //                    par.Range.Text += Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["NAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["FNAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["SNAMEDOL"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd();
                            //                    par.Range.Text += Convert.ToString(row["ADRES"]).TrimEnd() + "";
                            //                    par.Range.Text += "������� �������� ����������� ������. C���� ������, �� ������� ����� �������� ���������: " + Convert.ToString(row["SUMMA"]).TrimEnd() + "\n";
                            //                    sch_line += 5;
                            //                }

                            //                //if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                            //                //string priz = Convert.ToString(row["PRIZ"]).TrimEnd();
                            //                //if (priz.ToUpper().Equals("T"))
                            //            }
                            //        }
                            //        // ���� ������ �������������� � ������� ���, �� ��� � �����
                            //        if (sch_line == 6)
                            //        {
                            //            par.Range.Text += "��� ������������� ������� �� �������� � ������� ������ � ���������.";
                            //            sch_line++;
                            //            object oMissing = System.Reflection.Missing.Value;
                            //            par.Range.Delete(ref oMissing, ref oMissing);
                            //        }

                            //        prbWritingDBF.PerformStep();
                            //        prbWritingDBF.Refresh();
                            //        System.Windows.Forms.Application.DoEvents();
                            //    }

                            //    app.Visible = true;
                            //    //*************************************************
                            //}
                            #endregion
                        }
                        }
                        catch (OleDbException ole_ex)
                        {
                            foreach (OleDbError err in ole_ex.Errors)
                            {
                                MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                            }

                            if (nNewPackID > 0)
                            {
                                WriteLLog(conGIBDD, nNewPackID, "������! ������ �� ������� ��������� ��� ������.\n");
                                // �������� ������ ����-������
                                UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ����� �������� � �������
                            }

                            //return false;
                        }
                        catch (Exception ex)
                        {
                            //if (DBFcon != null) DBFcon.Close();
                            MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);

                            if (nNewPackID > 0)
                            {
                                WriteLLog(conGIBDD, nNewPackID, "������! ������ �� ������� ��������� ��� ������.\n");
                                // �������� ������ ����-������
                                UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ����� �������� � �������
                            }
                            //return false;
                        }
                        //return true;
                        # endregion
                    }
                }
            }
        }

        private void btnLoadPensData_Click(object sender, EventArgs e)
        {
            
           
        }


        private void btnWritePensDBF_Click(object sender, EventArgs e)
        {

            btnWritePensDBF.Enabled = false;

            if (bReadFromCopy)
            {
                con = new OleDbConnection(constr2);
            }

            con = new OleDbConnection(constrRDB);
            
            // ������������� ������� � �������� ��� ������ - ������� �� ������� �����. ��� ������� regl.sending_mode ���� ������

            //!!! ��� ��������������� - ����� �� ���� ������� �� �������� ������ - UPDATE ����� ����� ��������

            ///// �������� - ��� ������ � ������������� ��������� �������� ��������� ����������
            ///// �� ��������� ����������� �������� �������� � ������� � ������������ �������, ������� ���� ������� ���� �������
            //// d.docstatusid = 23 ��� ������ �������� ��������� ���������� 
            //string txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 23 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 3 and req_type.sndl_contr_id = " + pens_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// ��� �������������� ��������
            ////  d.docstatusid = 11 - ��� ������ �������������� ��������
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 11 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 2 and req_type.sndl_contr_id = " + pens_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// ���� ��� ������ 21 - ������ ��������
            //// ��� �������������� ��������
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 21 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 1 and req_type.sndl_contr_id = " + pens_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            DT_pens_reg = null;
            //DT_pens_doc = GetDataTableFromFB("SELECT DISTINCT c.PRIMARY_SITE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD, UPPER(a.name_d) as FIOVK, a.DATE_BORN_D as DATROZHD, a.SUM_ as SUMVZ, a.ADR_D as ADDR, a.ADR_D as ADDR, a.PK as FK_IP, a.PK_ID as FK_ID, a.uscode, a.FIO_SPI FROM IP a left join s_users c on (a.uscode=c.uscode) LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.SISP_KEY = '/1/3/' and a.DATE_IP_OUT is null and b.KOD != 1006 and (a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' and a.ssd is null and a.ssv is null", "TOFIND");
            //DT_pens_doc = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 15 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_NSI_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            //DT_pens_doc = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 15 and zapr_d.docstatusid = 2 and ip_d.docstatusid = 9 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            
            //// ������� ������� �������� �� ��
            //DT_pens_doc = GetDataTableFromFB("select pack.id as pack_id, 2 LITZDOLG, d_req.id ZAPROS, req.IPNO_NUM, req.div, req.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, d_req.doc_date DATZAPR,  req.ID_DBTR_ADR ADDR, req.ID_DBTR_BORN DATROZHD, req.ID_DBTRCLS,  req.DBTR_BORN_YEAR GOD, req.ID_DEBTSUM SUMMA, req.ID_DBTR_INN INNORG,  d_req.doc_number, req.ID_DEBTCLS_NAME VIDVZISK from dx_pack_o packo left join dx_pack pack on pack.id = packo.id join sendlist sl on pack.id = sl.dx_pack_id join o_ip req on sl.sendlist_o_id = req.id join document d_req on req.id = d_req.id join document ip_d on d_req.parent_id = ip_d.id join document dpack on pack.id = dpack.id join SPI on req.IP_EXEC_PRIST = spi.SUSER_ID where dpack.docstatusid = 23  and pack.agreement_id = 100 and packo.has_been_sent is null  and d_req.docstatusid != 19 and d_req.docstatusid != 15  and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");

            // ������� ������� �� ������������ ������ (��)
            // ext_request_id - ����� ����� ����� ������ ������� update
            DT_pens_doc = GetDataTableFromFB("select 100 agreement_id, ext_request_id,  pack_id,  2 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID where mvv_agreement_code = 100 and processed = 0 and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            
            // 100 - ����� ���������� � ��

            if (bReadFromCopy)
            {
                con = new OleDbConnection(constr1);
            }

            btnWritePensDBF.Enabled = true;

            int iDocCnt = 0;
            if (DT_pens_doc != null) iDocCnt = DT_pens_doc.Rows.Count;

            lblReadPensRowsValue.Text = iDocCnt.ToString();
            
            Int64 cnt;
            if (bDateFolderAdd)
            {
                CreatePathWithDate(pens_path);
            }
            else
            {
                FolderExist(pens_path);
            }

            string tablename = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString().PadLeft(2, '0') + DateTime.Today.Day.ToString().PadLeft(2, '0') + "_OSP_" + GetOSP_Num().ToString().PadLeft(2, '0') + "_1.dbf";
            cnt = WritePensToDBF(true, fullpath, tablename);
            lblPensDbfValue.Text = cnt.ToString(); 
            
        }

        
        private void dtpDatZapr1_krc_ValueChanged(object sender, EventArgs e)
        {
            if (dtpDatZapr1_krc.Value > dtpDatZapr2_krc.Value)
            {
                dtpDatZapr2_krc.Value = dtpDatZapr1_krc.Value;
            }

            DatZapr1_krc = dtpDatZapr1_krc.Value;
            DatZapr2_krc = dtpDatZapr2_krc.Value;
        }

        private void dtpDatZapr2_krc_ValueChanged(object sender, EventArgs e)
        {
            if (dtpDatZapr1_krc.Value > dtpDatZapr2_krc.Value)
            {
                dtpDatZapr1_krc.Value = dtpDatZapr2_krc.Value;
            }

            DatZapr1_krc = dtpDatZapr1_krc.Value;
            DatZapr2_krc = dtpDatZapr2_krc.Value;
        }

        
        private void UpdateZaprosTable(string txtConString, string txtUpdateZapros)
        {
            OleDbConnection UCon;
            OleDbTransaction UTran;
            OleDbCommand txtCmd;
            int nResult = 0;
            UCon = new OleDbConnection(txtConString);
            UCon.Open();
            UTran = UCon.BeginTransaction(IsolationLevel.RepeatableRead);
            try
            {
                txtCmd = UCon.CreateCommand();
                txtCmd.CommandText = txtUpdateZapros.Trim();
                txtCmd.Transaction = UTran;
                nResult = Convert.ToInt32(txtCmd.ExecuteScalar());
                if (nResult == -1)
                {
                    Exception ex = new Exception("Error operating with DB");
                    throw ex;
                }
                UTran.Commit();
                UCon.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception Thrown: " + e.ToString(), "��������!", MessageBoxButtons.OK);
                UTran.Rollback();
                if(UCon.State.Equals(ConnectionState.Open)) UCon.Close();
            }
        }


        private bool ChangeByte(string filename, byte Value, int position)
        {
            BinaryReader dataOut;
            BinaryWriter dataIn;
            FileStream fs;
            try
            {
                fs = new FileStream(filename, FileMode.Open);
                dataOut = new BinaryReader(fs);
            }
            catch(FileNotFoundException exc)
            {
                MessageBox.Show("������ ����������. ���� �� ������! Message: " + exc.ToString(), "��������!", MessageBoxButtons.OK);
                return false;
            }

            byte[] masBytes = dataOut.ReadBytes(Convert.ToInt32(fs.Length));
            dataOut.Close();
            fs.Close();
            fs.Dispose();

            if (masBytes[position - 1] != Value)
            {
                masBytes[position - 1] = Value;
                try
                {
                    fs = new FileStream(filename, FileMode.Create);
                    dataIn = new BinaryWriter(fs);
                }
                catch (FileNotFoundException exc)
                {
                    MessageBox.Show("������ ����������. ���� �� ������! Message: " + exc.ToString(), "��������!", MessageBoxButtons.OK);
                    return false;
                }
                dataIn.Seek(0, SeekOrigin.Begin);
                dataIn.Write(masBytes);
                dataIn.Close();
                fs.Close();
                fs.Dispose();
                return true;

            }
            else return true;
        }

              

        private void b_anspotd_Click(object sender, EventArgs e)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);
            decimal nParentID = 0;
            decimal nNewPackID = 0;

            openFileDialog1.Filter = "DBF �����(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();
            decimal nStatus = 0;
            int iRewriteState = 1; // ������� ����� ���������� ������� �� ������ (����������� �������� � ������������)
            bool bVFP_DBASE_local = false;
            DataTable tbl = null;
            DataSet ds = null;
            OleDbConnection DbaseCon;
            bool bEx = false;
            string txtFileDir;
            bool bNotIntTablesResp = false; // ���� ����� �� ������, ��������� ��� ������������ ������.

            if (res == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {
                    ChangeByte(openFileDialog1.FileName, 0x65, 30);
                    if (openFileDialog1.FileName.ToLower().Contains("opf"))
                    {
                        string tablename = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - 4);
                        tablename = tablename.Substring(tablename.LastIndexOf("\\") + 1);
                        txtFileDir = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - tablename.Length - 4);

                        try
                        {
                            ds = new DataSet();
                            tbl = ds.Tables.Add(tablename);
                            DBFcon = new OleDbConnection();
                            DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=RUSSIAN");
                            DBFcon.Open();
                            m_cmd = new OleDbCommand();
                            m_cmd.Connection = DBFcon;
                            m_cmd.CommandText = "SELECT * FROM " + tablename;
                            using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                            {
                                ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                                rdr.Close();
                            }

                            DBFcon.Close();
                        }
                        catch (OleDbException ole_ex)
                            {
                                foreach (OleDbError err in ole_ex.Errors)
                                {
                                    MessageBox.Show("������ ��� ������ � �������. ����� ����������� ��������� ������� ���������� ����. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                                }

                                bVFP_DBASE_local = true;
                            }
                        finally {
                            DBFcon.Dispose();
                            }

                            if (bVFP_DBASE_local)
                            {
                                try
                                {
                                    // ���� ��� ����� ������ 8 �������� - �� ���������� � ��������� �������
                                    string txtShortFileName = tablename;
                                    if (tablename.Length > 8)
                                    {
                                        txtShortFileName = tablename.Substring(0, 8) + ".dbf";
                                        File.Copy(openFileDialog1.FileName, txtFileDir + txtShortFileName);
                                    }
                                    else
                                    {
                                        txtShortFileName += ".dbf";
                                    }

                                    
                                    ds = new DataSet();
                                    tbl = ds.Tables.Add(tablename);
                                    DbaseCon = new OleDbConnection();
                                    DbaseCon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", txtFileDir);
                                    DbaseCon.Open();
                                    m_cmd = new OleDbCommand();
                                    m_cmd.Connection = DbaseCon;
                                    m_cmd.CommandText = "SELECT * FROM " + txtShortFileName;
                                    using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                                    {
                                        ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                                        rdr.Close();
                                    }

                                    DbaseCon.Close();
                                }
                                catch (OleDbException ole_ex)
                                {
                                    foreach (OleDbError err in ole_ex.Errors)
                                    {
                                        MessageBox.Show("������ ��� ������ � �������. ���� �������� ���������� �� �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                                    }
                                    bEx = true;
                                }
                                bVFP_DBASE_local = false;
                            }


                        
                    
                        # region "��������� �����"
                        // ���� ���� �������� ��� Exception
                            try
                            {
                                if (!bEx)
                                {

                                    Int32 iCnt = 0;
                                    OleDbTransaction tran;

                                    prbWritingDBF.Value = 0;
                                    prbWritingDBF.Maximum = tbl.Rows.Count;
                                    prbWritingDBF.Step = 1;

                                    //con.Open();
                                    //tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                                    DateTime dat1 = DateTime.Today;
                                    DateTime dat2 = DateTime.Today;

                                    dat1 = DatZapr1_potd.Date;
                                    dat2 = DatZapr2_potd.Date;
                                    /*
                                    if (tbl.Rows.Count > 0)
                                    {
                                        if (!(DateTime.TryParse(Convert.ToString(tbl.Rows[0]["DATZAPR1"]), out dat1)))
                                        {
                                            dat1 = DateTime.Today;
                                        }

                                        if (!(DateTime.TryParse(Convert.ToString(tbl.Rows[0]["DATZAPR2"]), out dat2)))
                                        {
                                            dat2 = DateTime.Today;
                                        }
                                    }
                                    */


                                    Decimal nID = 0;
                                    String txtID = "";
                                    nStatus = 0;

                                    decimal nAgreementID = 0;
                                    decimal nAgent_dept_id = 0;
                                    decimal nAgent_id = 0;
                                    decimal nDx_pack_id = 0;
                                    string txtAgreementCode = "";
                                    string txtAgentCode = "";
                                    string txtAgentDeptCode = "";


                                    // ���� ���� � �������� �� ������, �� �� ������� ������ ����������
                                    // ��b ������ �� ������� �� 14 ������ ��� ������ (��� �������)
                                    if (tbl.Rows.Count > 0)
                                    {
                                        decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["ZAPROS"]);
                                        if (FindSendlist(nFirstID, potd_id)) // ��������� �������� org - ���������� �� ������ ��������, �������� ���� ���������� �����
                                        {
                                            // ������ ��� ����� ������
                                            // �������� ���������: ����������, ����������, �������������
                                            DataTable dtParams = GetPackParams(con, nFirstID, potd_id);
                                            if ((dtParams != null) && (dtParams.Rows.Count > 0))
                                            {
                                                if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agreement_id"]), out nAgreementID))
                                                {
                                                    nAgreementID = 0;
                                                }

                                                if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agent_dept_id"]), out nAgent_dept_id))
                                                {
                                                    nAgent_dept_id = 0;
                                                }

                                                if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["agent_id"]), out nAgent_id))
                                                {
                                                    nAgent_id = 0;
                                                }

                                                if (!Decimal.TryParse(Convert.ToString(dtParams.Rows[0]["dx_pack_id"]), out nDx_pack_id))
                                                {
                                                    nDx_pack_id = 0;
                                                }
                                                //nAgreementID = Convert.ToDecimal(dtParams.Rows[0]["agreement_id"]);
                                                //nAgent_dept_id = Convert.ToDecimal(dtParams.Rows[0]["agent_dept_id"]);
                                                //nAgent_id = Convert.ToDecimal(dtParams.Rows[0]["agent_id"]);
                                                //nDx_pack_id = Convert.ToDecimal(dtParams.Rows[0]["dx_pack_id"]);
                                            }
                                        }


                                        if (nAgreementID == 0)
                                        {
                                            //GetAgr_by_Org - ���� �������� �������
                                            nAgreementID = 110;
                                            nAgent_id = GetAgent_ID(nAgreementID);
                                            nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                        }


                                        txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                        txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                        txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                        string txtEntityName = GetLegal_Name(potd_id);

                                        // ����� ������� ����� �������� �����
                                        // nNewPackID = ID_CreateDX_PACK_I(con, 1, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);

                                        // TODO: �������� ����� ������ �������, � �������� ������ �����
                                        frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD);
                                        nParentID = LogList.ShowForm();

                                        if (nParentID != -1)
                                        {
                                            // 1 - �����
                                            // 2 - ����� �������
                                            nNewPackID = CreateLLog(conGIBDD, 1, 2, txtAgreementCode, nParentID, "����� ������� �� " + txtEntityName + ".");

                                            // �������� � ��� ������ ���� � ������ ���������
                                            //WritePackLog(con, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                            //WritePackLog(con, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");
                                            WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ������ ��������� ������.\n");
                                            WriteLLog(conGIBDD, nNewPackID, "�������������� ����: " + openFileDialog1.FileName + "\n");


                                            foreach (DataRow row in tbl.Rows)
                                            {
                                                txtID = Convert.ToString(row["ZAPROS"]);
                                                if (!Decimal.TryParse(txtID, out nID))
                                                {
                                                    nID = 0;
                                                }
                                                if (FindZapros(nID))
                                                {
                                                    // ������� �������� ��������� � ���� ��������� ������ ������
                                                    try
                                                    {
                                                        string txtDatZap = "";
                                                        DateTime dtDatOtv, dtDatZap;


                                                        // ��������� ���� �������
                                                        txtDatZap = Convert.ToString(row["DATZAP"]);
                                                        if (!DateTime.TryParse(txtDatZap, out dtDatZap))
                                                        {
                                                            dtDatZap = DateTime.MaxValue;
                                                        }

                                                        bNotIntTablesResp = false; // ������ ��� ������ ����� �� ������������ ������
                                                        //if (dtDatZap < dtIntTablesDeplmntDate)
                                                        //{
                                                        //    bNotIntTablesResp = true;
                                                        //}
                                                        //else
                                                        //{
                                                        //    bNotIntTablesResp = false;
                                                        //}

                                                        dtDatOtv = DateTime.Now;

                                                        string txtOtvet;
                                                        string txtResponse = "";

                                                        txtResponse = "� ������������ � " + PKOSP_GetOrgConvention(potd_id);
                                                        txtResponse += " ������� �����: ";

                                                        if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                                                        {
                                                            txtResponse += "��� ������ �� ��������";
                                                            nStatus = 7; // ��� ������
                                                        }
                                                        else
                                                        {
                                                            txtResponse += "�����: " + Convert.ToString(row["ADRES"]).TrimEnd() + ".\n";
                                                            txtResponse += "������������ ������������: " + Convert.ToString(row["NAMEORG"]).TrimEnd() + ".\n";
                                                            txtResponse += "��������������� ������������: " + Convert.ToString(row["ADRORG"]).TrimEnd() + ".\n";
                                                            txtResponse += "���� ������ ������� ������: " + Convert.ToString(row["DATST"]).TrimEnd() + ".\n";
                                                            txtResponse += "���� ��������� ������� ������: " + Convert.ToString(row["DATFN"]).TrimEnd() + ".\n";
                                                            txtResponse += "�����������: " + Convert.ToString(row["KOMMENT"]).TrimEnd() + ".\n";
                                                            nStatus = 20; // ������� �����
                                                        }

                                                        txtOtvet = txtResponse;

                                                        if (bNotIntTablesResp)
                                                        {
                                                            if (InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, potd_id, ref iRewriteState, nNewPackID))
                                                            {
                                                                iCnt++;
                                                                WritePackLog(con, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");
                                                                prbWritingDBF.PerformStep();
                                                                prbWritingDBF.Refresh();
                                                                System.Windows.Forms.Application.DoEvents();
                                                            }
                                                            else
                                                            {
                                                                // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                                nStatus = 15; // ������
                                                                WritePackLog(con, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, potd_id, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                            {
                                                                iCnt++;
                                                                //WritePackLog(con, nNewPackID, "��������� ����� # " + nID.ToString() + "\n");
                                                                WriteLLog(conGIBDD, nNewPackID, "��������� ����� �� ������ # " + nID.ToString() + "\n");

                                                                prbWritingDBF.PerformStep();
                                                                prbWritingDBF.Refresh();
                                                                System.Windows.Forms.Application.DoEvents();
                                                            }
                                                            else
                                                            {
                                                                // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                                nStatus = 15; // ������
                                                                //WritePackLog(con, nNewPackID, "������! ����� # " + nID.ToString() + " ���������� �� �������.\n");
                                                                WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ # " + nID.ToString() + " ���������� �� �������.\n");
                                                            }
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                                                        if (nNewPackID > 0)
                                                        {
                                                            WriteLLog(conGIBDD, nNewPackID, "������! �������� ������ ������� ��������� ����������.\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "�������� �������� = " + iCnt.ToString() + "\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");
                                                            if (nID > 0)
                                                            {
                                                                WriteLLog(conGIBDD, nNewPackID, "ID ������� = " + nID.ToString() + "\n");
                                                            }
                                                        }

                                                    }
                                                }
                                                else
                                                {
                                                    // ����� �� ������� ���������, ���� �� ��� ���-�� � ������� ��������
                                                    //WritePackLog(con, nNewPackID, "������! ����� # " + nID.ToString() + " ���������� �� �������.\n");
                                                    if (nNewPackID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "������! ����� �� ������ ID = " + nID.ToString() + " �� ������� ��������� �.�. �� ��������� ������-��������.\n");
                                                    }
                                                }

                                            }
                                            //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                            //WritePackLog(con, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                            //WritePackLog(con, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");
                                            WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                            WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " ���������� ��������� ������.\n");
                                            WriteLLog(conGIBDD, nNewPackID, "���������� �������: " + iCnt.ToString() + "\n");

                                            // ���������� ���������� ������������ ��������
                                            UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                            // �������� ������ ����-������
                                            UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                            // �������� ������ ����-��������
                                            UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - �������� �����

                                            //// ���� ��� ��, �� ����� �������� ������ ������
                                            //if (nNewPackID > 0)
                                            //{
                                            //    SetDocumentStatus(nNewPackID, 70);
                                            //}
                                        }
                                    }

                                    MessageBox.Show("���������� �������: " + iCnt.ToString() + ".\n ������ ����� ����������� ������ �������.", "���������", MessageBoxButtons.OK);

                                    //**********������������**�������**pens************
                                    //���� ��������� ������ � ���� + ����������� ���������� ����� 
                                    //��� ���������� �� ���������. 

                                    //������ ���� ���������
                                    DataTable dtspi = ds.Tables.Add("SPI");

                                    // �������� ���� ��� �� ������� �������.
                                    // � ���� ������ �� ������ - ����� �� DataTable
                                    // �������� ������ �������
                                    string[] cols = new string[] { "NOMSPI" };
                                    dtspi = SelectDistinct(tbl, cols);

                                    // ���� ���������� ��� HTML
                                    prbWritingDBF.Value = 0;
                                    prbWritingDBF.Maximum = dtspi.Rows.Count;
                                    prbWritingDBF.Step = 1;
                                    Int32 spi = 0;

                                    ReportMaker report = new ReportMaker();
                                    report.StartReport();
                                    foreach (DataRow drspi in dtspi.Rows)
                                    {
                                        report.AddToReport("<h3>");
                                        report.AddToReport("������ ������� �� ������� ��-� � ������������������� ������ � ���<br />");
                                        report.AddToReport("������ �� ��� �� " + DateTime.Today.ToShortDateString() + "<br />");
                                        // report.AddToReport("�� ������ � " + dat1.ToShortDateString() + " �� " + dat2.ToShortDateString() + "<br />");
                                        spi = Convert.ToInt32(drspi["NOMSPI"]);
                                        report.AddToReport("��-�: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "<br />");
                                        report.AddToReport("</h3>");

                                        foreach (DataRow row in tbl.Rows)
                                        {
                                            if (spi == Convert.ToInt32(row["NOMSPI"]))
                                            {
                                                report.AddToReport("<br />");
                                                report.AddToReport(Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FNAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["NAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["SNAMEDOL"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd() + "<br />");
                                                if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                                                    report.AddToReport("��� ������ �� ��������<br />");
                                                else
                                                {
                                                    report.AddToReport("������������ ������������: " + Convert.ToString(row["NAMEORG"]).TrimEnd() + ".<br />");
                                                    report.AddToReport("��������������� ������������: " + Convert.ToString(row["ADRORG"]).TrimEnd() + ".<br />");

                                                    try
                                                    {
                                                        report.AddToReport("���� ������ ������� ������: " + Convert.ToDateTime(row["DATST"]).ToShortDateString() + ".<br />");
                                                        report.AddToReport("���� ��������� ������� ������: " + Convert.ToDateTime(row["DATFN"]).ToShortDateString() + ".<br />");
                                                    }
                                                    catch
                                                    {
                                                        // ��� ����� catch - �� ����� ���� ���� DateTime ������� ���������

                                                    }
                                                    //par.Range.Text += "���� ������ ������� ������: " + Convert.ToDateTime(row["DATST"]).ToShortDateString() + ".";
                                                    //par.Range.Text += "���� ��������� ������� ������: " + Convert.ToDateTime(row["DATFN"]).ToShortDateString() + ".";

                                                    report.AddToReport("�����������: " + Convert.ToString(row["KOMMENT"]).TrimEnd() + "<br />");
                                                }
                                            }
                                            
                                        }
                                        report.SplitNewPage();

                                        prbWritingDBF.PerformStep();
                                        prbWritingDBF.Refresh();
                                        System.Windows.Forms.Application.DoEvents();

                                    }
                                    report.EndReport();
                                    report.ShowReport();

                                    # region "OLD_FMS_REPORT"
                                    //                                if (OooIsInstall)
                                    //                            {
                                    //                                //OOo start
                                    //                                OOo_Writer OOo_cld = new OOo_Writer();
                                    //                                OOo_cld.OOo_Potd(tablename, ds, con, this);
                                    //                            }
                                    //                            else
                                    //                            {
                                    //                                //      ������ ��� �����

                                    //                                Microsoft.Office.Interop.Word.ApplicationClass app = new ApplicationClass();

                                    //                                object s1 = "";
                                    //                                object fl = false;
                                    //                                object t = WdNewDocumentType.wdNewBlankDocument;
                                    //                                object fl2 = true;

                                    //                                Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref s1, ref fl, ref t, ref fl2);

                                    //                                Paragraph par;


                                    //                                int spi;
                                    //                                int sch_line;
                                    //                                int fl_fst = 1;

                                    //                                DateTime datepens = new DateTime();


                                    //                                prbWritingDBF.Value = 0;
                                    //                                prbWritingDBF.Maximum = dtspi.Rows.Count;
                                    //                                prbWritingDBF.Step = 1;

                                    //                                foreach (DataRow drspi in dtspi.Rows)
                                    //                                {
                                    //                                    sch_line = 0;
                                    //                                    if (fl_fst == 1)
                                    //                                    {
                                    //                                        sch_line = 1;
                                    //                                        fl_fst = 0;
                                    //                                        par = doc.Paragraphs[1];
                                    //                                    }
                                    //                                    else
                                    //                                    {
                                    //                                        object oMissing = System.Reflection.Missing.Value;
                                    //                                        par = doc.Paragraphs.Add(ref oMissing);
                                    //                                        object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                                    //                                        par.Range.InsertBreak(ref oPageBreak);
                                    //                                    }

                                    //                                    par.Range.Font.Name = "Courier";
                                    //                                    par.Range.Font.Size = 8;
                                    //                                    float a = par.Range.PageSetup.RightMargin;
                                    //                                    float b = par.Range.PageSetup.LeftMargin;
                                    //                                    float c = par.Range.PageSetup.TopMargin;
                                    //                                    float d = par.Range.PageSetup.BottomMargin;

                                    //                                    par.Range.PageSetup.RightMargin = 30;
                                    //                                    par.Range.PageSetup.LeftMargin = 30;
                                    //                                    par.Range.PageSetup.TopMargin = 20;
                                    //                                    par.Range.PageSetup.BottomMargin = 20;

                                    //                                    par.Range.Text += "������ ������� �� ������� ��-� � ������������������� ������ � ���\n";
                                    //                                    par.Range.Text += "������ �� ��� �� " + DateTime.Today.ToShortDateString() + "\n";
                                    //                                    par.Range.Text += "�� ������ � " + dat1.ToShortDateString() + " �� " + dat2.ToShortDateString() + "\n";

                                    //                                    spi = Convert.ToInt32(drspi["NOMSPI"]);

                                    //                                    sch_line = 0;
                                    //                                    if (fl_fst == 1)
                                    //                                    {
                                    //                                        sch_line = 1;
                                    //                                        fl_fst = 0;
                                    //                                    }
                                    //                                    //par.Range.Text += " ";
                                    //                                    par.Range.Text += "��-�: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";
                                    //                                    //par.Range.Text += GetOSP_Name();
                                    //                                    sch_line += 8;

                                    //                                    foreach (DataRow row in tbl.Rows)
                                    //                                    {
                                    //                                        if (spi == Convert.ToInt32(row["NOMSPI"]))
                                    //                                        {
                                    //                                            par.Range.Text += Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FNAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["NAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["SNAMEDOL"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd();
                                    //                                            if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                                    //                                                par.Range.Text += "��� ������ �� ��������\n";
                                    //                                            else
                                    //                                            {
                                    //                                                par.Range.Text += "������������ ������������: " + Convert.ToString(row["NAMEORG"]).TrimEnd() + ".";
                                    //                                                par.Range.Text += "��������������� ������������: " + Convert.ToString(row["ADRORG"]).TrimEnd() + ".";

                                    //                                                try 
                                    //                                                {
                                    //                                                    par.Range.Text += "���� ������ ������� ������: " + Convert.ToDateTime(row["DATST"]).ToShortDateString() + ".";
                                    //                                                    par.Range.Text += "���� ��������� ������� ������: " + Convert.ToDateTime(row["DATFN"]).ToShortDateString() + ".";                                                
                                    //                                                }
                                    //                                                catch
                                    //                                                {

                                    //                                                }
                                    //                                                //par.Range.Text += "���� ������ ������� ������: " + Convert.ToDateTime(row["DATST"]).ToShortDateString() + ".";
                                    //                                                //par.Range.Text += "���� ��������� ������� ������: " + Convert.ToDateTime(row["DATFN"]).ToShortDateString() + ".";

                                    //                                                par.Range.Text += "�����������: " + Convert.ToString(row["KOMMENT"]).TrimEnd() + "\n";
                                    //                                                sch_line++;
                                    //                                            }
                                    //                                            sch_line += 3;
                                    //                                        }
                                    //                                    }

                                    //                                    prbWritingDBF.PerformStep();
                                    //                                    prbWritingDBF.Refresh();
                                    //                                    System.Windows.Forms.Application.DoEvents();
                                    //                                }

                                    //                                app.Visible = true;
                                    //                                //*************************************************
                                    //                            }
                                    # endregion
                                }
                            }

                            catch (OleDbException ole_ex)
                            {
                                foreach (OleDbError err in ole_ex.Errors)
                                {
                                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                                }
                                
                                if (nNewPackID > 0)
                                {
                                    WriteLLog(conGIBDD, nNewPackID, "������! ������ �� ������� ��������� ��� ������.\n");
                                    // �������� ������ ����-������
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ����� �������� � �������
                                }

                                //return false;
                            }
                            catch (Exception ex)
                            {
                                //if (DBFcon != null) DBFcon.Close();
                                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);

                                if (nNewPackID > 0)
                                {
                                    WriteLLog(conGIBDD, nNewPackID, "������! ������ �� ������� ��������� ��� ������.\n");
                                    // �������� ������ ����-������
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ����� �������� � �������
                                }

                                //return false;
                            }
                            //return true;
                    # endregion
                    }
                }
            }
        }

        private void btnWritePotdDBF_Click(object sender, EventArgs e)
        {
            btnWritePotdDBF.Enabled = false;

            if (bReadFromCopy)
            {
                con = new OleDbConnection(constr2);
            }

            con = new OleDbConnection(constrRDB);

            // ��� ��������������� - ������ ��� update �������� ������� ����� �������
            
            ///// �������� - ��� ������ � ������������� ��������� �������� ��������� ����������
            ///// �� ��������� ����������� �������� �������� � ������� � ������������ �������, ������� ���� ������� ���� �������
            //// d.docstatusid = 23 ��� ������ �������� ��������� ���������� 
            //string txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 23 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 3 and req_type.sndl_contr_id = " + potd_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// ��� �������������� ��������
            ////  d.docstatusid = 11 - ��� ������ �������������� ��������
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 11 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 2 and req_type.sndl_contr_id = " + potd_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// ���� ��� ������ 21 - ������ ��������
            //// ��� �������������� ��������
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 21 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 1 and req_type.sndl_contr_id = " + potd_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);
                             

            //select 2 LITZDOLG, zapr_d.id ZAPROS, ip.div, ip.ID_DBTR_NAME FIOVK, ip.IPNO NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 15 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_NSI_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))
            //ReadPOTDData(DatZapr1_potd, DatZapr2_potd);
            //DT_potd_doc = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 206 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_NSI_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            //DT_potd_doc = GetDataTableFromFB("select pack.id as pack_id, 2 LITZDOLG, d_req.id ZAPROS, req.IPNO_NUM, req.div, req.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, d_req.doc_date DATZAPR,  req.ID_DBTR_ADR ADDR, req.ID_DBTR_BORN DATROZHD, req.ID_DBTRCLS,  req.DBTR_BORN_YEAR GOD, req.ID_DEBTSUM SUMMA, req.ID_DBTR_INN INNORG,  d_req.doc_number, req.ID_DEBTCLS_NAME VIDVZISK from dx_pack_o packo left join dx_pack pack on pack.id = packo.id join sendlist sl on pack.id = sl.dx_pack_id join o_ip req on sl.sendlist_o_id = req.id join document d_req on req.id = d_req.id join document ip_d on d_req.parent_id = ip_d.id join document dpack on pack.id = dpack.id join SPI on req.IP_EXEC_PRIST = spi.SUSER_ID where dpack.docstatusid = 23  and pack.agreement_id = 110 and packo.has_been_sent is null  and d_req.docstatusid != 19 and d_req.docstatusid != 15   and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");

            // ������� ������� �� ������������ ������ (��)
            // ext_request_id - ����� ����� ����� ������ ������� update
            DT_potd_doc = GetDataTableFromFB("select 110 agreement_id, ext_request_id,  pack_id,  2 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID where mvv_agreement_code = 110 and processed = 0 and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            
            // 110

            if (bReadFromCopy)
            {
                con = new OleDbConnection(constr1);
            }

            btnWritePotdDBF.Enabled = true;

            lblReadPotdRowsValue.Text = (DT_potd_doc.Rows.Count).ToString();

            
            Int64 cnt;
            if (bDateFolderAdd)
            {
                CreatePathWithDate(potd_path);
            }
            else
            {
                FolderExist(potd_path);
            }

            string tablename = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString().PadLeft(2, '0') + DateTime.Today.Day.ToString().PadLeft(2, '0') + "_OSP_" + GetOSP_Num().ToString().PadLeft(2, '0') + "_2.dbf";            
            cnt = WritePotdToDBF(true, fullpath, tablename);
            lblPotdDbfValue.Text = cnt.ToString(); 
           

            //string tablename = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString().PadLeft(2, '0') + DateTime.Today.Day.ToString().PadLeft(2, '0') + "_OSP_" + GetOSP_Num().ToString().PadLeft(2, '0') + "_2.dbf";
            //string tablename_n = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString().PadLeft(2, '0') + DateTime.Today.Day.ToString().PadLeft(2, '0') + "_OSP_" + GetOSP_Num().ToString().PadLeft(2, '0') + "_2.dbf";

            //Int64 cnt = WritePotdToDBF(true, fullpath, tablename);
            //    //WritePensToDBF(false, fullpath, tablename);
            //lblPotdDbfValue.Text = cnt.ToString();

            //if (File.Exists(string.Format(@"{0}\{1}", fullpath, tablename_n)))
            //{
            //    DialogResult rv = MessageBox.Show("�� ���� " + string.Format(@"{0}\{1}", fullpath, tablename_n) + ", ��������� � ���������������� �����, ���������� ����. ������� ���?", "��������", MessageBoxButtons.YesNo);
            //    if (rv == DialogResult.Yes)
            //    {
            //        File.Delete(string.Format(@"{0}\{1}", fullpath, tablename_n));
            //        File.Move(string.Format(@"{0}\{1}", fullpath, tablename), string.Format(@"{0}\{1}", fullpath, tablename_n));
            //    }
            //}
            //else
            //{
            //    File.Move(string.Format(@"{0}\{1}", fullpath, tablename), string.Format(@"{0}\{1}", fullpath, tablename_n));
            //}
        }

        private void b_findgibd_Click(object sender, EventArgs e)
        {
            /*
            if (bReadFromCopy)
            {
                con = new OleDbConnection(constr2);
            }*/
            /***������� ���� ������������ �� �� ������� �����***/
            //DT_gibd_reg = GetDataTableFromFB("select b.num_id, a.date_ip_in, a.name_v, a.sisp_ttl from ip a, id b where a.pk_id = b.pk and a.sisp_key like '%/1/22/37%' and a.date_ip_out is null ", "TOFIND");
            
            /*
            if(cb_agibd.Checked == true)
                DT_gibd_reg = GetDataTableFromFB("select distinct b.num_id, c.nomid, a.date_ip_in, a.name_v, a.sisp_ttl, a.num_ip, a.name_d, a.adr_d, a.uscode, a.pk fk_ip, b.pk fk_id, c.summ, c.nomid from ip a, id b, gibdd c where a.pk_id = b.pk and b.num_id = c.nomid and a.sisp_key like '%/1/22/37%' and a.date_ip_out is null order by a.uscode", "TOFIND");
            else
                DT_gibd_reg = GetDataTableFromFB("select distinct b.num_id, c.nomid, a.date_ip_in, a.name_v, a.sisp_ttl, a.num_ip, a.name_d, a.adr_d, a.uscode, a.pk fk_ip, b.pk fk_id, c.summ, c.nomid  from ip a, id b, gibdd c where a.pk_id = b.pk and b.num_id = c.nomid and a.sisp_key like '%/1/22/37%' and a.date_ip_out is null and (a.date_ip_in >= '" + DatZapr1_gibd.ToShortDateString() + "' AND a.date_ip_in <= '" + DatZapr2_gibd.ToShortDateString() + "') order by a.uscode", "TOFIND");
            */

            /*
            if (bReadFromCopy)
            {
                con = new OleDbConnection(constr1);
            }*/

            b_zaprgibd.Enabled = true;
        }

        public int idCount(DataTable DT_work, int nomspi) 
        {
            int summ=0;
            DataRow[] dt;
            dt = DT_work.Select("uscode =" + nomspi.ToString());
            summ = dt.Length;
            return summ;
        }


        private string Money_ToStr(decimal nMoney)
        {
            string txtResult = "";
            txtResult = nMoney.ToString("N2").Replace(".", " ���. ");
            txtResult = txtResult.Replace(",", " ���. ") + " ���.";

            return txtResult;
        }

        private string Money_ToStr(double nMoney)
        {
            string txtResult = "";
            txtResult = nMoney.ToString("N2").Replace(".", " ���. ");
            txtResult = txtResult.Replace(",", " ���. ") + " ���.";

            return txtResult;
        }




        private void b_loadgibd_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "DBF �����(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();
            OleDbConnection conGIBDD;

            if (res == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {
                    
                    //ChangeByte(openFileDialog1.FileName, 0x65, 30);
                    string tablename = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - 4);
                    tablename = tablename.Substring(tablename.LastIndexOf("\\") + 1);

                    // ��������� �� ����� ����� ���� ���. ��������� (�������)
                    string txtDateIsh = tablename.Substring(6, 2) + '.' + tablename.Substring(4, 2) + '.' + tablename.Substring(0, 4);

                    DateTime dtDateIsh;
                    if (!DateTime.TryParse(txtDateIsh, out dtDateIsh)) {
                        dtDateIsh = DateTime.MinValue;
                    }

                    // ��������� �� ����� ����� ��� ����� ���������
                    string txtIshNumber = tablename.Substring(9, tablename.Length - 9);

                    try
                    {
                        DataSet ds = new DataSet();
                        DataTable tbl = ds.Tables.Add(tablename);
                        DBFcon = new OleDbConnection();
                        DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=RUSSIAN");
                        DBFcon.Open();
                        m_cmd = new OleDbCommand();
                        m_cmd.Connection = DBFcon;
                        m_cmd.CommandText = "SELECT * FROM " + tablename;
                        using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                        {
                            ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                            rdr.Close();
                        }

                        DBFcon.Close();

                        

                        conGIBDD = new OleDbConnection(constrGIBDD);

                        // ��������� �� ���� �� � ������ ��� ����������� ������ � ��������� ������� txtIshNumber
                        // �������� ������ ���� ����������� ��������
                        ArrayList alLoadedReestrs = GetLoadedReestrs(conGIBDD);

                        if (alLoadedReestrs.Contains(txtIshNumber))
                        {
                            Exception ex = new Exception("������. ������������� ������� �������� ��������� ������ " + txtIshNumber + ".");
                            throw ex;
                        }
                        
                        if ((conGIBDD == null) || (conGIBDD.State == ConnectionState.Closed))
                            conGIBDD.Open();
                        

                        l_loadgibd.Text = tbl.Rows.Count.ToString();
                        prbWritingDBF.Value = 0;
                        prbWritingDBF.Maximum = tbl.Rows.Count;
                        prbWritingDBF.Step = 1;

                        Int32 iCnt = 0;
                        OleDbTransaction tran;

                        tran = conGIBDD.BeginTransaction(IsolationLevel.ReadCommitted);
                        
                        foreach (DataRow row in tbl.Rows)
                        {
                            
                            m_cmd = new OleDbCommand();
                            m_cmd.Connection = conGIBDD;
                            m_cmd.Transaction = tran;

                            // ���� ������� ����������� � �.�. � �������� � ���� ��������
                            // ���� - ������ ����� ���������
                            // �������������� ���� - ������ ������ �������
                            // ��������� �����, ����, �����, �������, ���, ���� ������, ����� �������, ���� �������

                            //m_cmd.CommandText = "INSERT INTO GIBDD (NOMID, PK)";
                            //m_cmd.CommandText += " VALUES(:NOMID, GEN_ID(GEN_GIBDD_ID, 1))";

                            //m_cmd.CommandText = "INSERT INTO GIBDD_PLATEZH (NUMBER, NOMID, DATID, SUMM, SUMM_DOC, FIO_D, DATE_DOC, ISH_NUMBER, DATE_ISH, FL_USE)";
                            //m_cmd.CommandText += " VALUES(:NUMBER, :NOMID, :DATID, :SUMM, :SUMM_DOC, :FIO_D, :DATE_DOC, :ISH_NUMBER, :DATE_ISH, :FL_USE)";

                            m_cmd.CommandText = "INSERT INTO GIBDD_PLATEZH (NUMBER, NOMID, DATID, SUMM, SUMM_DOC, FIO_D, DATE_DOC, ISH_NUMBER, DATE_ISH, FL_USE, NUM_DOC, BORN_D, DATE_VH)";
                            m_cmd.CommandText += " VALUES(:NUMBER, :NOMID, :DATID, :SUMM, :SUMM_DOC, :FIO_D, :DATE_DOC, :ISH_NUMBER, :DATE_ISH, :FL_USE, :NUM_DOC, :BORN_D, :DATE_VH)";
                            
                            


                            string txtNumber = Convert.ToString(row["Number"]).TrimEnd();
                            string txtGibddIDNumber = txtNumber.Substring(1, 7);
                            
                            string txtDatID = Convert.ToString(row["Date_exec"]);
                            DateTime dtDatID;
                            
                            Double nSum;
                            string txtSum = Convert.ToString(row["Summa"]);

                            string txtFioD = Convert.ToString(row["Plat_name"]).TrimEnd();
                            string txtDateDoc = Convert.ToString(row["Date_doc"]);
                            DateTime dtDateDoc;

                            string txtNumDoc = Convert.ToString(row["Num_doc"]).TrimEnd();

                            string txtBornD = Convert.ToString(row["Date_plat"]);
                            DateTime dtBornD;

                            if (!DateTime.TryParse(txtBornD, out dtBornD))
                            {
                                dtBornD = Convert.ToDateTime("01.01.1800");
                            }

                            Double nSumDoc;
                            string txtSumDoc = Convert.ToString(row["Summa_doc"]);
                            
                            if (txtGibddIDNumber[0] == '0')
                            {
                                txtGibddIDNumber = txtGibddIDNumber.Substring(1, 6);
                            }

                            if (!DateTime.TryParse(txtDatID, out dtDatID))
                            {
                                dtDatID = DateTime.MinValue;
                            }

                            if(!Double.TryParse(txtSum, out nSum)){
                                nSum = -1;
                            }

                            if (!DateTime.TryParse(txtDateDoc, out dtDateDoc))
                            {
                                dtDateDoc = DateTime.MinValue;
                            }
                            
                            if (!Double.TryParse(txtSumDoc, out nSumDoc))
                            {
                                nSumDoc = -1;
                            }

                            m_cmd.Parameters.Add(new OleDbParameter(":NUMBER", txtNumber));
                            m_cmd.Parameters.Add(new OleDbParameter(":NOMID", txtGibddIDNumber));
                            m_cmd.Parameters.Add(new OleDbParameter(":DATID", dtDatID)); // ���� ��
                            m_cmd.Parameters.Add(new OleDbParameter(":SUMM", nSum));     // ����� ��
                            m_cmd.Parameters.Add(new OleDbParameter(":SUMM_DOC", nSumDoc));   //����� ������
                            m_cmd.Parameters.Add(new OleDbParameter(":FIO_D", txtFioD));
                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_DOC", dtDateDoc)); // ���� ������
                            

                            m_cmd.Parameters.Add(new OleDbParameter(":ISH_NUMBER", txtIshNumber));   // ��� �����
                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_ISH", dtDateIsh));   // ���� ��� 
                            m_cmd.Parameters.Add(new OleDbParameter(":FL_USE", Convert.ToInt32(0)));   // ���� - �����/�������
                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_DOC", txtNumDoc));   // ����� ���������
                            m_cmd.Parameters.Add(new OleDbParameter(":BORN_D", dtBornD));   // ���� �������� ��������
                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_VH", DateTime.Today));   // ���� �������� �������
                            
                            
                            
                            int result = m_cmd.ExecuteNonQuery();

                            if (result != -1)
                            {
                                iCnt++;
                                prbWritingDBF.PerformStep();
                            }
                        }
                        tran.Commit();
                        conGIBDD.Close();
                        MessageBox.Show("������ ������� ���������.", "���������", MessageBoxButtons.OK);
   
                    }
                    catch (OleDbException ole_ex)
                    {
                        foreach (OleDbError err in ole_ex.Errors)
                        {
                            MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                        }
                        //return false;
                    }
                    catch (Exception ex)
                    {
                        //if (DBFcon != null) DBFcon.Close();
                        MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                        //return false;
                    }
                    //return true;

                }
            }
        }
            
        private void button11_Click(object sender, EventArgs e)
        {
            //DateTime dat1, dat2, start, end;
            //dat1 = Convert.ToDateTime("14.05.2008");
            //dat2 = Convert.ToDateTime("29.05.2008");
            //start = DateTime.Now;
            //DT_pens_reg = GetDataTableFromFB("SELECT DISTINCT a.USCODE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD, a.name_d as FIOVK, a.DATE_BORN_D as DATROZHD, a.ADR_D as ADDR, a.PK as FK_IP, a.PK_ID as FK_ID FROM IP a LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.DATE_IP_OUT is null and b.KOD != 1006 and (a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' AND a.NUM_IP NOT LIKE '%!%'", "TOFIND",IsolationLevel.ReadCommitted);
            //DT_pens_doc = GetDataTableFromFB("SELECT DISTINCT a.USCODE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD, a.name_d as FIOVK, a.DATE_BORN_D as DATROZHD, a.ADR_D as ADDR, b.PK as FK_DOC, a.PK as FK_IP, a.PK_ID as FK_ID FROM IP a LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.DATE_IP_OUT is null and b.KOD = 1006 and (b.DATE_DOC >= '" + dat1.ToShortDateString() + "' AND b.DATE_DOC <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' AND a.NUM_IP NOT LIKE '%!%'", "TOFIND", IsolationLevel.ReadCommitted);
            //end = DateTime.Now;
            //lblTime1.Text = Convert.ToString(((TimeSpan)(end - start)).TotalMilliseconds);


            //start = DateTime.Now;
            //DT_reg = GetDataTableFromFB("SELECT DISTINCT a.NAME_D as FIOVK, a.NUM_IP as ZAPROS, a.VIDD_KEY as LITZDOLG, a.DATE_BORN_D as GOD, a.USCODE as NOMSPI, a.NUM_IP as NOMIP, a.SUM_ as SUMMA, a.WHY as VIDVZISK, a.INND as INNORG, a.DATE_IP_IN as DATZAPR, a.ADR_D as ADDR,a.TEXT_PP as OSNOKON, a.PK as FK_IP, a.PK_ID as FK_ID FROM IP a WHERE a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "' AND NUM_IP NOT LIKE '%��' AND a.NUM_IP NOT LIKE '%��' AND NUM_IP NOT LIKE '%!%' ORDER BY a.NAME_D", "TOFIND", IsolationLevel.ReadCommitted);
            //DT_okon = GetDataTableFromFB("SELECT DISTINCT a.NAME_D as FIOVK, a.NUM_IP as ZAPROS, a.VIDD_KEY as LITZDOLG, a.DATE_BORN_D as GOD, a.USCODE as NOMSPI, a.NUM_IP as NOMIP, a.SUM_ as SUMMA, a.WHY as VIDVZISK, a.INND as INNORG, a.DATE_IP_OUT as DATZAPR, a.ADR_D as ADDR,a.TEXT_PP as OSNOKON FROM IP a WHERE a.DATE_IP_OUT is not null AND NUM_IP NOT LIKE '%!%' and DATE_IP_OUT <= '" + dat2.ToShortDateString() + "' and DATE_IP_OUT >= '" + dat1.ToShortDateString() + "'", "TOFIND", IsolationLevel.ReadCommitted);
            //end = DateTime.Now;
            //lblTime2.Text = Convert.ToString(((TimeSpan)(end - start)).TotalMilliseconds);

            //start = DateTime.Now;
            //DT_ktfoms_reg = GetDataTableFromFB(" SELECT DISTINCT a.USCODE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD, a.name_d as FIOVK, a.DATE_BORN_D as DATROZHD, a.ADR_D as ADDR, a.PK as FK_IP, a.PK_ID as FK_ID FROM IP a LEFT JOIN DOCUMENT b ON b.FK = a.PK WHERE (a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' AND a.NUM_IP NOT LIKE '%��' AND a.NUM_IP NOT LIKE '%��' AND a.NUM_IP NOT LIKE '%!%'", "TOFIND", IsolationLevel.ReadCommitted);
            //DT_ktfoms_doc = GetDataTableFromFB("SELECT DISTINCT a.USCODE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD, a.name_d as FIOVK, a.DATE_BORN_D as DATROZHD, a.ADR_D as ADDR, b.PK as FK_DOC, a.PK as FK_IP, a.PK_ID as FK_ID FROM IP a LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.DATE_IP_OUT is null and b.KOD = 1010 and (b.DATE_DOC >= '" + dat1.ToShortDateString() + "' AND b.DATE_DOC <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' AND a.NUM_IP NOT LIKE '%!%'", "TOFIND", IsolationLevel.ReadCommitted);
            //end = DateTime.Now;
            //lblTime3.Text = Convert.ToString(((TimeSpan)(end - start)).TotalMilliseconds);
            //OdbcConnection MyodbcConn = new OdbcConnection(@"Driver={Microsoft dBase Driver (*.dbf)};CollatingSequence=ASCII;PageTimeout=5;UserCommitSync=Yes;MaxScanRows=8;DefaultDir=c:\\basebank\\cred_org;Deleted=1;Statistics=0;FIL=dBase IV;UID=admin;MaxBufferSize=2048;Threads=3;SafeTransactions=0");
            OdbcConnection MyodbcConn = new OdbcConnection(@"Driver={Microsoft dBase Driver (*.dbf)};collatingsequence=ASCII;defaultdir=C:\basebank\cred_org;deleted=1;driverid=277;fil=dBase IV;filedsn=C:\basebank\cred_org\tofind.dbf.dsn;maxbuffersize=2048;maxscanrows=8;pagetimeout=600;safetransactions=0;statistics=0;threads=3;uid=admin;usercommitsync=Yes");
            
            DataSet myDS = new DataSet();
            
            DataTable myTbl = myDS.Tables.Add("Tofind");

            //string txtSql = "Select NOMOSP, LITZDOLG, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DatZapr1, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FLOKON, OSNOKON from TOFIND"; 
            string txtSql = "SELECT * FROM TOFIND"; 

            try
            {
                MyodbcConn.Open();
                //OdbcTransaction odbcTran = MyodbcConn.BeginTransaction(IsolationLevel.RepeatableRead);
                //OdbcCommand odbcCmd = new OdbcCommand(txtSql, MyodbcConn, odbcTran);
                OdbcCommand odbcCmd = new OdbcCommand(txtSql, MyodbcConn);

                using (OdbcDataReader odbcRdr = odbcCmd.ExecuteReader(CommandBehavior.Default))
                {
                    myDS.Load(odbcRdr, LoadOption.OverwriteChanges, myTbl);
                }

                //odbcTran.Rollback();
                MyodbcConn.Close();

            }
            catch (OdbcException odbc_ex)
            {
                foreach (OdbcError odbc_err in odbc_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + odbc_err.Message + "Native Error: " + odbc_err.NativeError + "Source: " + odbc_err.Source + "SQL State   : " + odbc_err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }

            MessageBox.Show("��������� �����: " + myTbl.Rows.Count.ToString(), "��������!", MessageBoxButtons.OK);
        }

    public class ComboItem
    {
        public ComboItem(decimal id, string name)
        {
            this.id = id;
            this.name = name;
        }
        private decimal id;
        public decimal Id
        {
            get
            {
                return this.id;
            }
            set
            {
                this.id = value;
            }
        }

        private string name;
        public string Name
        {
            get
            {
                return this.name;
            }
            set
            {
                this.name = value;
            }
        }
        }

        
        private void tabPotd_Click(object sender, EventArgs e)
        {

        }

        private void b_Load_basegibd_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "DBF �����(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {

                    string tablename = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - 4);
                    tablename = tablename.Substring(tablename.LastIndexOf("\\") + 1);

                    try
                    {
                        DataSet ds = new DataSet();
                        DataTable tbl = ds.Tables.Add(tablename);
                        DBFcon = new OleDbConnection();
                        DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=RUSSIAN");
                        DBFcon.Open();
                        m_cmd = new OleDbCommand();
                        m_cmd.Connection = DBFcon;
                        m_cmd.CommandText = "SELECT * FROM " + tablename;
                        using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                        {
                            ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                            rdr.Close();
                        }

                        DBFcon.Close();
                        l_loadgibd.Text = tbl.Rows.Count.ToString();

                        Int32 iCnt = 0;
                        OleDbTransaction tran;

                        prbWritingDBF.Value = 0;
                        prbWritingDBF.Maximum = tbl.Rows.Count;
                        prbWritingDBF.Step = 1;

                        if (con != null && con.State != ConnectionState.Closed) con.Close();
                        con.Open();
                        tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                        //int Gibid = 0;
                        //string Gibstr = "";
                        //string Gibstr2 = "";

                        foreach (DataRow row in tbl.Rows)
                        {
                            //if(Convert.ToString(row["NOMID"]).Trim()=="106199")
                            //    Gibid = 1;                           
                            m_cmd = new OleDbCommand();
                            m_cmd.Connection = con;
                            m_cmd.Transaction = tran;

                            m_cmd.CommandText = "INSERT INTO GIBDD (NOMID, SUMM)";
                            m_cmd.CommandText += " VALUES(:NOMID,:SUMM)";
                            m_cmd.Parameters.Add(new OleDbParameter(":NOMID", Convert.ToString(row["NOMID"])));
                            m_cmd.Parameters.Add(new OleDbParameter(":SUMM", Convert.ToDecimal(row["SUMVZISK"])));

                            int result = m_cmd.ExecuteNonQuery();
                            if (result != -1)
                            {
                                iCnt++;
                                prbWritingDBF.PerformStep();
                            }
                        }
                        tran.Commit();
                        con.Close();
                        MessageBox.Show("���������� �������: " + iCnt.ToString() + "\n. ������ ����� ����������� ������ �������.", "���������", MessageBoxButtons.OK);

                    }
                    catch (OleDbException ole_ex)
                    {
                        foreach (OleDbError err in ole_ex.Errors)
                        {
                            MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                        }
                        //return false;
                    }
                    catch (Exception ex)
                    {
                        //if (DBFcon != null) DBFcon.Close();
                        MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                        //return false;
                    }
                    //return true;

                }
            }
        }

        private void b_delgibdd_Click(object sender, EventArgs e)
        {
            /***�������� �������� ��� ��� �� �� ������� ��������� ������� � �����***/
            //int vid_d = 1;
            //DateTime dtDate;
            DateTime DatZapr = DateTime.Today;
            String uid = System.Guid.NewGuid().ToString();
            String osp_name = GetOSP_Name();
            uid = cutEnd(uid, 30);
            Int32 iCnt = 0;
            OleDbTransaction tran;

            prbWritingDBF.Value = 0;
            prbWritingDBF.Maximum = DT_gibd_reg.Rows.Count;
            prbWritingDBF.Step = 1;


            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);


                foreach (DataRow row in DT_gibd_reg.Rows)
                {
                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = con;
                    m_cmd.Transaction = tran;

                    m_cmd.CommandText = "DELETE FROM GIBDD WHERE NOMID LIKE '" + Convert.ToString(row["NOMID"]) + "'";

                    if (m_cmd.ExecuteNonQuery() != -1)
                    {
                        iCnt++;
                        prbWritingDBF.PerformStep();
                    }
                    else
                    {
                        ;
                    }
                }

                tran.Commit();

                con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    //tran.Rollback();
                    con.Close();
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            /***�������� �������� ��� ��� �� �� ������� ��������� ������� � �����***/
            int vid_d = 1;
            DateTime dtDate;
            DateTime DatZapr = DateTime.Today;
            String uid = System.Guid.NewGuid().ToString();
            String osp_name = GetOSP_Name();
            uid = cutEnd(uid, 30);
            Int32 iCnt = 0;
            OleDbTransaction tran;

            //DatZapr = Convert.ToDateTime(lb_gibd.SelectedItem);
            try
            {
                DT_gibd_rst = GetDataTableFromFB("select distinct c.nomid, c.summ, c.fio_d, c.base_t, a.num_ip, c.date_z, a.uscode, a.name_d, a.name_v, a.adr_d, a.pk fk_ip, b.pk fk_id, b.num_id from ip a, id b, gibdd c where a.pk_id = b.pk and b.num_id = c.nomid and a.sisp_key like '%/1/22/37%' and a.date_ip_out is null and c.fl_use = 1 and date_use = '" + DatZapr.ToShortDateString() + "' and b.d_id = c.dateid order by a.uscode", "TOFIND");
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }


            prbWritingDBF.Value = 0;
            //prbWritingDBF.Maximum = DT_gibd_reg.Rows.Count;
            prbWritingDBF.Maximum = DT_gibd_rst.Rows.Count;
            prbWritingDBF.Step = 1;


            try
            {
                string txtGibdName = "����� �����";
                string txtGibdConv = "����������� �� ������ �����������";

                DateTime dat1 = DateTime.Today;
                DateTime dat2 = DateTime.Today;

                dat1 = DatZapr1_gibd.Date;
                dat2 = DatZapr2_gibd.Date;

                if (OooIsInstall)
                {
                    //OOo start
                    OOo_Writer OOo_cld = new OOo_Writer();
                    OOo_cld.OOo_Gbdd(DatZapr, DT_gibd_rst, con, this);
                }
                else
                {

                    Microsoft.Office.Interop.Word.ApplicationClass app = new ApplicationClass();

                    object s1 = "";
                    object fl = false;
                    object t = WdNewDocumentType.wdNewBlankDocument;
                    object fl2 = true;

                    Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref s1, ref fl, ref t, ref fl2);
                    doc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                    Paragraph par = doc.Content.Paragraphs[1];

                    par.Range.Font.Name = "Courier";
                    par.Range.Font.Size = 8;
                    float a = par.Range.PageSetup.RightMargin;
                    float b = par.Range.PageSetup.LeftMargin;
                    float c = par.Range.PageSetup.TopMargin;
                    float d = par.Range.PageSetup.BottomMargin;

                    par.Range.PageSetup.RightMargin = 30;
                    par.Range.PageSetup.LeftMargin = 30;
                    par.Range.PageSetup.TopMargin = 20;
                    par.Range.PageSetup.BottomMargin = 20;

                    //par.Range.Text = "������!";
                    //for (int i = 1; i < 200; i++)
                    //par.Range.Text += Convert.ToString(i - 1);

                    int spi;
                    int sch_line = 1;
                    int fl_fst = 1;
                    int idcnt;
                    int totline = 0;

                    spi = 999;

                    string nline = "";

                    prbWritingDBF.Value = 0;
                    //prbWritingDBF.Maximum = DT_gibd_reg.Rows.Count;
                    prbWritingDBF.Maximum = DT_gibd_rst.Rows.Count;
                    prbWritingDBF.Step = 1;

                    //foreach (DataRow row in DT_gibd_reg.Rows)
                    foreach (DataRow row in DT_gibd_rst.Rows)
                    {
                        if (spi < Convert.ToInt32(row["USCODE"]))
                        {
                            if (cb_prgibd.Checked == true)
                            {
                                while (sch_line > 61)
                                    sch_line = sch_line - 61;

                                nline = "";
                                for (int i = sch_line; i < 61; i++)
                                    par.Range.Text += "";
                                //nline += "\n";
                                //par.Range.Text += Convert.ToString(i - 1);
                                //par.Range.Text += nline;
                            }
                            else
                            {
                                string bord = "";
                                for (int j = 0; j < 100; j++)
                                    bord += "*";
                                par.Range.Text += bord;
                                par.Range.Text += "";
                                sch_line++;
                                sch_line++;

                            }

                            spi = 999;
                        }
                        if (spi > Convert.ToInt32(row["USCODE"]))
                        {
                            if (cb_prgibd.Checked == false)
                            {
                                totline = sch_line;
                                while (totline > 61)
                                    totline = totline - 61;
                                idcnt = idCount(DT_gibd_rst, Convert.ToInt32(row["USCODE"]));
                                if ((idcnt + 11 + totline) > 61)
                                {
                                    for (int i = (totline); i < 61; i++)
                                    {
                                        //par.Range.Text += Convert.ToString(i - 1);
                                        par.Range.Text += "";
                                        sch_line++;
                                    }
                                }
                            }

                            par.Range.Text += "������ �������������� ����������, ���� �� ������� �������.";
                            par.Range.Text += "����������� �� ������ ������, ���������� �� �����.\n";
                            par.Range.Text += "���� ������������ " + DatZapr.ToShortDateString() + "\n";

                            spi = Convert.ToInt32(row["USCODE"]);

                            if (cb_prgibd.Checked == true)
                            {
                                sch_line = 0;
                                if (fl_fst == 1)
                                {
                                    sch_line = 1;
                                    fl_fst = 0;
                                }
                            }
                            //par.Range.Text += " ";
                            par.Range.Text += GetSpiName2(Convert.ToInt32(row["USCODE"])) + "\n";
                            par.Range.Text += "����� ��       �������     ����������     ����� ��       ���� �������� � ���� �����      ���� ��������";
                            //par.Range.Text += GetOSP_Name();
                            sch_line += 8;
                        }
                        if (spi == Convert.ToInt32(row["USCODE"]))
                        {
                            // ������ �����-�� svn ������!
                            string txtResponse = Convert.ToString(row["NOMID"]) /*+ "  " + Money_ToStr(Convert.ToDecimal(row["summ"]))*/ + "  " + Convert.ToString(row["FIO_D"] + "  " + Convert.ToString(row["name_v"]) + "  " + Convert.ToString(row["NUM_IP"])) + "  " + Convert.ToString(Convert.ToDateTime(row["BASE_T"]).ToShortDateString()) + "  " + Convert.ToString(Convert.ToDateTime(row["DATE_Z"]).ToShortDateString());
                            //sch_line++;
                            //string txtResponse = Convert.ToString(row["BASE_T"]) + " " + Convert.ToString(row["DATE_Z"]);
                            par.Range.Text += txtResponse;
                            sch_line++;
                            if (txtResponse.Length > 200)
                            {
                                sch_line++; // ���� ��� ������� ������
                            }
                        }

                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }



                    app.Visible = true;
                    //*************************************************
                }
                //con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    //tran.Rollback();
                    con.Close();
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void panel10_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnLoadKrcData_Click(object sender, EventArgs e)
        {
            btnWriteKrcDBF.Enabled = false;
            btnWriteKrcZapros.Enabled = false;

            if (bReadFromCopy)
            {
                con = new OleDbConnection(constr2);
            }

            ReadKRCData(DatZapr1_krc, DatZapr2_krc);

            if (bReadFromCopy)
            {
                con = new OleDbConnection(constr1);
            }

            btnWriteKrcDBF.Enabled = true;
            btnWriteKrcZapros.Enabled = true;

            lblReadKrcRowsValue.Text = DT_krc_reg.Rows.Count.ToString();
            
        }

        private void btnWriteKrcZapros_Click(object sender, EventArgs e)
        {
            int zapros = InsertKRCZapros(DatZapr1_krc, DatZapr2_krc);
            lblKrcZaprosValue.Text = Convert.ToString(zapros);
        }

        private void btnWriteKrcDBF_Click(object sender, EventArgs e)
        {
            Int64 cnt;
            if (bDateFolderAdd)
            {
                CreatePathWithDate(krc_path);
            }
            else
            {
                FolderExist(krc_path);
            }

            string tablename = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString().PadLeft(2, '0') + DateTime.Today.Day.ToString().PadLeft(2, '0') + "_OSP_" + GetOSP_Num().ToString().PadLeft(2, '0') + ".dbf";
            cnt = WriteKrcToDBF(true, fullpath, tablename);
            lblKrcDbfValue.Text = cnt.ToString();
        }

        private void b_anskrc_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "DBF �����(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {
                    ChangeByte(openFileDialog1.FileName, 0x65, 30);
                    if (openFileDialog1.FileName.ToLower().Contains("osp"))
                    {
                        string tablename = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - 4);
                        tablename = tablename.Substring(tablename.LastIndexOf("\\") + 1);

                        try
                        {
                            DataSet ds = new DataSet();
                            DataTable tbl = ds.Tables.Add(tablename);
                            DBFcon = new OleDbConnection();
                            DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=RUSSIAN");
                            DBFcon.Open();
                            m_cmd = new OleDbCommand();
                            m_cmd.Connection = DBFcon;
                            m_cmd.CommandText = "SELECT * FROM " + tablename;
                            using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                            {
                                ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                                rdr.Close();
                            }

                            DBFcon.Close();

                            Int32 iCnt = 0;
                            OleDbTransaction tran;

                            prbWritingDBF.Value = 0;
                            prbWritingDBF.Maximum = tbl.Rows.Count;
                            prbWritingDBF.Step = 1;

                            DateTime dat1 = DateTime.Today;
                            DateTime dat2 = DateTime.Today;

                            dat1 = DatZapr1_krc.Date;
                            dat2 = DatZapr2_krc.Date;

                            string txtKrcName = cutEnd((GetLegal_Name(krc_id)).Trim(), 200);
                            String uid = System.Guid.NewGuid().ToString();
                            String osp_name = GetOSP_Name();
                            int vid_d = 1;
                            /*
                            if (tbl.Rows.Count > 0)
                            {
                                if (!(DateTime.TryParse(Convert.ToString(tbl.Rows[0]["DATZAPR1"]), out dat1)))
                                {
                                    dat1 = DateTime.Today;
                                }

                                if (!(DateTime.TryParse(Convert.ToString(tbl.Rows[0]["DATZAPR2"]), out dat2)))
                                {
                                    dat2 = DateTime.Today;
                                }
                            }
                            */
                            if (con != null && con.State != ConnectionState.Closed) con.Close();
                            con.Open();
                            tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                            #region "krc_reg"

                            //foreach (DataRow row in tbl.Rows)
                            //{
                            //    m_cmd = new OleDbCommand();
                            //    m_cmd.Connection = con;
                            //    m_cmd.Transaction = tran;

                            //    m_cmd.CommandText = "INSERT INTO ZAPROS (PK, NUM_ZAPR, SUB_DIV, FIO_SPI, NUM_IP, DATE_ZAPR, VID_D, INN_D, NAME_D, DATE_R, NUM_RES, DATE_RES, RESULT, FK_DOC, FK_IP, FK_ID, NUM_ID, TEXT, DATE_BEG, DATE_END, ADRESS, NUM_PACK, NUM_ZAPR_IN_PACK, STATUS, DATE_SEND, TEXT_ERROR, USCODE, CONVENTION, WHYRESPONS, WHYPREPARE, PASSPORT, SUMM, TEMP, FK_LEGAL, ADRESAT)";
                            //    m_cmd.CommandText += " VALUES(GEN_ID(S_N_ZAPROS, 1), :NUM_ZAPR, :SUB_DIV, :FIO_SPI, :NUM_IP, :DATE_ZAPR, :VID_D, :INN_D, :NAME_D, :DATE_R, :NUM_RES, :DATE_RES, :RESULT, :FK_DOC, :FK_IP, :FK_ID, :NUM_ID, :TEXT, :DATE_BEG, :DATE_END, :ADRESS, :NUM_PACK, :NUM_ZAPR_IN_PACK, :STATUS, :DATE_SEND, :TEXT_ERROR, :USCODE, :CONVENTION, :WHYRESPONS, :WHYPREPARE, :PASSPORT, :SUMM, :TEMP, :FK_LEGAL, :ADRESAT)";


                            //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR", Convert.ToString(row["ZAPROS"]).Trim()));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":SUB_DIV", cutEnd(osp_name, 100).Trim()));

                            //    OleDbCommand spi_name_cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE USCODE = '" + Convert.ToString(row["NOMSPI"]) + "'", con, tran);
                            //    String spi_name = Convert.ToString(spi_name_cmd.ExecuteScalar());

                            //    m_cmd.Parameters.Add(new OleDbParameter(":FIO_SPI", cutEnd(spi_name, 100).Trim()));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_IP", cutEnd(Convert.ToString(row["ZAPROS"]).Trim(), 40)));


                            //    //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// ������� ��� ���� ����������� ��
                            //    //{
                            //    //    DatZapr = DateTime.Today;
                            //    //}
                            //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                            //    // � ��� ��� ���� ������ �������������

                            //    vid_d = 1; // ���. ����
                            //    //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                            //    //{
                            //    //    vid_d = 1;// ���. ����
                            //    //}

                            //    m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                            //    //m_cmd.Parameters.Add(new OleDbParameter(":INN_D", Convert.ToString(row["INNORG"])));
                            //    m_cmd.Parameters.Add(new OleDbParameter(":INN_D", ""));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", cutEnd(Convert.ToString(row["NAMEDOL"]).Trim(), 100)));

                            //    /*
                            //    if (!DateTime.TryParse(Convert.ToString(row["DATROZHD"]), out dtDate))
                            //    {
                            //        dtDate = DateTime.MaxValue;
                            //    }
                            //    */
                            //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", DateTime.MaxValue));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// ��� ����� ������ �� ������� ���-��

                            //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DateTime.Today));// ���� ������, �� ������ �������� ��� ������

                            //    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(1)));// (0 - ������� �� ���������������, 1 - ��� ���. �� ��������, ������ 1 - ���� ���-� �� ��������) (��� ��� � FIND - ��� ������ 1, ���������� 0)

                            //    Int32 iKey = -1;
                            //    //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                            //    //{
                            //    //    iKey = 0;
                            //    //}

                            //    // TODO: �������, ��� ��� ������� ������ �� DOCUMENTS!!!
                            //    // ���� ��� � ������� DOCUMENTS ��������� ������, ���� �� �����
                            //    m_cmd.Parameters.Add(new OleDbParameter(":FK_DOC", iKey));

                            //    iKey = 0;
                            //    m_cmd.Parameters.Add(new OleDbParameter(":FK_IP", iKey));
                            //    m_cmd.Parameters.Add(new OleDbParameter(":FK_ID", iKey));


                            //    //OleDbCommand num_id_cmd = new OleDbCommand("Select NUM_ID from ID where ID.PK = " + Convert.ToString(row["FK_ID"]), con, tran);
                            //    //String num_id = Convert.ToString(num_id_cmd.ExecuteScalar());

                            //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ID", Convert.ToString(row["NOMID"])));

                            //    if (Convert.ToInt32(row["SUMPL"])!=0)
                            //        m_cmd.Parameters.Add(new OleDbParameter(":TEXT", "������� �������� ����� �������������: " + Convert.ToString(row["SUMPL"]) + " . ���� �������: " + Convert.ToString(row["DATPL"])));// ����� ������ - ������
                            //    else
                            //        m_cmd.Parameters.Add(new OleDbParameter(":TEXT", "�� ������ � " + Convert.ToString(row["DATZAPR1"]) + " �� " + Convert.ToString(row["DATZAPR1"]) + " �������� �� ���������"));// ����� ������ - ������

                            //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", DatZapr1_krc));
                            //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", DatZapr2_krc));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADRES"]).Trim(), 250)));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                            //    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "�����"));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", ""));

                            //    if (!(Int32.TryParse(Convert.ToString(row["NOMSPI"]), out iKey)))
                            //    {
                            //        iKey = 0;
                            //    }
                            //    m_cmd.Parameters.Add(new OleDbParameter(":USCODE", iKey));

                            //    //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));


                            //    m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", ""));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":WHYRESPONS", ""));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":WHYPREPARE", ""));

                            //    // ����� � ������� persons � ��� ���� ������ ������
                            //    // ������� persons �� ip.fk; � ������� physical �� person.tablename + person.FK
                            //    // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                            //    //

                            //    m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// ����� ������ ����. ������

                            //    Double sum = 0;
                            //    //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                            //    //{
                            //    //    sum = 0;
                            //    //}
                            //    m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                            //    // TODO: ����� ��� ��� NOMIP � tofind
                            //    String txtNOMIP = Convert.ToString(row["ZAPROS"]).Trim();
                            //    if (txtNOMIP.Trim() != "")
                            //    {
                            //        String[] strings = txtNOMIP.Split('/');
                            //        if (!(Int32.TryParse(Convert.ToString(strings[2]), out iKey)))
                            //        {
                            //            iKey = 0;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        iKey = 0;
                            //    }
                            //    m_cmd.Parameters.Add(new OleDbParameter(":TEMP", iKey));

                            //    //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));
                            //    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                            //    //OleDbCommand legal_id_cmd = new OleDbCommand("Select PK from LEGAL where UPPER(FULL_NAME) like UPPER('"+ Convert.ToString(row["NAME_V"]) +"')", con, tran);
                            //    //String legal_id = Convert.ToString(legal_id_cmd.ExecuteScalar());

                            //    m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", krc_id));
                            //    //m_cmd.Parameters[":FK_LEGAL"].Value = Convert.ToInt32(Legal_List[i].Trim());

                            //    m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", txtKrcName));

                            //    //m_cmd.Parameters[":ADRESAT"].Value = Convert.ToString(Legal_Name_List[i]);
                            //    if (m_cmd.ExecuteNonQuery() != -1)
                            //    {
                            //        iCnt++;
                            //        prbWritingDBF.PerformStep();
                            //    }
                            //}
                            #endregion

                            //tran.Commit();
                            //con.Close();


                            foreach (DataRow row in tbl.Rows)
                            {

                                m_cmd = new OleDbCommand();
                                m_cmd.Connection = con;
                                m_cmd.Transaction = tran;
                                m_cmd.CommandText = "UPDATE ZAPROS SET RESULT = :RESULT, TEXT = :TEXT, DATE_RESP = :DATE_RESP, DATE_RES = :DATOTV, STATUS = '�����'";
                                //m_cmd.CommandText += " WHERE USCODE = " + Convert.ToString(Convert.ToInt32(row["NOMSPI"]));
                                //m_cmd.CommandText += " AND TEMP = " + Convert.ToString(Convert.ToInt32(Convert.ToInt32(row["NOMIP"])));
                                m_cmd.CommandText += " WHERE NUM_IP = '" + Convert.ToString(row["ZAPROS"]).TrimEnd() + "'";
                                m_cmd.CommandText += " AND FK_LEGAL = " + Convert.ToString(krc_id).TrimEnd();


                                string txtResponse = "";
                                //string priz = Convert.ToString(row["PRIZ"]).TrimEnd();
                                //if (priz.ToUpper().Equals("T"))
                                //int priz = 0;

                                //if (!(int.TryParse(Convert.ToString(row["PRIZ"]), out priz)))
                                //{
                                //    priz = 2;
                                //}

                                ////int priz = Convert.ToInt32(row["PRIZ"]);
                                //if (priz == 1)
                                //{
                                //    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(1)));
                                //    txtResponse = "������� �������� ����������� ������.\n";
                                //    txtResponse += "�����: " + Convert.ToString(row["ADRES"]).TrimEnd() + "\n";
                                //    txtResponse += "C���� ������, �� ������� ����� �������� ���������: " + Convert.ToString(row["SUMMA"]).TrimEnd() + ". " + Convert.ToString(row["KOMMENT"]).TrimEnd() + "\n";

                                //}
                                //else
                                //{
                                //    if (priz == 0)
                                //    {
                                //        m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));
                                //        txtResponse = "��� ������ � �������� �� ������� �� " + Convert.ToDateTime(row["DATZAP"]).ToShortDateString();
                                //    }
                                //    else
                                //    {
                                //        m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));
                                //        txtResponse = "��� ������ � �������� �� ������� �� " + Convert.ToDateTime(row["DATZAP"]).ToShortDateString() + " " + Convert.ToString(row["SUMMA"]).TrimEnd();
                                //    }
                                //}
                                if (Convert.ToInt32(row["SUMPL"]) != 0)
                                {
                                    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(1)));
                                    m_cmd.Parameters.Add(new OleDbParameter(":TEXT", "������� �������� ����� �������������: " + Convert.ToString(row["SUMPL"]) + " . ���� �������: " + Convert.ToString(row["DATPL"])));// ����� ������ - ������
                                }
                                else
                                {
                                    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));
                                    m_cmd.Parameters.Add(new OleDbParameter(":TEXT", "�� ������� ��������������� ��������� �������� �� ���������"));// ����� ������ - ������
                                }

                                //m_cmd.Parameters.Add(new OleDbParameter(":TEXT", cutEnd(txtResponse, 500)));

                                m_cmd.Parameters.Add(new OleDbParameter(":DATE_RESP", DateTime.Today));

                                DateTime tmpDate = DateTime.Today;
                                /*
                                if (row.Table.Columns.Contains("DATOTV"))
                                {
                                    if (!(DateTime.TryParse(Convert.ToString(row["DATOTV"]), out tmpDate)))
                                        tmpDate = DateTime.Today;
                                }
                                */
                                m_cmd.Parameters.Add(new OleDbParameter(":DATOTV", tmpDate));

                                int result = m_cmd.ExecuteNonQuery();
                                if (result != -1)
                                {
                                    iCnt++;
                                    prbWritingDBF.PerformStep();
                                    prbWritingDBF.Refresh();
                                    System.Windows.Forms.Application.DoEvents();
                                }
                                //if (result > 0)
                                //{
                                //    MessageBox.Show("���, �������� �����!", "��������!", MessageBoxButtons.OK);
                                //}
                            }
                            tran.Commit();
                            con.Close();
                            MessageBox.Show("���������� �������: " + iCnt.ToString() + ".\n ������ ����� ����������� ������ �������.", "���������", MessageBoxButtons.OK);

                            //**********������������**�������**pens************
                            //���� ��������� ������ � ���� + ����������� ���������� ����� 
                            //��� ���������� �� ���������. 

                            //������ ���� ���������
                            DataTable dtspi = ds.Tables.Add("SPI");

                            DBFcon.Open();
                            m_cmd = new OleDbCommand();
                            m_cmd.Connection = DBFcon;
                            // ������� ������ �� ���� priz ����� �� ���������� ������ ��������
                            //m_cmd.CommandText = "SELECT DISTINCT NOMSPI FROM " + tablename + " WHERE priz = '1'";
                            //m_cmd.CommandText = "SELECT DISTINCT NOMSPI FROM " + tablename + " WHERE priz = '1'";
                            m_cmd.CommandText = "SELECT DISTINCT NOMSPI FROM " + tablename + " WHERE sumpl > 0";

                            using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                            {
                                ds.Load(rdr, LoadOption.OverwriteChanges, dtspi);
                                rdr.Close();
                            }

                            DBFcon.Close();

                            if (OooIsInstall)
                            {
                                //OOo start
                                OOo_Writer OOo_cld = new OOo_Writer();
                                OOo_cld.OOo_Krc(tablename, ds, con, this);
                            }
                            //else
                            //{
                            //    //      ������ ��� �����

                            //    Microsoft.Office.Interop.Word.ApplicationClass app = new ApplicationClass();

                            //    object s1 = "";
                            //    object fl = false;
                            //    object t = WdNewDocumentType.wdNewBlankDocument;
                            //    object fl2 = true;

                            //    Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref s1, ref fl, ref t, ref fl2);

                            //    Paragraph par;

                            //    int spi;
                            //    int sch_line;
                            //    int fl_fst = 1;

                            //    string nline = "";

                            //    prbWritingDBF.Value = 0;
                            //    prbWritingDBF.Maximum = dtspi.Rows.Count;
                            //    prbWritingDBF.Step = 1;

                            //    foreach (DataRow drspi in dtspi.Rows)
                            //    {

                            //        sch_line = 0;
                            //        if (fl_fst == 1)
                            //        {
                            //            sch_line = 1;
                            //            fl_fst = 0;
                            //            par = doc.Paragraphs[1];
                            //        }
                            //        else
                            //        {
                            //            object oMissing = System.Reflection.Missing.Value;
                            //            par = doc.Paragraphs.Add(ref oMissing);
                            //            object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                            //            par.Range.InsertBreak(ref oPageBreak);
                            //        }

                            //        par.Range.Font.Name = "Courier";
                            //        par.Range.Font.Size = 8;
                            //        float a = par.Range.PageSetup.RightMargin;
                            //        float b = par.Range.PageSetup.LeftMargin;
                            //        float c = par.Range.PageSetup.TopMargin;
                            //        float d = par.Range.PageSetup.BottomMargin;

                            //        par.Range.PageSetup.RightMargin = 30;
                            //        par.Range.PageSetup.LeftMargin = 30;
                            //        par.Range.PageSetup.TopMargin = 20;
                            //        par.Range.PageSetup.BottomMargin = 20;

                            //        par.Range.Text += "������ ������� �� �������� ��-� � ��� � ������� ������\n";
                            //        par.Range.Text += "�� ������ � ��� �� " + DateTime.Today.ToShortDateString() + "\n";
                            //        //par.Range.Text += "�� ������ � " + dat1.ToShortDateString() + " �� " + dat2.ToShortDateString() + "\n";

                            //        spi = Convert.ToInt32(drspi["NOMSPI"]);

                            //        par.Range.Text += "��-�: " + GetSpiName3(Convert.ToInt32(drspi["NOMSPI"])) + "\n";

                            //        sch_line += 6;

                            //        foreach (DataRow row in tbl.Rows)
                            //        {
                            //            if (spi == Convert.ToInt32(row["NOMSPI"]))
                            //            {
                            //                //int priz = Convert.ToInt32(row["PRIZ"]);
                            //                int priz = 0;

                            //                if (!(int.TryParse(Convert.ToString(row["PRIZ"]), out priz)))
                            //                {
                            //                    priz = 2;
                            //                }

                            //                if (priz == 1)
                            //                {
                            //                    par.Range.Text += Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["NAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["FNAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["SNAMEDOL"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd();
                            //                    par.Range.Text += Convert.ToString(row["ADRES"]).TrimEnd() + "";
                            //                    par.Range.Text += "������� �������� ����������� ������. C���� ������, �� ������� ����� �������� ���������: " + Convert.ToString(row["SUMMA"]).TrimEnd() + "\n";
                            //                    sch_line += 5;
                            //                }

                            //                //if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                            //                //string priz = Convert.ToString(row["PRIZ"]).TrimEnd();
                            //                //if (priz.ToUpper().Equals("T"))
                            //            }
                            //        }
                            //        // ���� ������ �������������� � ������� ���, �� ��� � �����
                            //        if (sch_line == 6)
                            //        {
                            //            par.Range.Text += "��� ������������� ������� �� �������� � ������� ������ � ���������.";
                            //            sch_line++;
                            //            object oMissing = System.Reflection.Missing.Value;
                            //            par.Range.Delete(ref oMissing, ref oMissing);
                            //        }

                            //        prbWritingDBF.PerformStep();
                            //        prbWritingDBF.Refresh();
                            //        System.Windows.Forms.Application.DoEvents();
                            //    }

                            //    app.Visible = true;
                            //    //*************************************************
                            //}
                        }
                        catch (OleDbException ole_ex)
                        {
                            foreach (OleDbError err in ole_ex.Errors)
                            {
                                MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                            }
                            //return false;
                        }
                        catch (Exception ex)
                        {
                            //if (DBFcon != null) DBFcon.Close();
                            MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                            //return false;
                        }
                        //return true;
                    }
                }
            }
        }

        /// <summary>
        /// Get count of lines in given text file
        /// </summary>
        private int GetFilesLineCnt(string filename)
        {
            int count = 0;
            string line;
            TextReader reader = new StreamReader(filename);
            while ((line = reader.ReadLine()) != null)
            {
                count++;
            }
            reader.Close();
            return count;
        }


        /// <summary>
        /// Split file based on max size in MB
        /// </summary>
        private int SplitFile(string filePath)
        {
            string wokingCopy = string.Empty;
            string xmlIndexFile = string.Empty;
            double numOfNewFiles = 1;
            double maxFileSplitLines = 15000; //100 MB

            //Output document writer
            //StreamWriter sw = null;
            
            nodePathDic.Clear();
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("�������� ���� ��� ���������� ��� �� ��������� ������.");
                return 0;
            }

            FileInfo fi = new FileInfo(filePath);
            //double origFileSize = (double)fi.Length;
            double origLineCount = Convert.ToDouble(GetFilesLineCnt(filePath));

            //numOfNewFiles = Math.Ceiling(origFileSize / maxFileSplitSize);
            numOfNewFiles = Math.Ceiling(origLineCount / maxFileSplitLines);
            
            string filePathPart = filePath.Substring(0, filePath.Length - 10);
            string fiName = "";
            if (fi.Name.Length > 4)
            {
                fiName = fi.Name.Substring(0, fi.Name.Length - 4);
            }
            else fiName = fi.Name;
            
            string filePart = filePathPart + "/" + fiName + ".part1" + fi.Extension;
            int fileCnt = 1;
            long writeFilePosition = 0;

            if (numOfNewFiles > 1)
            {

                using (StreamReader sr = new StreamReader(filePath, Encoding.UTF8))
                {
                    int pos = 0;
                    int LinesCnt = 0;
                    filePart = filePathPart + "/" + fiName + ".part" + fileCnt + fi.Extension;
                    //Read each line in XML document as regular file stream.
                    StreamWriter sw = new StreamWriter(filePart, false);

                    Regex rx = new Regex(@"<", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                    string nodeName = string.Empty;
                    do
                    {

                        string line = sr.ReadLine();
                        LinesCnt++;
                        pos += Encoding.UTF8.GetByteCount(line) + 2;// 2 extra bites for end of line chars.


                        MatchCollection m = rx.Matches(line);
                        //Save index of this node into dictionary
                        foreach (Match mt in m)
                        {
                            nodeName = line.Split(' ').Length == 0 ? line.Substring(1, line.LastIndexOf('>') - 1) : line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0];
                            if (!nodeName.Contains("?xml") && !nodePathDic.ContainsKey(pos + mt.Index))
                            {
                                nodePathDic.Add(pos + mt.Index, nodeName);
                            }
                            break;
                        }

                        sw.WriteLine(line);
                        sw.Flush();
                        writeFilePosition = sw.BaseStream.Position;

                        //If we at the limit of new file let's get a last node and write it to this file 
                        //and create a new split file.
                        if (LinesCnt > maxFileSplitLines * fileCnt)
                        {
                            //if (pos > maxFileSplitSize * fileCnt) {
                            int lastNodeStartPosition = 0;
                            string lastNodeName = string.Empty;
                            string ln = string.Empty;
                            string completeLastNode = GetLastNode(filePath, out lastNodeStartPosition, out lastNodeName);

                            //Some synchronization. TODO: needs to be optimized but it works "AS IS"
                            do
                            {
                                //Skip rest of the node....
                                ln = sr.ReadLine();
                                if (ln == null)
                                    break;
                                LinesCnt++;
                                pos += Encoding.UTF8.GetByteCount(ln) + 2;
                            } while (!ln.Contains(lastNodeName));

                            //Get position where we will begin to read again in our original XML file. We want to skip to the end of last 
                            //complete node we wrote to the file.
                            long swPosition = (writeFilePosition - (nodePathDic.Keys[nodePathDic.Count - 1] - lastNodeStartPosition)) + 2;
                            sw.BaseStream.Position = swPosition >= 0 ? swPosition : 0;
                            sw.Write("\n");
                            sw.WriteLine("<!-- End of " + filePathPart + "/" + fiName + ".part" + fileCnt + fi.Extension + ". " + fileCnt + " out of " + numOfNewFiles + " -->");

                            sw.WriteLine(completeLastNode + "\n\n");
                            sw.WriteLine(nodePathDic.Values[0].Replace("<", "</"));

                            filePart = filePathPart + "/" + fiName + ".part" + (++fileCnt) + fi.Extension;
                            sw.Flush();
                            sw.Close();

                            sw = new StreamWriter(filePart, false);
                            sw.WriteLine(nodePathDic.Values[0]);
                            sw.WriteLine("<!-- Start of " + filePathPart + "/" + fiName + ".part" + fileCnt + fi.Extension + ". " + fileCnt + " out of " + numOfNewFiles + " -->");
                            sw.Flush();
                        }
                    } while (!sr.EndOfStream);

                    //Clean up...
                    sw.Flush();
                    sw.Close();
                    sr.Close();
                    sw.Close();
                }

                return fileCnt;
            }

            else return 0;
        }

        /// <summary>
        /// Try to find complete last node at the end of this page.
        /// </summary>
        /// <param name="fileSourceName">original XML file to split</param>
        /// <param name="lastNodeStartPosition">position</param>
        /// <param name="nodeName">node name</param>
        /// <returns></returns>
        private string GetLastNode(string fileSourceName, out int lastNodeStartPosition, out string nodeName)
        {
            int lastIdx = nodePathDic.Count - 1;
            string output = string.Empty;

            //Check to avoid error.
            if (lastIdx < 0)
            {
                lastNodeStartPosition = 0;
                nodeName = "";
                return ("");
            }

        redo:
            try
            {
                while (!nodePathDic.Values[lastIdx].Trim().Contains("</") && lastIdx >= 0)
                {
                    lastIdx--;
                }
            }
            catch
            {
                MessageBox.Show("Error: \nTry to change file size to larger number. \nOnly found: " + nodePathDic.Count + " elements");
                lastNodeStartPosition = 0;
                nodeName = "";
                return null;
            }

            nodeName = nodePathDic.Values[lastIdx].Trim();
            if (nodeName.IndexOf("</") != 0)
            {
                lastIdx--;
                //String may not be formated well so let's try again
                goto redo;
            }

            lastNodeStartPosition = nodePathDic.Keys[lastIdx];

            //Get this node at position we found it should be at
            using (FileStream fs = new FileStream(fileSourceName, FileMode.Open, FileAccess.Read))
            {
                fs.Seek(lastNodeStartPosition, SeekOrigin.Begin);
                try
                {
                    using (XmlReader reader = XmlReader.Create(fs))
                    {
                        reader.MoveToContent();
                        XmlDocument d = new XmlDocument();
                        d.Load(reader.ReadSubtree());
                        output = d.InnerXml;
                        reader.Close();
                    }
                }
                catch (Exception anyEr)
                {
                    //Just in case we want to see were our pointer is...
                    byte[] buffer = new byte[512];
                    fs.Read(buffer, 0, buffer.Length);
                    Console.Out.WriteLine("Error:{0}, node:{1}", anyEr.Message, Encoding.UTF8.GetString(buffer));
                }
                finally
                {
                    fs.Close();
                }
            }
            //Add some spaces between ending nodes.
            output = output.Replace("><", ">\n<");
            return output;
        }

        private void btnSplit_Click(object sender, EventArgs e)
        {
            int nNumOfNewFiles = 0;
            string filePath = "";
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "All files|*.*";
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;
                if (filePath.Trim().Length > 10)
                {
                    nNumOfNewFiles = SplitFile(filePath);
                    if(nNumOfNewFiles != 0)
                        MessageBox.Show("��������. ���� �������� �� " + nNumOfNewFiles.ToString() + " ������.\n������ ������������ ���� ����� � ��������� .partX\n X - ���������� ����� �����");
                    else MessageBox.Show("��������. ���� �� ��������� ������ �� �����.\n");
                }
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            //tabKRC.Hide();
            //tabGibd.Hide();
        }

        private decimal FindIDNum(string txtNomID, double nSumID, DateTime dtDatID)
        {
            OleDbTransaction tran;
            string txtSql = "";
            decimal res = -1;
            try
            {

                    if (con != null && con.State != ConnectionState.Closed) con.Close();
                    con.Open();
                    tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                    //txtSql = "select d_ip_d.id from doc_ip_doc d_ip_d left join doc_ip d_ip on d_ip_d.id = d_ip.id left join document d on d_ip_d.id = d.id where ((d_ip_d.id_debtcls = 22) or (d_ip_d.id_debtcls = 37) or (d_ip_d.id_debtcls = 38) or (d_ip_d.id_debtcls = 45)) and (d.docstatusid != -1) and (d.docstatusid != 7) and (d.docstatusid != 10) and d_ip.id_debtsum = " + nSumID.ToString() + " and d_ip_d.id_docdate = '" + dtDatID.ToShortDateString() + "' and d_ip_d.id_docno = '" + Convert.ToString(txtNomID) + "'";
                    txtSql = "select i_id.ip_id from i_id left join document d on i_id.ip_id = d.id where ((i_id.debtcls = 22) or (i_id.debtcls = 37) or (i_id.debtcls = 38) or (i_id.debtcls = 45)) and (d.docstatusid != -1) and (d.docstatusid != 7) and (d.docstatusid != 10) and i_id.debtsum = " + nSumID.ToString().Replace(',', '.') + " and i_id.id_docdate = '" + dtDatID.ToShortDateString() + "' and i_id.id_docno = '" + Convert.ToString(txtNomID) + "'";
                    OleDbCommand cmd = new OleDbCommand(txtSql, con, tran);
                    res = Convert.ToDecimal(cmd.ExecuteScalar());
                    tran.Rollback();
                    con.Close();
            }
            catch (OleDbException ole_ex)
            {
                //if (tran != null) {
                    //tran.Rollback();
                //}

                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }

            
            if (con != null)
            {
                con.Close();
            }

            return res;
        }

        private decimal FindIDNum(string txtQuery)
        {

            decimal res = -1;
            try
            {

                    if (con != null && con.State != ConnectionState.Closed) con.Close();
                    con.Open();
                    OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                    OleDbCommand cmd = new OleDbCommand("select d_ip_d.id from doc_ip_doc d_ip_d left join doc_ip d_ip on d_ip_d.id = d_ip.id left join document d on d_ip_d.id = d.id where (d.docstatusid != -1) and (d.docstatusid != 7) and (d.docstatusid != 10) and d_ip_d.id_docno = '" + Convert.ToString(txtQuery) + "'", con, tran);
                    res = Convert.ToDecimal(cmd.ExecuteScalar());
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
        

        private void btnFindNumID_Click(object sender, EventArgs e)
        {
            
                    string tablename = "GIBDD_PLATEZH";
                    OleDbConnection ConG;
                    Int32 iCnt = 0;

                    try
                    {
                        ConG = new OleDbConnection(constrGIBDD);

                        // ������� ��� �������� � ��� ��������������
                        DeleteUsedGibddPlat(ConG, true, true);

                        DataSet ds = new DataSet();
                        string txtSql = "SELECT * FROM " + tablename + " WHERE  FL_USE = 0"; // ������� ������ ��, ������� ��� �� ���� �����������
                        DataTable tbl = GetDataTableFromFB(constrGIBDD, txtSql, tablename, IsolationLevel.RepeatableRead);
                        if (tbl != null)
                        {
                            l_loadgibd.Text = tbl.Rows.Count.ToString();
                        }

                        prbWritingDBF.Value = 0;
                        prbWritingDBF.Maximum = tbl.Rows.Count;
                        prbWritingDBF.Step = 1;

                        string txtContent = "";

                        //AlterIndxI_ID(con, true);

                        foreach (DataRow row in tbl.Rows)
                        {
                            string txtNomID = Convert.ToString(row["NOMID"]).TrimEnd();
                            DateTime dtDatID = Convert.ToDateTime(row["DATID"]);
                            Double nSumID = Convert.ToDouble(row["SUMM"]);
                            Double nSumDoc = Convert.ToDouble(row["SUMM_DOC"]);
                            DateTime dtDateDoc = Convert.ToDateTime(row["DATE_DOC"]);
                            DateTime dtExtDocDate = Convert.ToDateTime(row["DATE_ISH"]);
                            string txtExtDocNum = Convert.ToString(row["ISH_NUMBER"]);
                            string txtFIO_D = Convert.ToString(row["FIO_D"]).TrimEnd();
                            string txtNumKvit = Convert.ToString(row["NUM_DOC"]).TrimEnd();
                            DateTime dtBornD = Convert.ToDateTime(row["BORN_D"]);
                            DateTime dtReestrVhodDate = Convert.ToDateTime(row["DATE_VH"]);

                            // ������� ������ � ���� �� ��� ��� ������ ������ ������ ��
                            decimal id = FindIDNum(txtNomID, nSumID, dtDatID);
                            if (id > 0)
                            {
                                iCnt++;
                                // ��������� ������ ������� �� ���
                                int i = 0;
                                while (txtFIO_D.IndexOf("  ") != -1)
                                {
                                    txtFIO_D = txtFIO_D.Replace("  ", " ");
                                    i++;
                                    if (i > 200)
                                    {
                                        break;
                                    }
                                }

                                // ��������� txtContent
                                txtContent = "������� " + txtFIO_D + " "; ;
                                if (!dtBornD.Equals(Convert.ToDateTime("01.01.1800")))
                                {
                                    txtContent += "(���� �������� " + dtBornD.ToShortDateString() + ") ";
                                }
                                txtContent += dtDateDoc.ToShortDateString() + " ������� " + Money_ToStr(nSumDoc) + " � ��������� �� ������ " + txtNumKvit + " �� �� � " + txtNomID + " �� " + dtDatID.ToShortDateString() + ".";

                                //MessageBox.Show("����� �� � " + txtNomID + ". IP_ID = " + id.ToString(), "��������!", MessageBoxButtons.OK);
                                string txtNumber = Convert.ToString(row["NUMBER"]);
                                UpdateGibddPlatezh(ConG, txtNumber, 1);
                                ID_InsertPlatDocTo_PK_OSP(con, 1, 86011011483815, nSumDoc, DateTime.Today, id, dtDateDoc, txtNumKvit, mvd_id, txtFIO_D);
                                ID_InsertOtherIP_DocTo_PK_OSP(con, 1, 86011011483815, dtReestrVhodDate, id, dtExtDocDate, txtExtDocNum, txtContent, mvd_id);
                            }

                            prbWritingDBF.PerformStep();
                            prbWritingDBF.Refresh();
                            System.Windows.Forms.Application.DoEvents();
                        }
                        
                        //AlterIndxI_ID(con, false);


                        l_zaprgibd.Text = iCnt.ToString();


                        MessageBox.Show("������ ������� ���������.", "���������", MessageBoxButtons.OK);

                    }
                    catch (OleDbException ole_ex)
                    {
                        foreach (OleDbError err in ole_ex.Errors)
                        {
                            MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                        }
                        //return false;
                    }
                    catch (Exception ex)
                    {
                        //if (DBFcon != null) DBFcon.Close();
                        MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                        //return false;
                    }
                }

        private void b_zaprgibd_Click(object sender, EventArgs e)
        {

        }

        private void l_zaprgibd_Click(object sender, EventArgs e)
        {

        }

        private string GetIPNum(OleDbConnection con, string txtCode)
        {

            Decimal code;
            if (!Decimal.TryParse(txtCode, out code))
            {
                code = -1;
            }
            string res = "";
            try
            {
                if (code != -1)
                {
                    if (con != null && con.State != ConnectionState.Closed) con.Close();
                    con.Open();
                    OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                    //OleDbCommand cmd = new OleDbCommand("Select ipno from O_IP WHERE ID = " + Convert.ToString(code), con, tran);
                    OleDbCommand cmd = new OleDbCommand("select ip_d.doc_number NOMIP from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id where req.id = " + Convert.ToString(code), con, tran);
                    res = Convert.ToString(cmd.ExecuteScalar());
                    tran.Rollback();
                    con.Close();
                }
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", System.Windows.Forms.MessageBoxButtons.OK);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", System.Windows.Forms.MessageBoxButtons.OK);
            }
            return res;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            //frmRewriteDialog frmRwD = new frmRewriteDialog();
            //int iResult = frmRwD.ShowForm();
            //lblTest.Text = iResult.ToString();
        }

        private void btnLoadKrcDbfFile_Click(object sender, EventArgs e)
        {
            // �������� ������ �� ���������� ������ �� ������ �� ���
            DT_krc_oplat = null;
            Int32 iCnt = 0;
            openFileDialog1.Filter = "DBF �����(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();

            if (res == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {

                    //ChangeByte(openFileDialog1.FileName, 0x65, 30);
                    string tablename = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - 4);
                    tablename = tablename.Substring(tablename.LastIndexOf("\\") + 1);

                    // ��������� �� ����� ����� ���� ���. ��������� (�������)
                    string txtDateIsh = tablename.Substring(6, 2) + '.' + tablename.Substring(4, 2) + '.' + tablename.Substring(0, 4);

                    DateTime dtDateIsh; 
                    if (!DateTime.TryParse(txtDateIsh, out dtDateIsh)) // +
                    {
                        dtDateIsh = DateTime.MinValue;
                    }

                    // ��������� �� ����� ����� ��� ����� ���������
                    string txtIshNumber = tablename.Substring(9, tablename.Length - 9); // +

                    try
                    {
                        DataSet ds = new DataSet();
                        DataTable tbl = ds.Tables.Add(tablename);
                        DBFcon = new OleDbConnection();
                        DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=RUSSIAN");
                        DBFcon.Open();
                        m_cmd = new OleDbCommand();
                        m_cmd.Connection = DBFcon;
                        m_cmd.CommandText = "SELECT * FROM " + tablename;
                        using (OleDbDataReader rdr = m_cmd.ExecuteReader(CommandBehavior.Default))
                        {
                            ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                            rdr.Close();
                        }

                        DBFcon.Close();
                        if (tbl != null)
                        {
                            lblKrc_load.Text = tbl.Rows.Count.ToString();
                            DT_krc_oplat = tbl;
                            MessageBox.Show("������ ������� ����������, ����������� ��������� ������.", "���������", MessageBoxButtons.OK);
                        }
                        // ������ ��������� - ������ ����� �������� ������


                        prbWritingDBF.Value = 0;
                        if (DT_krc_oplat != null)
                        {
                                prbWritingDBF.Maximum = DT_krc_oplat.Rows.Count;
                                prbWritingDBF.Step = 1;

                                string txtContent = "";

                                //AlterIndxI_ID(con, true);

                                foreach (DataRow row in DT_krc_oplat.Rows)
                                {
                                    string txtNumber = Convert.ToString(row["Number"]).TrimEnd();// +
                                    string txtGibddIDNumber = txtNumber; // txtNumber.Substring(1, 7); // +

                                    string txtDatID = Convert.ToString(row["Date_exec"]);// +
                                    DateTime dtDatID;

                                    Double nSum;
                                    string txtSum = Convert.ToString(row["Summa"]); // +

                                    string txtFioD = Convert.ToString(row["Plat_name"]).TrimEnd();// +
                                    string txtDateDoc = Convert.ToString(row["Date_doc"]); // +
                                    DateTime dtDateDoc;

                                    string txtNumDoc = Convert.ToString(row["Num_doc"]).TrimEnd();// +

                                    string txtBornD = Convert.ToString(row["Date_plat"]); // +
                                    DateTime dtBornD;

                                    if (!DateTime.TryParse(txtBornD, out dtBornD)) // +
                                    {
                                        dtBornD = Convert.ToDateTime("01.01.1800");
                                    }

                                    Double nSumDoc;
                                    string txtSumDoc = Convert.ToString(row["Summa_doc"]);// +

                                    // ���� �������������� - ���� �� ������� ��� ����� �� ����� ����
                                    //if (txtGibddIDNumber[0] == '0') // +
                                    //{
                                    //    txtGibddIDNumber = txtGibddIDNumber.Substring(1, 6);
                                    //}

                                    if (!DateTime.TryParse(txtDatID, out dtDatID)) // +
                                    {
                                        dtDatID = DateTime.MinValue;
                                    }

                                    if (!Double.TryParse(txtSum, out nSum)) // +
                                    {
                                        nSum = -1;
                                    }

                                    if (!DateTime.TryParse(txtDateDoc, out dtDateDoc)) // +
                                    {
                                        dtDateDoc = DateTime.MinValue;
                                    }

                                    if (!Double.TryParse(txtSumDoc, out nSumDoc))// +
                                    {
                                        nSumDoc = -1;
                                    }

                                    DateTime dtReestrVhodDate = DateTime.Today;

                                    //m_cmd.Parameters.Add(new OleDbParameter(":NUMBER", txtNumber));
                                    //m_cmd.Parameters.Add(new OleDbParameter(":NOMID", txtGibddIDNumber));
                                    //m_cmd.Parameters.Add(new OleDbParameter(":DATID", dtDatID)); // ���� ��
                                    //m_cmd.Parameters.Add(new OleDbParameter(":SUMM", nSum));     // ����� ��
                                    //m_cmd.Parameters.Add(new OleDbParameter(":SUMM_DOC", nSumDoc));   //����� ������
                                    //m_cmd.Parameters.Add(new OleDbParameter(":FIO_D", txtFioD));
                                    //m_cmd.Parameters.Add(new OleDbParameter(":DATE_DOC", dtDateDoc)); // ���� ������


                                    //m_cmd.Parameters.Add(new OleDbParameter(":ISH_NUMBER", txtIshNumber));   // ��� �����
                                    //m_cmd.Parameters.Add(new OleDbParameter(":DATE_ISH", dtDateIsh));   // ���� ��� 
                                    //m_cmd.Parameters.Add(new OleDbParameter(":FL_USE", Convert.ToInt32(0)));   // ���� - �����/�������
                                    //m_cmd.Parameters.Add(new OleDbParameter(":NUM_DOC", txtNumDoc));   // ����� ���������
                                    //m_cmd.Parameters.Add(new OleDbParameter(":BORN_D", dtBornD));   // ���� �������� ��������
                                    //m_cmd.Parameters.Add(new OleDbParameter(":DATE_VH", DateTime.Today));   // ���� �������� �������

                                
                                    // ������� ������ � ���� �� ��� ��� ������ ������ ������ ��
                                    decimal id = FindIDNum(txtGibddIDNumber, nSum, dtDatID);
                                    if (id > 0)
                                    {
                                        iCnt++;
                                        // ��������� ������ ������� �� ���
                                        int i = 0;
                                        while (txtFioD.IndexOf("  ") != -1)
                                        {
                                            txtFioD = txtFioD.Replace("  ", " ");
                                            i++;
                                            if (i > 200)
                                            {
                                                break;
                                            }
                                        }

                                        // ��������� txtContent
                                        txtContent = "������� " + txtFioD + " "; ;
                                        if (!dtBornD.Equals(Convert.ToDateTime("01.01.1800")))
                                        {
                                            txtContent += "(���� �������� " + dtBornD.ToShortDateString() + ") ";
                                        }
                                        txtContent += dtDateDoc.ToShortDateString() + " ������� " + Money_ToStr(nSumDoc) + " � ��������� �� ������ " + txtNumDoc + " �� �� � " + txtGibddIDNumber + " �� " + dtDatID.ToShortDateString() + ".";

                                        //MessageBox.Show("����� �� � " + txtNomID + ". IP_ID = " + id.ToString(), "��������!", MessageBoxButtons.OK);
                                        
                                        //UpdateGibddPlatezh(ConG, txtNumber, 1);

                                        ID_InsertPlatDocTo_PK_OSP(con, 1, 86011011483815, nSumDoc, DateTime.Today, id, dtDateDoc, txtNumDoc, krc_id, txtFioD);
                                        ID_InsertOtherIP_DocTo_PK_OSP(con, 1, 86011011483815, dtReestrVhodDate, id, dtDateIsh, txtIshNumber, txtContent, krc_id);
                                    }

                                prbWritingDBF.PerformStep();
                                prbWritingDBF.Refresh();
                                System.Windows.Forms.Application.DoEvents();
                            }

                            //AlterIndxI_ID(con, false);

                            lblKRC_z.Text = iCnt.ToString();

                            MessageBox.Show("������ ������� ���������.", "���������", MessageBoxButtons.OK);
                        }

                    }
                    catch (OleDbException ole_ex)
                    {
                        foreach (OleDbError err in ole_ex.Errors)
                        {
                            MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                        }
                        //return false;
                    }
                    catch (Exception ex)
                    {
                        //if (DBFcon != null) DBFcon.Close();
                        MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                        //return false;
                    }
                    //return true;

                }
            }


        }


        private decimal CreateLLog(OleDbConnection gibdd_con, decimal nStatus, int nPackType, string txtAgreement_code, decimal nParent_ID, string txtLOG)
        {
            decimal nID = 0;
            decimal nOspNum = 0;
            OleDbCommand cmd, cmdIns;
            OleDbTransaction tran = null;

            try
            {
                if (gibdd_con != null && gibdd_con.State != ConnectionState.Closed) gibdd_con.Close();
                gibdd_con.Open();
                tran = gibdd_con.BeginTransaction(IsolationLevel.ReadCommitted);

                // �������� ����� ����
                cmd = new OleDbCommand("SELECT gen_id(GEN_LOCAL_LOG_ID, 1) FROM RDB$DATABASE", gibdd_con, tran);
                nID = Convert.ToDecimal(cmd.ExecuteScalar());

                // �������� OSPNUM
                nOspNum = 10000 + GetOSP_Num();

                // �������� DOCUMENT
                cmdIns = new OleDbCommand();
                cmdIns.Connection = gibdd_con;
                cmdIns.Transaction = tran;
                cmdIns.CommandText = "insert into LOCAL_LOGS (ID, OSPNUM, PACKDATE, PACK_TYPE, CONV_CODE, PACK_STATUS, PACK_COUNT, PARENT_ID, LOG)";
                cmdIns.CommandText += " VALUES (:ID ,:OSPNUM, :PACKDATE, :PACK_TYPE, :CONV_CODE, :PACK_STATUS, 0, :PARENT_ID, :TXT_LOG)";

                cmdIns.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));
                cmdIns.Parameters.Add(new OleDbParameter(":OSPNUM", Convert.ToInt32(nOspNum)));
                cmdIns.Parameters.Add(new OleDbParameter(":PACKDATE", DateTime.Now));
                cmdIns.Parameters.Add(new OleDbParameter(":PACK_TYPE", nPackType));
                cmdIns.Parameters.Add(new OleDbParameter(":CONV_CODE", txtAgreement_code));
                cmdIns.Parameters.Add(new OleDbParameter(":PACK_STATUS", nStatus));
                cmdIns.Parameters.Add(new OleDbParameter(":PARENT_ID", nParent_ID));
                cmdIns.Parameters.Add(new OleDbParameter(":TXT_LOG", txtLOG));
                

                if (cmdIns.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to document table LOCAL_LOGS = " + nID.ToString());
                    throw ex;
                }

                tran.Commit();
                gibdd_con.Close();

            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (gibdd_con != null)
                {
                    gibdd_con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
                if (gibdd_con != null)
                {
                    gibdd_con.Close();
                }
            }

            return nID;
        }

        private bool WriteLLog(OleDbConnection gibdd_con, decimal nLLogID, string txtText)
        {
            //string txtSql = "Update DX_PACK set PLAIN_LOG = PLAIN_LOG || '" + txtText + "' where ID = " + nPackID.ToString();
            string txtSql = "Update LOCAL_LOGS set LOG = LOG || '" + txtText + "' where ID = " + nLLogID.ToString();
            return UpdateSqlExecute(gibdd_con, txtSql);
        }

        private bool UpdateLLogStatus(OleDbConnection gibdd_con, decimal nLLogID, int Status)
        {
            //string txtSql = "Update DX_PACK set PLAIN_LOG = PLAIN_LOG || '" + txtText + "' where ID = " + nPackID.ToString();
            string txtSql = "Update LOCAL_LOGS set PACK_STATUS = " + Status.ToString() + " where ID = " + nLLogID.ToString();
            return UpdateSqlExecute(gibdd_con, txtSql);
        }

        private bool UpdateLLogCount(OleDbConnection gibdd_con, decimal nLLogID, int Count)
        {
            //string txtSql = "Update DX_PACK set PLAIN_LOG = PLAIN_LOG || '" + txtText + "' where ID = " + nPackID.ToString();
            string txtSql = "Update LOCAL_LOGS set PACK_COUNT = " + Count.ToString() + " where ID = " + nLLogID.ToString();
            return UpdateSqlExecute(gibdd_con, txtSql);
        }

        private bool AppendLLogCount(OleDbConnection gibdd_con, decimal nLLogID, int Count)
        {
            //string txtSql = "Update DX_PACK set PLAIN_LOG = PLAIN_LOG || '" + txtText + "' where ID = " + nPackID.ToString();
            string txtSql = "Update LOCAL_LOGS set PACK_COUNT = PACK_COUNT + " + Count.ToString() + " where ID = " + nLLogID.ToString();
            return UpdateSqlExecute(gibdd_con, txtSql);
        }

        private bool UpdateLLogParent(OleDbConnection gibdd_con, decimal nLLogID, decimal nParentID)
        {
            //string txtSql = "Update DX_PACK set PLAIN_LOG = PLAIN_LOG || '" + txtText + "' where ID = " + nPackID.ToString();
            string txtSql = "Update LOCAL_LOGS set PARENT_ID = " + nParentID.ToString() + " where ID = " + nLLogID.ToString();
            return UpdateSqlExecute(gibdd_con, txtSql);
        }

        private bool UpdateLLogParentStatus(OleDbConnection gibdd_con, decimal nLLogID, int Status)
        {
            //string txtSql = "Update DX_PACK set PLAIN_LOG = PLAIN_LOG || '" + txtText + "' where ID = " + nPackID.ToString();
            string txtSql = "Update LOCAL_LOGS set PACK_STATUS = " + Status.ToString() + " where ID = (select PARENT_ID from LOCAL_LOGS where ID = " + nLLogID.ToString() + ")";
            return UpdateSqlExecute(gibdd_con, txtSql);
        }

        private decimal GetLLogStatus(OleDbConnection gibdd_con, decimal nLLogID)
        {
            string txtSql = "SELECT  PACK_STATUS from LOCAL_LOGS where ID = " + nLLogID.ToString();
            string txtResult = "";
            txtResult = SelectSqlScalar(gibdd_con, txtSql);
            // �������� 
            if (txtResult.Trim() == "") return 0;
            return Convert.ToDecimal(txtResult);
        }

        private String SelectSqlScalar(OleDbConnection con, string txtSql)
        {
            String res = "";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand(txtSql, con, tran);
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
            

        private bool UpdateLLogFlag(OleDbConnection gibdd_con, decimal nLLogID, int Status, string txtFlagName)
        {
            string txtSql = "";

            switch (txtFlagName)
            {
                case "FL_FIND":
                    txtSql = "Update LOCAL_LOGS set FL_FIND = " + Status.ToString() + " where ID = (select first 1 parent_id from local_logs where id =" + nLLogID.ToString() + ")";
                    break;
                case "FL_NOFIND":
                    txtSql = "Update LOCAL_LOGS set FL_NOFIND = " + Status.ToString() + " where ID = (select first 1 parent_id from local_logs where id =" + nLLogID.ToString() + ")";
                    break;
                case "FL_E_TOFIND":
                    txtSql = "Update LOCAL_LOGS set FL_E_TOFIND = " + Status.ToString() + " where ID = (select first 1 parent_id from local_logs where id =" + nLLogID.ToString() + ")";
                    break;
            }

            if (txtSql != "")
                return UpdateSqlExecute(gibdd_con, txtSql);
            else return false;
        }
        
        private decimal CopyLLogParent(OleDbConnection gibdd_con, decimal nOldLLogID, string txtNewAgreementCode)
        {
            bool bUpdated = true;
            decimal nID = 0;
            OleDbCommand cmd, cmdIns;
            OleDbTransaction tran = null;

            try
            {
                if (gibdd_con != null && gibdd_con.State != ConnectionState.Closed) gibdd_con.Close();
                gibdd_con.Open();
                tran = gibdd_con.BeginTransaction(IsolationLevel.ReadCommitted);

                // �������� ����� ����
                cmd = new OleDbCommand("SELECT gen_id(GEN_LOCAL_LOG_ID, 1) FROM RDB$DATABASE", gibdd_con, tran);
                nID = Convert.ToDecimal(cmd.ExecuteScalar());

                if (nID * nOldLLogID != 0)
                {
                    cmdIns = new OleDbCommand();
                    cmdIns.Connection = gibdd_con;
                    cmdIns.Transaction = tran;
                    cmdIns.CommandText = "insert into LOCAL_LOGS (ID, OSPNUM, PACKDATE, PACK_TYPE, CONV_CODE, PACK_STATUS, PACK_COUNT, PARENT_ID, LOG) select :new_ID as ID, OSPNUM, PACKDATE, PACK_TYPE, :newAgrCode as CONV_CODE, PACK_STATUS, PACK_COUNT, PARENT_ID, LOG from LOCAL_LOGS WHERE ID = :old_ID";

                    cmdIns.Parameters.Add(new OleDbParameter(":new_ID", nID));
                    cmdIns.Parameters.Add(new OleDbParameter(":newAgrCode", txtNewAgreementCode));
                    cmdIns.Parameters.Add(new OleDbParameter(":old_ID", nOldLLogID));

                    if (cmdIns.ExecuteNonQuery() == -1)
                    {
                        bUpdated = false;
                    }
                }

                tran.Commit();
                gibdd_con.Close();

                if (!bUpdated)
                {
                    Exception ex = new Exception("Error. Can't copy doc_deposit id = " + nOldLLogID.ToString());
                        throw ex;
                }
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("������ ��� ������ � �������. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "��������!", MessageBoxButtons.OK);
                }
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (gibdd_con != null)
                {
                    gibdd_con.Close();
                }
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (gibdd_con != null)
                {
                    gibdd_con.Close();
                }
                MessageBox.Show("������ ����������. Message: " + ex.ToString(), "��������!", MessageBoxButtons.OK);
            }
            return nID;

        }

        private void ��������������������������������������ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // �������� ���� ������ � ���� ���������
            frmSelectDate SelectDate = new frmSelectDate();
            DatePeriod DateStartEnd;
            DateStartEnd = SelectDate.ShowForm();

            // ��������� ��� ������������ �� ����� Cancel  DateStartEnd.DateStart/End = DateTime.MinValue
            if (!((DateStartEnd.DateStart == DateTime.MinValue) && (DateStartEnd.DateEnd == DateTime.MinValue)))
            {
                DataTable dtReport = null;
                ReportMaker report = new ReportMaker();
                report.StartReport();
                report.AddToReport("<h3>");
                report.AddToReport("������ ����� ������ � ������������ ��������� � �������� � " + GetOSP_Name() + "<br />");
                report.AddToReport("</h3>");
                report.AddToReport("<p>���� ������������: " + DateTime.Now.ToString() + "</p>");
                string txtSql = "select ll.id, ll.packdate, agr.name_agreement, ll.pack_count, ls.status_name,  ll.pack_type, ll_resp.pack_type, pt.type as resp_type, ls_resp.status_name as resp_status, ll_resp.packdate as resp_date, ll_resp.pack_count as resp_count, ll.fl_find, ll.fl_nofind, ll.fl_e_tofind from local_logs ll join agreements agr on ll.conv_code = agr.agreement_code join logs_status ls on ll.pack_status = ls.id left join local_logs ll_resp on ll_resp.parent_id = ll.id left join pack_type pt on ll_resp.pack_type = pt.id left join logs_status ls_resp on ls_resp.id = ll_resp.pack_status and ll.packdate >= '" + DateStartEnd.DateStart.ToShortDateString() + "' and ll.packdate <='" + DateStartEnd.DateEnd.ToShortDateString() + "' where ll.pack_type = 1 order by ll.id";

                dtReport = GetDataTableFromFB(constrGIBDD, txtSql, "report", IsolationLevel.Unspecified);
                DateTime dPackDate, dRespDate;
                string txtPackDate, txtRespDate, txtAgr, txtReqCount, txtRespCount, txtReqStatus, txtRespStatus, txtRespType;
                Decimal nReqCount, nRespCount;

                if (dtReport != null)
                {
                    report.AddToReport("<table border=\"1\" cellpadding=\"2\" style=\"border:1px #000000 solid;\"><tbody>");

                    report.AddToReport("<tr>");
                    report.AddToReport("<td>���� �������� ��������</td>");
                    report.AddToReport("<td>��� ��������</td>");
                    report.AddToReport("<td>���������� ����������� ��������</td>");
                    report.AddToReport("<td>������ ������ ��������</td>");
                    report.AddToReport("<td>���� ��������� �������</td>");
                    report.AddToReport("<td>��� �������</td>");
                    report.AddToReport("<td>���������� ������������ �������</td>");
                    report.AddToReport("<td>������ ������ �������</td>");
                    report.AddToReport("</tr>");


                    foreach (DataRow row in dtReport.Rows)
                    {
                        txtPackDate = Convert.ToString(row["PACKDATE"]).Trim();
                        if (DateTime.TryParse(txtPackDate, out dPackDate))
                        {
                            txtPackDate = dPackDate.ToShortDateString();
                        }
                        else txtPackDate = "";

                        txtAgr = Convert.ToString(row["NAME_AGREEMENT"]).Trim();

                        txtReqCount = Convert.ToString(row["PACK_COUNT"]).Trim();
                        if (!Decimal.TryParse(txtReqCount, out nReqCount))
                        {
                            txtReqCount = "0";
                        }

                        txtReqStatus = Convert.ToString(row["STATUS_NAME"]).Trim();

                        txtRespDate = Convert.ToString(row["RESP_DATE"]).Trim();
                        if (DateTime.TryParse(txtRespDate, out dRespDate))
                        {
                            txtRespDate = dRespDate.ToShortDateString();
                        }
                        else txtRespDate = "";

                        txtRespType = Convert.ToString(row["RESP_TYPE"]).Trim();

                        txtRespCount = Convert.ToString(row["RESP_COUNT"]).Trim();
                        if (!Decimal.TryParse(txtRespCount, out nRespCount))
                        {
                            txtRespCount = "0";
                        }


                        txtRespStatus = Convert.ToString(row["RESP_STATUS"]).Trim();

                        report.AddToReport("<tr>");
                        report.AddToReport("<td>" + txtPackDate + "</td>");
                        report.AddToReport("<td>" + txtAgr + "</td>");
                        report.AddToReport("<td>" + txtReqCount + "</td>");
                        report.AddToReport("<td>" + txtReqStatus + "</td>");
                        report.AddToReport("<td>" + txtRespDate + "</td>");
                        report.AddToReport("<td>" + txtRespType + "</td>");
                        report.AddToReport("<td>" + txtRespCount + "</td>");
                        report.AddToReport("<td>" + txtRespStatus + "</td>");
                        report.AddToReport("</tr>");
                    }

                    report.AddToReport("</tbody></table>");
                    //report.SplitNewPage();
                }

                report.EndReport();
                report.ShowReport();
            }

        }

    }

}

