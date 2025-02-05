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
        String[] Legal_Сonv_List;

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
            // выбрать строки с настройками из app.config
            Properties.Settings s = new Properties.Settings();
            // подключаемся через прочитанную из настроек строку подключения
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

            //Загрузка листбокса для рееестров ГИБДД
            //load_gibdlist();

            //***Скрываю*GIBDD***
            //tabControl1.Controls.Remove(tabGibd);
            //tabControl1.Update();

            //***Скрываю*KRC***
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
                //tss = ConvertDOS("абвгдеёжз");
            }*/

            //OOo_Writer child = new OOo_Writer();
            //OooIsInstall = child.isOOoInstalled();
        }

        // функция получает на входе путь, возвращает пусть + текущая дата, заодно создавая саму папку с текущей датой
        private void CreatePathWithDate(string txtPathWithoutDate){

            string txtCurrDateFolder = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString().PadLeft(2, '0') + DateTime.Today.Day.ToString().PadLeft(2, '0');

            if (Directory.Exists(string.Format(@"{0}\{1}", txtPathWithoutDate, txtCurrDateFolder)))
            {
                    // Нужно сделать новый путь с суффиксом _mmss для release_name
                    DateTime dtFixNowDate = DateTime.Now;
                    string suffix = "_" + dtFixNowDate.Hour.ToString().PadLeft(2, '0') + dtFixNowDate.Minute.ToString().PadLeft(2, '0') + dtFixNowDate.Second.ToString().PadLeft(2, '0');
                    DialogResult rv = MessageBox.Show("По пути " + string.Format(@"{0}\{1}", txtPathWithoutDate, txtCurrDateFolder) + ", указанном в конфигурационном файле, существует файл. Будет файл будет выгружен в папку " + txtCurrDateFolder + suffix, "Внимание", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return tbl;
        }

        private DataTable GetDataTableFromFB(string txtSql, string tblName, IsolationLevel islLevel)
        {
            DataSet ds = new DataSet();
            DataTable tbl = ds.Tables.Add(tblName);
            try
            {
                // проверить подключение - а то может статься что не закрыли
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return tbl;
        }

        private DataTable GetDataTableFromFB(string txtSql, string tblName){
            DataSet ds = new DataSet();
            DataTable tbl = ds.Tables.Add(tblName);
            try
            {
                // проверить подключение - а то может статься что не закрыли
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
            string txtStatus = "отправлен";
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
                    txtKtfomsConv = "соглашением об обмене информацией по обязательному медицинскому страхованию от 1 ноября 2006 года";
                }

                // раскидываем txtKtfomsName на строчку одну не больше 120 символов и 2-ю не больше 200 символов

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
                        txtStatus = "отправлен";
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


                        //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// реально это дата регистрации ИП
                        //{
                        //    DatZapr = DateTime.Today;
                        //}
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                        // а вот тут надо строку анализировать

                        vid_d = 1; // физ. лицо
                        //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                        //{
                        //    vid_d = 1;// физ. лицо
                        //}

                        m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                        //m_cmd.Parameters.Add(new OleDbParameter(":INN_D", Convert.ToString(row["INNORG"])));
                        m_cmd.Parameters.Add(new OleDbParameter(":INN_D", ""));

                        m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 100)));

                        // TODO: сделать проверку на корректность года рождения и если ошибка - то писать ошибку в пксп
                        if (vid_d == 1) // проверка только для физ. лиц
                        {
                            nBirthYear = parseBirthDate(Convert.ToString(row["DATROZHD"]));
                            if (nBirthYear == 0)
                            {
                                nResult = 1;
                                txtStatus = "Ошибка в запросе";
                                txtText = "Отсутствует корректно введенная дата или год рождения (формат ##.##.#### или ####)";
                                bFizBadYear = true;
                            }
                        }
                        if (!DateTime.TryParse(Convert.ToString(row["DATROZHD"]), out dtDate))
                        {
                            dtDate = DateTime.MaxValue;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// это номер ответа от взаимод орг-ии

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// дата ответа, не забыть обновить при ответе

                        m_cmd.Parameters.Add(new OleDbParameter(":RESULT", nResult));

                        //m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ДОЛЖНИК НЕ ИДЕНТИФИЦИРОВАН, 1 - НЕТ ИНФ. ПО ДОЛЖНИКУ, БОЛЬШЕ 1 - ЕСТЬ ИНФ-Я ПО ДОЛЖНИКУ) (ВСЕ ЧТО В FIND - ВСЕ БОЛЬШЕ 1, ИЗНАЧАЛЬНО 0)

                        Int32 iKey = -1;
                        //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                        //{
                        //    iKey = 0;
                        //}

                        // TODO: ПРОВЕРЬ, ЧТО ЭТО РЕАЛЬНО ССЫЛКА НА DOCUMENTS!!!
                        // надо еще в таблицу DOCUMENTS вставлять запись, пока не когда
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

                        m_cmd.Parameters.Add(new OleDbParameter(":TEXT", txtText));// текст ответа по-умолчанию пустой
                        //m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// текст ответа - пустой

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", dat1));
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", dat2));

                        m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250)));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                        //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                        m_cmd.Parameters.Add(new OleDbParameter(":STATUS", txtStatus));
                        //m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "отправлен"));

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                        if (txtKtfomsConv2.Length > 0)
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", txtKtfomsConv2));
                        }
                        else
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", txtText)); // если нет ошибки то будет пустой, а если есть то будет text
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

                        // лезем в таблицу persons и там ищем нужные данные
                        // таблицу persons по ip.fk; в таблицу physical по person.tablename + person.FK
                        // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                        //

                        m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// потом напишу пасп. данные

                        Double sum = 0;
                        //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                        //{
                        //    sum = 0;
                        //}
                        m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                        // TODO: пишет что нет NOMIP в tofind
                        // теперь NOMIP можно взять из IPNO_NUM
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con.State != System.Data.ConnectionState.Closed) con.Close();
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
            string txtStatus = "отправлен";
            string txtText = "";
            int nResult = 0;
            int nBirthYear = 0;

            prbWritingDBF.Value = 0;
            if (DT_pens_doc != null)
            {
                // автоматом больше не выгребаем
                //prbWritingDBF.Maximum = DT_pens_reg.Rows.Count + DT_pens_doc.Rows.Count; // *Legal_List.Length;
                prbWritingDBF.Maximum = DT_pens_doc.Rows.Count;
            }
            // автоматом больше не выгребаем
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
                    txtPensConv = "соглашением № 10/03-08-1778 о порядке информационного обмена и взаимодействия между УФССП по РК и Отделением Пенсионного фонда РФ по РК от 24 марта 2008 года";
                }
                txtPensConv = cutEnd(txtPensConv, 120);

                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                #region "penss_reg"
                // автоматом больше не выгребаем
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


                //    //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// реально это дата регистрации ИП
                //    //{
                //    //    DatZapr = DateTime.Today;
                //    //}
                //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                //    // а вот тут надо строку анализировать

                //    vid_d = 1; // физ. лицо
                //    //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                //    //{
                //    //    vid_d = 1;// физ. лицо
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

                //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// это номер ответа от взаимод орг-ии

                //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// дата ответа, не забыть обновить при ответе

                //    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ДОЛЖНИК НЕ ИДЕНТИФИЦИРОВАН, 1 - НЕТ ИНФ. ПО ДОЛЖНИКУ, БОЛЬШЕ 1 - ЕСТЬ ИНФ-Я ПО ДОЛЖНИКУ) (ВСЕ ЧТО В FIND - ВСЕ БОЛЬШЕ 1, ИЗНАЧАЛЬНО 0)

                //    Int32 iKey = -1;
                //    //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                //    //{
                //    //    iKey = 0;
                //    //}

                //    // TODO: ПРОВЕРЬ, ЧТО ЭТО РЕАЛЬНО ССЫЛКА НА DOCUMENTS!!!
                //    // надо еще в таблицу DOCUMENTS вставлять запись, пока не когда
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

                //    m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// текст ответа - пустой

                //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", dat1));
                //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", dat2));

                //    m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250)));

                //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                //    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                //    m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "отправлен"));

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

                //    // лезем в таблицу persons и там ищем нужные данные
                //    // таблицу persons по ip.fk; в таблицу physical по person.tablename + person.FK
                //    // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                //    //

                //    m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// потом напишу пасп. данные

                //    Double sum = 0;
                //    //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                //    //{
                //    //    sum = 0;
                //    //}
                //    m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                //    // TODO: пишет что нет NOMIP в tofind
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
                        txtStatus = "отправлен";
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


                        //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// реально это дата регистрации ИП
                        //{
                        //    DatZapr = DateTime.Today;
                        //}
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                        // а вот тут надо строку анализировать

                        vid_d = 1; // физ. лицо
                        //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                        //{
                        //    vid_d = 1;// физ. лицо
                        //}

                        m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                        //m_cmd.Parameters.Add(new OleDbParameter(":INN_D", Convert.ToString(row["INNORG"])));
                        m_cmd.Parameters.Add(new OleDbParameter(":INN_D", ""));

                        m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 100)));

                        // TODO: сделать проверку на корректность года рождения и если ошибка - то писать ошибку в пксп
                        if (vid_d == 1) // проверка только для физ. лиц
                        {
                            nBirthYear = parseBirthDate(Convert.ToString(row["DATROZHD"]));
                            if (nBirthYear == 0)
                            {
                                nResult = 1; // а текст и так уже какой надо
                                txtStatus = "Ошибка в запросе";
                                txtText = "Отсутствует корректно введенная дата или год рождения (формат ##.##.#### или ####)";
                                bFizBadYear = true;
                            }
                        }

                        if (!DateTime.TryParse(Convert.ToString(row["DATROZHD"]), out dtDate))
                        {
                            dtDate = DateTime.MaxValue;
                        }

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// это номер ответа от взаимод орг-ии

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// дата ответа, не забыть обновить при ответе

                        m_cmd.Parameters.Add(new OleDbParameter(":RESULT", nResult));
                        //m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ДОЛЖНИК НЕ ИДЕНТИФИЦИРОВАН, 1 - НЕТ ИНФ. ПО ДОЛЖНИКУ, БОЛЬШЕ 1 - ЕСТЬ ИНФ-Я ПО ДОЛЖНИКУ) (ВСЕ ЧТО В FIND - ВСЕ БОЛЬШЕ 1, ИЗНАЧАЛЬНО 0)

                        Int32 iKey = -1;
                        //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                        //{
                        //    iKey = 0;
                        //}

                        // TODO: ПРОВЕРЬ, ЧТО ЭТО РЕАЛЬНО ССЫЛКА НА DOCUMENTS!!!
                        // надо еще в таблицу DOCUMENTS вставлять запись, пока не когда
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

                        m_cmd.Parameters.Add(new OleDbParameter(":TEXT", txtText));// текст ответа по-умолчанию пустой
                        //m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// текст ответа - пустой

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", DatZapr1_pens));
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", DatZapr2_pens));

                        m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250)));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                        //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                        m_cmd.Parameters.Add(new OleDbParameter(":STATUS", txtStatus));
                        //m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "отправлен"));

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                        m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", txtText)); // если нет ошибки то будет пустой, а если есть то будет text
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

                        // лезем в таблицу persons и там ищем нужные данные
                        // таблицу persons по ip.fk; в таблицу physical по person.tablename + person.FK
                        // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                        //

                        m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// потом напишу пасп. данные

                        Double sum = 0;
                        //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                        //{
                        //    sum = 0;
                        //}
                        m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                        // TODO: пишет что нет NOMIP в tofind
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
            string txtStatus = "отправлен";
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
                    txtPotdConv = "соглашением об обмене информацией по наличии получаемой должником пенсии от 27 марта 2008 г.";
                }
                
                if (con != null && con.State != ConnectionState.Closed) con.Close();

                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                #region "potd_doc"

                foreach (DataRow row in DT_potd_doc.Rows)
                {
                    bFizBadYear = false;
                    nResult = 0;
                    txtStatus = "отправлен";
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


                    //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// реально это дата регистрации ИП
                    //{
                    //    DatZapr = DateTime.Today;
                    //}
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                    // а вот тут надо строку анализировать

                    vid_d = 1; // физ. лицо
                    //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                    //{
                    //    vid_d = 1;// физ. лицо
                    //}

                    m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                    //m_cmd.Parameters.Add(new OleDbParameter(":INN_D", Convert.ToString(row["INNORG"])));
                    m_cmd.Parameters.Add(new OleDbParameter(":INN_D", ""));

                    m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 100)));

                    if (vid_d == 1) // проверка только для физ. лиц
                    {
                        nBirthYear = parseBirthDate(Convert.ToString(row["DATROZHD"]));
                        if (nBirthYear == 0)
                        {
                            nResult = 1; // а текст и так уже какой надо
                            txtStatus = "Ошибка в запросе";
                            txtText = "Отсутствует корректно введенная дата или год рождения (формат ##.##.#### или ####)";
                            bFizBadYear = true;
                        }
                    }

                    if (!DateTime.TryParse(Convert.ToString(row["DATROZHD"]), out dtDate))
                    {
                        dtDate = DateTime.MaxValue;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// это номер ответа от взаимод орг-ии

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// дата ответа, не забыть обновить при ответе

                    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", nResult));
                    //m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ДОЛЖНИК НЕ ИДЕНТИФИЦИРОВАН, 1 - НЕТ ИНФ. ПО ДОЛЖНИКУ, БОЛЬШЕ 1 - ЕСТЬ ИНФ-Я ПО ДОЛЖНИКУ) (ВСЕ ЧТО В FIND - ВСЕ БОЛЬШЕ 1, ИЗНАЧАЛЬНО 0)

                    Int32 iKey = -1;
                    //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                    //{
                    //    iKey = 0;
                    //}

                    // TODO: ПРОВЕРЬ, ЧТО ЭТО РЕАЛЬНО ССЫЛКА НА DOCUMENTS!!!
                    // надо еще в таблицу DOCUMENTS вставлять запись, пока не когда
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

                    m_cmd.Parameters.Add(new OleDbParameter(":TEXT", txtText));// текст ответа по-умолчанию пустой
                    //m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// текст ответа - пустой

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", dat1));
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", dat2));

                    m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250)));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                    m_cmd.Parameters.Add(new OleDbParameter(":STATUS", txtStatus));
                    //m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "отправлен"));

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                    m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", txtText)); // если нет ошибки то будет пустой, а если есть то будет text
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

                    // лезем в таблицу persons и там ищем нужные данные
                    // таблицу persons по ip.fk; в таблицу physical по person.tablename + person.FK
                    // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                    //

                    m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// потом напишу пасп. данные

                    Double sum = 0;
                    //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                    //{
                    //    sum = 0;
                    //}
                    m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                    // TODO: пишет что нет NOMIP в tofind
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    txtKrcConv = "соглашением об обмене информацией о суммах, оплаченных должниками по исполнительным документам.";
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


                    //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// реально это дата регистрации ИП
                    //{
                    //    DatZapr = DateTime.Today;
                    //}
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                    // а вот тут надо строку анализировать

                    vid_d = 1; // физ. лицо
                    //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                    //{
                    //    vid_d = 1;// физ. лицо
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

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// это номер ответа от взаимод орг-ии

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// дата ответа, не забыть обновить при ответе

                    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ДОЛЖНИК НЕ ИДЕНТИФИЦИРОВАН, 1 - НЕТ ИНФ. ПО ДОЛЖНИКУ, БОЛЬШЕ 1 - ЕСТЬ ИНФ-Я ПО ДОЛЖНИКУ) (ВСЕ ЧТО В FIND - ВСЕ БОЛЬШЕ 1, ИЗНАЧАЛЬНО 0)

                    Int32 iKey = -1;
                    //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                    //{
                    //    iKey = 0;
                    //}

                    // TODO: ПРОВЕРЬ, ЧТО ЭТО РЕАЛЬНО ССЫЛКА НА DOCUMENTS!!!
                    // надо еще в таблицу DOCUMENTS вставлять запись, пока не когда
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

                    m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// текст ответа - пустой

                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", dat1));
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", dat2));

                    m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250)));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                    m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "отправлен"));

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

                    // лезем в таблицу persons и там ищем нужные данные
                    // таблицу persons по ip.fk; в таблицу physical по person.tablename + person.FK
                    // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                    //

                    m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// потом напишу пасп. данные

                    Double sum = 0;
                    //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                    //{
                    //    sum = 0;
                    //}
                    m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                    // TODO: пишет что нет NOMIP в tofind
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }

            return iCnt;
        }

        //private int InsertZapros()
        private int InsertZapros(DataTable dtTable, DateTime DatZapr1_param, DateTime DatZapr2_param)
        {
            // главное что тут надо понять - что запрос реально вставляется только для Сбербанка - если он первый в Legal_List
            int vid_d = 1;
            DateTime dtDate;
            DateTime DatZapr;
            String uid = System.Guid.NewGuid().ToString();
            String osp_name = GetOSP_Name();
            Decimal osp_num = GetOSP_Num();
            String osp_h_pristav = GetOSP_H_Pristav();
            String legal_branch = GetLegal_Branch(Convert.ToInt32(Legal_List[0].Trim())); // наверное branch это номер
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
                    
                    string txtStatus = "Запрос готов к отправке";
                    
                    txtTextError = "";
                    
                    if (Legal_Сonv_List[0].Length > 100)
                    {
                        txtTextError = Legal_Сonv_List[0].Substring(100);
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


                    if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// реально это дата регистрации ИП
                    {
                        DatZapr = DateTime.Today;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DatZapr));

                    // а вот тут надо строку анализировать

                    vid_d = 2;
                    if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                    {
                        vid_d = 1;// физ. лицо
                    }

                    m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                    m_cmd.Parameters.Add(new OleDbParameter(":INN_D", Convert.ToString(row["INNORG"])));

                    // сомотрим есть ли фамилия, имя и отчество
                    // если чего-то нет, то пишем что staus = 'Ответ' и text = 'Не удалось определить фамилию, имя и отчество должника. Электронный запрос не может быть обработан Сбербанком. Необходимо направить запрос в бумажном виде.'
                    string txtFIO = Convert.ToString(row["FIOVK"]).Trim();
                    string[] Names;
                    Names = parseFIO(txtFIO);
                    if (!(Names.Length > 2) || (Names[0].Trim().Equals("")) || (Names[1].Trim().Equals("")) || (Names[2].Trim().Equals("")))
                    {
                        txtStatus = "Ответ";
                        bNotFullFIO = true;
                        txtZaprosText = "'Не удалось определить фамилию или имя или отчество должника. Электронный запрос не может быть обработан Сбербанком. Необходимо направить запрос в бумажном виде.'";
                    }

                    m_cmd.CommandText += ":NAME_D, :DATE_BORN, NULL, NULL, NULL, :FK_DOC, :FK_IP, :FK_ID, NULL, :NUM_ID," + txtZaprosText + ", NULL, NULL, NULL,";
                    
                    m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", cutEnd(txtFIO, 100)));

                    nBirthYear = parseBirthDate(Convert.ToString(row["GOD"]));
                    if (nBirthYear == 0)
                    {
                        txtStatus = "Ошибка в запросе";
                        txtTextError = "Отсутствует дата рождения (формат ##.##.####)";
                    }

                    if (!DateTime.TryParse(Convert.ToString(row["GOD"]), out dtDate))
                    {
                        dtDate = DateTime.MaxValue;
                    }
                    m_cmd.Parameters.Add(new OleDbParameter(":DATE_BORN", dtDate));

                    // теперь следующие 3 не нужны
                    //m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// это номер ответа от взаимод орг-ии

                    //m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// дата ответа, не забыть обновить при ответе

                    //3.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ДОЛЖНИК НЕ ИДЕНТИФИЦИРОВАН, 1 - НЕТ ИНФ. ПО ДОЛЖНИКУ, БОЛЬШЕ 1 - ЕСТЬ ИНФ-Я ПО ДОЛЖНИКУ) (ВСЕ ЧТО В FIND - ВСЕ БОЛЬШЕ 1, ИЗНАЧАЛЬНО 0)

                    Int32 iKey = -1;
                    //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                    //{
                    //    iKey = 0;
                    //}

                    // TODO: ПРОВЕРЬ, ЧТО ЭТО РЕАЛЬНО ССЫЛКА НА DOCUMENTS!!!
                    // надо еще в таблицу DOCUMENTS вставлять запись, пока не когда
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

                    //m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// текст ответа - пустой
                    
                    m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", cutEnd(Convert.ToString(Legal_Name_List[0]).Trim(), 200)));
                    //m_cmd.Parameters[":ADRESAT"].Value = Convert.ToString(Legal_Name_List[i]);

                    if (Legal_Сonv_List[0].Length > 100)
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
                    // проверить где заполняется Legal_NameList[0];
                    
                    
                    if (Legal_Сonv_List[0].Length > 100)
                    {
                        m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_Сonv_List[0].Substring(0, 100)));
                    }
                    else
                    {
                        m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_Сonv_List[0].Substring(0, Legal_Сonv_List[0].Length)));
                    }

                    
                    // лезем в таблицу persons и там ищем нужные данные
                    // таблицу persons по ip.fk; в таблицу physical по person.tablename + person.FK
                    // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                    //

                    m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// потом напишу пасп. данные
                    
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

                    // надо вытащить из FIOVK отдельно Ф, И и О
                    // 

                    // это все мы уже выше сделали
                    //string txtFIO = Convert.ToString(row["FIOVK"]).Trim();
                    //string[] Names;
                    //Names = parseFIO(txtFIO);

                    if (Names.Length > 0) m_cmd.Parameters.Add(new OleDbParameter(":FAM", cutEnd(Convert.ToString(Names[0]), 30)));
                    else m_cmd.Parameters.Add(new OleDbParameter(":FAM", ""));

                    if (Names.Length > 1) m_cmd.Parameters.Add(new OleDbParameter(":IM", cutEnd(Convert.ToString(Names[1]), 30)));
                    else m_cmd.Parameters.Add(new OleDbParameter(":IM", ""));

                    if (Names.Length > 2)
                    {
                        // все что осталось - отчетсво. склеиваем, обрезаем до 30 символов и в базу
                        string txtOt = "";
                        
                        for (int j = 2; j < Names.Length; j++)
                        {
                            txtOt += Names[j] + ' ';
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":OT", cutEnd(txtOt.TrimEnd(), 30)));

                        //if (Names.Length > 3)
                        //{
                        //    MessageBox.Show("Отчество необычное очень! Message:" + txtOt, "Внимание!", MessageBoxButtons.OK);

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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    string txtStatus = "отправлен";
                    bJurBadInn = false;
                    bFizBadYear = false;
                    int nBirthYear = 0;
                    
                    for (int i = 1; i < Legal_List.Length; i++)
                    {
                        // для doc тут надо будет проверить есть ли ИНН у юр. лица и если нет то
                        // изменить статус, result и text
                        // при выгрузке в DBF повторить проверку так-же
                        // в Insert добавить поля TEXT, RESULT, (STATUS уже есть)
                        // NULL, 0 по-умолчанию

                        txtText = "";
                        nResult = 0;
                        nBirthYear = 0;
                        txtStatus = "отправлен";
                        
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


                        if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// реально это дата регистрации ИП
                        {
                            DatZapr = DateTime.Today;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DatZapr));

                        // а вот тут надо строку анализировать

                        vid_d = 2;
                        txtLitzdolg = Convert.ToString(row["LITZDOLG"]).Trim();
                        if (txtLitzdolg.StartsWith("/1/"))
                        {
                            vid_d = 1;// физ. лицо
                        }

                        m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                        txtInnd = Convert.ToString(row["INNORG"]).Trim();
                        // если это REG и юр. лицо
                        if (bRegTable && (vid_d == 2))
                        {
                            if (txtInnd.Trim().Length < 10)
                            {
                                // вот тут и делаем все что нужно чтобы сообщить приставу о том что у него беда с ИНН и запрос не уйдет
                                nResult = 1; // а текст и так уже какой надо
                                txtText = "Ошибка! У должника - юр.лица не заполнен ИНН. Электронный запрос не может быть отправлен. Для отправки электронного запроса необходимо заполнить поле ИНН и создать запрос снова.\n";
                                txtStatus = "Ошибка в запросе";
                                bJurBadInn = true;
                            }
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":INN_D", txtInnd));

                        txtName_d = cutEnd(Convert.ToString(row["FIOVK"]).Trim(), 100);
                        m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", txtName_d));

                        // TODO: сделать проверку на корректность года рождения и если ошибка - то писать ошибку в пксп
                        if (vid_d == 1) // проверка только для физ. лиц
                        {
                            nBirthYear = parseBirthDate(Convert.ToString(row["GOD"]));
                            if (nBirthYear == 0)
                            {
                                nResult = 1; // а текст и так уже какой надо
                                txtStatus = "Ошибка в запросе";
                                txtText = "Отсутствует корректно введенная дата или год рождения (формат ##.##.#### или ####)";
                                bFizBadYear = true;
                            }
                        }

                        if (!DateTime.TryParse(Convert.ToString(row["GOD"]), out dtDate))
                        {
                            dtDate = DateTime.MaxValue;
                        }

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// это номер ответа от взаимод орг-ии

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// дата ответа, не забыть обновить при ответе

                        // m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ДОЛЖНИК НЕ ИДЕНТИФИЦИРОВАН, 1 - НЕТ ИНФ. ПО ДОЛЖНИКУ, БОЛЬШЕ 1 - ЕСТЬ ИНФ-Я ПО ДОЛЖНИКУ) (ВСЕ ЧТО В FIND - ВСЕ БОЛЬШЕ 1, ИЗНАЧАЛЬНО 0)

                        m_cmd.Parameters.Add(new OleDbParameter(":RESULT", nResult));

                        //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                        //{
                        //    iKey = 0;
                        //}

                        // TODO: ПРОВЕРЬ, ЧТО ЭТО РЕАЛЬНО ССЫЛКА НА DOCUMENTS!!!
                        // надо еще в таблицу DOCUMENTS вставлять запись, пока не когда

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

                        m_cmd.Parameters.Add(new OleDbParameter(":TEXT", txtText));// текст ответа по-умолчанию пустой

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", DatZapr1_param));
                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", DatZapr2_param));

                        txtAddr = cutEnd(Convert.ToString(row["ADDR"]).Trim(), 250);
                        m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", txtAddr));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", uid));

                        m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                        //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                        m_cmd.Parameters.Add(new OleDbParameter(":STATUS", txtStatus));

                        m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                        if (Legal_Сonv_List[i].Length > 100)
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", Legal_Сonv_List[i].Substring(100)));
                        }
                        else
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", txtText)); // если нет ошибки то будет пустой, а если есть то будет text
                        }

                        if (!(Int32.TryParse(Convert.ToString(row["USCODE"]), out iUSCODE)))
                        {
                            iUSCODE = 0;
                        }
                        m_cmd.Parameters.Add(new OleDbParameter(":USCODE", iUSCODE));

                        //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));

                        if (Legal_Сonv_List[i].Length > 100)
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_Сonv_List[i].Substring(0, 100)));
                        }
                        else
                        {
                            m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_Сonv_List[i].Substring(0, Legal_Сonv_List[i].Length)));
                        }

                        m_cmd.Parameters.Add(new OleDbParameter(":WHYRESPONS", ""));

                        m_cmd.Parameters.Add(new OleDbParameter(":WHYPREPARE", ""));

                        // лезем в таблицу persons и там ищем нужные данные
                        // таблицу persons по ip.fk; в таблицу physical по person.tablename + person.FK
                        // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                        //

                        m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// потом напишу пасп. данные

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

                        // вставить запрос по частным предпринимателям как по юр.лицам если есть ИНН
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

                            // а вот тут надо строку анализировать

                            vid_d = 2; // запрос о ЧП = как юр.лице
                            m_cmd.Parameters.Add(new OleDbParameter(":VID_D", vid_d));

                            m_cmd.Parameters.Add(new OleDbParameter(":INN_D", txtInnd));

                            m_cmd.Parameters.Add(new OleDbParameter(":NAME_D", txtName_d));

                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_R", dtDate));

                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// это номер ответа от взаимод орг-ии

                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DatZapr));// дата ответа, не забыть обновить при ответе

                            m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));// (0 - ДОЛЖНИК НЕ ИДЕНТИФИЦИРОВАН, 1 - НЕТ ИНФ. ПО ДОЛЖНИКУ, БОЛЬШЕ 1 - ЕСТЬ ИНФ-Я ПО ДОЛЖНИКУ) (ВСЕ ЧТО В FIND - ВСЕ БОЛЬШЕ 1, ИЗНАЧАЛЬНО 0)

                            m_cmd.Parameters.Add(new OleDbParameter(":FK_DOC", iFK_DOC));

                            m_cmd.Parameters.Add(new OleDbParameter(":FK_IP", iFK_IP));

                            m_cmd.Parameters.Add(new OleDbParameter(":FK_ID", iFK_ID));

                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_ID", txtNum_Id));

                            m_cmd.Parameters.Add(new OleDbParameter(":TEXT", ""));// текст ответа - пустой

                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", DatZapr1_param));
                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", DatZapr2_param));

                            m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", txtAddr));

                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", uid));

                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                            //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                            m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "отправлен"));

                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_SEND", DateTime.Today));

                            if (Legal_Сonv_List[i].Length > 100)
                            {
                                m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", Legal_Сonv_List[i].Substring(100)));
                            }
                            else
                            {
                                m_cmd.Parameters.Add(new OleDbParameter(":TEXT_ERROR", ""));
                            }

                            m_cmd.Parameters.Add(new OleDbParameter(":USCODE", iUSCODE));

                            //m_cmd.Parameters.Add(new OleDbParameter(":FK_LEGAL", Convert.ToInt32(Legal_List[0].Trim())));

                            if (Legal_Сonv_List[i].Length > 100)
                            {
                                m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_Сonv_List[i].Substring(0, 100)));
                            }
                            else
                            {
                                m_cmd.Parameters.Add(new OleDbParameter(":CONVENTION", Legal_Сonv_List[i].Substring(0, Legal_Сonv_List[i].Length)));
                            }

                            m_cmd.Parameters.Add(new OleDbParameter(":WHYRESPONS", ""));

                            m_cmd.Parameters.Add(new OleDbParameter(":WHYPREPARE", ""));

                            // лезем в таблицу persons и там ищем нужные данные
                            // таблицу persons по ip.fk; в таблицу physical по person.tablename + person.FK
                            // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                            //

                            m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// потом напишу пасп. данные

                            
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }

            return iCnt;
        }




        private void ReadPOTDData(DateTime dat1, DateTime dat2)
        {
            // выгребаем неоконченные возбужденные за период или с вынесенными в заданный период постановлениями об ограничении выезда
            //DT_ktfoms = GetDataTableFromFB("SELECT DISTINCT a.USCODE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD,  a.name_d as FIOVK, a.DATE_BORN_D as DATROZHD, a.ADR_D as ADDR, b.PK as FK_DOC, a.PK as FK_IP, a.PK_ID as FK_ID, b.KOD, b.DATE_DOC FROM IP a LEFT JOIN DOCUMENT b ON b.FK = a.PK WHERE a.DATE_IP_OUT is not null and a.text_pp is not null and ((a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') or (b.DATE_DOC >= '" + dat1.ToShortDateString() + "' AND b.DATE_DOC <= '" + dat2.ToShortDateString() + "' AND b.KOD = 1010))  AND a.VIDD_KEY LIKE '/1/%'", "TOFIND");
            //DT_pens_reg = GetDataTableFromFB("SELECT DISTINCT a.USCODE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD, a.name_d as FIOVK, a.DATE_BORN_D as DATROZHD, a.ADR_D as ADDR, a.PK as FK_IP, a.PK_ID as FK_ID FROM IP a LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.DATE_IP_OUT is null and b.KOD != 1006 and (a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' AND a.NUM_IP NOT LIKE '%!%'", "TOFIND and a.NUM_IP not in (select a.NUM_IP from IP a LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.DATE_IP_OUT is null and b.KOD = 1006 and (b.DATE_DOC >= '" + dat1.ToShortDateString() + "' AND b.DATE_DOC <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%')");
            DT_potd_doc = GetDataTableFromFB("SELECT DISTINCT c.PRIMARY_SITE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD, UPPER(a.name_d) as FIOVK, a.DATE_BORN_D as DATROZHD, a.ADR_D as ADDR, b.PK as FK_DOC, a.PK as FK_IP, a.PK_ID as FK_ID, a.uscode, a.FIO_SPI FROM IP a left join s_users c on (a.uscode=c.uscode) LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.DATE_IP_OUT is null and b.KOD = 1045 and (b.DATE_DOC >= '" + dat1.ToShortDateString() + "' AND b.DATE_DOC <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' and a.ssd is null and a.ssv is null order by FIOVK", "TOFIND");
        }

        private void ReadKRCData(DateTime dat1, DateTime dat2)
        {            
            //DT_krc_reg = GetDataTableFromFB("select a.sdc as nomotd, (select first 1 c.PRIMARY_SITE from s_users c where c.uscode = a.uscode) as nomspi, a.num_ip as zapros, a.num_id as nomid, a.date_id_send as datid, a.name_d as namedol, a.DATE_BORN_D as born, a.sum_ as sumvz, a.adr_d as addr from ip a WHERE a.DATE_IP_OUT is null and a.name_v like '%КРЦ%' and (a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' and a.ssd is null and a.ssv is null", "TOFIND");
            //DT_krc_reg = GetDataTableFromFB("select distinct c.PRIMARY_SITE as NOMSPI, a.sdc as nomotd, a.num_ip as zapros, a.num_id as nomid, a.date_id_send as datid, a.name_d as namedol, a.DATE_BORN_D as born, a.sum_ as sumvz, a.adr_d as addr, a.PK as FK_IP, a.PK_ID as FK_ID, a.uscode, a.FIO_SPI from ip a left join s_users c on (a.uscode=c.uscode) WHERE a.DATE_IP_OUT is null and a.name_v like '%КРЦ%' and (a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' and a.ssd is null and a.ssv is null", "TOFIND");

            DT_krc_reg = GetDataTableFromFB("select distinct c.PRIMARY_SITE as NOMSPI, a.sdc as nomotd, a.num_ip as zapros, a.num_id as nomid, a.date_id_send as datid, a.name_d as namedol, a.DATE_BORN_D as born, a.sum_ as sumvz, a.adr_d as addr, a.PK as FK_IP, a.PK_ID as FK_ID, a.uscode, a.FIO_SPI from ip a left join s_users c on (a.uscode=c.uscode) WHERE a.DATE_IP_OUT is null and (a.name_v like '%КРЦ%' or a.name_v = 'Комплексный Расчетный Центр ООО') and a.VIDD_KEY LIKE '/1/%' and a.ssd is null and a.ssv is null", "TOFIND");
        }

        private bool ReestrOutWord(DataTable dtReg, string dir)
        {
            Decimal nYear = DateTime.Today.Year;
            DateTime dtDate;
            string bankname = "Карельское ОСБ N 8628 АК СБ РФ";
            //string bankadres = "г.Петрозаводск, ул.Антикайнена, д.2";
            string bankadres = "";
            string ospadres = GetOSP_Adres().ToUpper();
            string ospname = GetOSP_Name().ToUpper();

            DataRow[] FizRows = dtReg.Select("LITZDOLG LIKE '/1/*'", "FIOVK");

            if (File.Exists(string.Format(@"{0}\{1}", dir, DateTime.Today.ToShortDateString() + ".doc")))
                File.Delete(string.Format(@"{0}\{1}", dir, DateTime.Today.ToShortDateString() + ".doc"));

            using (StreamWriter sw = new StreamWriter(dir + "\\" + DateTime.Today.ToShortDateString() + ".doc", true, Encoding.GetEncoding(1251)))
            {

                sw.WriteLine("              МИНЮСТ                      " + bankname);
                sw.WriteLine("   ФЕДЕРАЛЬНАЯ СЛУЖБА СУДЕБНЫХ ПРИСТАВОВ  " + bankadres);
                sw.WriteLine("   УПРАВЛЕНИЕ ФССП ПО РЕСПУБЛИКЕ КАРЕЛИЯ  ");
                sw.WriteLine("");
                sw.WriteLine("  " + ospname);
                sw.WriteLine("  ");
                sw.WriteLine("  " + ospadres);
                sw.WriteLine("");
                sw.WriteLine("    Исх.N ________от _________   ;");
                sw.WriteLine("");
                sw.WriteLine("                                З А П Р О С");
                sw.WriteLine("   На исполнении в " + ospname);
                sw.WriteLine("   находятся исполнительные документы на должников : ");
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
                sw.WriteLine("  Просим Вас в семидневный срок сообщить о счетах должников,");
                sw.WriteLine("  зарегистрированных в Вашем банке.");
                sw.WriteLine("  ");
                sw.WriteLine("  ");
                sw.WriteLine("  Исполнитель:  ");

                sw.Flush();
                sw.Close();
            }

            using (StreamReader sr = new StreamReader(dir + "\\" + DateTime.Today.ToShortDateString() + ".doc", Encoding.GetEncoding(1251)))
            {

                // пример для Ворда

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
            string bankname = "Карельское ОСБ N 8628 АК СБ РФ";
            //string bankadres = "г.Петрозаводск, ул.Антикайнена, д.2";
            string bankadres = "";
            string ospadres = GetOSP_Adres().ToUpper();
            string ospname = GetOSP_Name().ToUpper();

            DataRow[] FizRows = dtReg.Select("LITZDOLG LIKE '/1/*'", "FIOVK");

            if (File.Exists(string.Format(@"{0}\{1}", dir, DateTime.Today.ToShortDateString() + ".doc")))
                File.Delete(string.Format(@"{0}\{1}", dir, DateTime.Today.ToShortDateString() + ".doc"));

            using (StreamWriter sw = new StreamWriter(dir + "\\" + DateTime.Today.ToShortDateString() + ".doc", true, Encoding.GetEncoding(1251)))
            {

                sw.WriteLine("              МИНЮСТ                      " + bankname);
                sw.WriteLine("   ФЕДЕРАЛЬНАЯ СЛУЖБА СУДЕБНЫХ ПРИСТАВОВ  " + bankadres);
                sw.WriteLine("   УПРАВЛЕНИЕ ФССП ПО РЕСПУБЛИКЕ КАРЕЛИЯ  ");
                sw.WriteLine("");
                sw.WriteLine("  " + ospname);
                sw.WriteLine("  ");
                sw.WriteLine("  " + ospadres);
                sw.WriteLine("");
                sw.WriteLine("    Исх.N ________от _________   ;");
                sw.WriteLine("");
                sw.WriteLine("                                З А П Р О С");
                sw.WriteLine("   На исполнении в " + ospname);
                sw.WriteLine("   находятся исполнительные документы на должников : ");
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
                sw.WriteLine("  Просим Вас в семидневный срок сообщить о счетах должников,");
                sw.WriteLine("  зарегистрированных в Вашем банке.");
                sw.WriteLine("  ");
                sw.WriteLine("  ");
                sw.WriteLine("  Исполнитель:  ");

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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", System.Windows.Forms.MessageBoxButtons.OK);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", System.Windows.Forms.MessageBoxButtons.OK);
            }
            return res;
        }

        private string PKOSP_GetOrgConvention(decimal org_id){
            string res = "< нет значения в базе данных >";
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return res;

        }

        private String PK_OSP_GetSPI_Name(Decimal code)
        {
            String res = "нет значения в базе данных";
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                // переписать select по-новому
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                // переписать select по-новому
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return res;
        }



        private String GetLegal_Name(int code)
        {
            String res = "нет значения в базе данных";
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return res;
        }


        private String GetLegal_Conv(int code)
        {
            String res = "нет значения в базе данных";
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return res;
        }


        private String GetLegal_Name(Decimal code)
        {
            String res = "нет значения в базе данных";
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return res;
        }

        private string GetLegalNameByAgrCode(string txtAgreementCode){
            String res = "нет значения в базе данных";
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return res;
        }
        


        private String GetLegal_Adr(Decimal code)
        {
            String res = "нет значения в базе данных";
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return res;
        }
        
        private String GetLegal_Conv(Decimal code)
        {
            String res = "нет значения в базе данных";
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return res;
        }


        
        
        private string[] parseFIO(string txtFIO)
        {
            string[] Names;

            try
            {
                // надо вытащить из FIOVK отдельно Ф, И и О
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
                MessageBox.Show("Ошибка при попытке разбить строку ФИО на 3 части. Message: " + ex.Message + "Source: " + ex.Source, "Внимание!", MessageBoxButtons.OK);
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
                    // вобще говоря, если строчка в нормальное число не транслируется, то вырезаем все точки 
                    // и если получилось 4 цифры и они попали в интервал 1900 - 9999
                    
                    // вырезать из строки все точки, оставшеееся рассматривать как год
                    string[] strData = txtDateBornD.Split('.');
                    txtDateBornD = "";
                    foreach (string str in strData)
                    {
                        txtDateBornD += str.Trim();
                    }

                    // вырезаем нули с начала строки - так как если это год рождения, то они нам не нужны - они незначащие
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
                MessageBox.Show("Ошибка при попытке определить год рождения должника. Message: " + ex.Message + "Source: " + ex.Source, "Внимание!", MessageBoxButtons.OK);
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

        //        // вобще говоря, если строчка в нормальное число не транслируется, то вырезаем все точки 
        //        // и если получилось 4 цифры и они попали в интервал 1900 - 9999

        //        // бывает, что год рождения помещается в конец строки, которая начинается с "  .  ."


        //        if (txtDateBornD.Length == 10)
        //        {
        //            if (txtDateBornD.Substring(0, 6) == "  .  .")
        //            {
        //                txtDateBornD = txtDateBornD.Substring(6, 4);
        //            }
        //        }
        //        else
        //        {
        //            // вырезать из строки все точки, оставшеееся рассматривать как год
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
        //        MessageBox.Show("Ошибка при попытке определить год рождения должника. Message: " + ex.Message + "Source: " + ex.Source, "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }

                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }

                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                // TODO: почему-то здесь зациклевается
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("По указанному пути в конфигурационном файле отсутствует директория для ведения архива. В эту папку файл сохранён не будет.", "Внимание", MessageBoxButtons.OK);
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

                        // если есть tofind  в корне, то удалить его
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
                        // вот тут надо будет перекодировать!!!
                        Process proc = new Process();
                        proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");
                        //proc.StartInfo.FileName = string.Format(@"{0}\{1}", "C:\\Program Files\\SSP\\InstallInfoChange", "fox622.exe ");

                        proc.StartInfo.Arguments = string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), tofind_name) + " " + string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), release_name);
                        proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";

                        proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                        proc.Start();

                        System.Threading.Thread.Sleep(5000);// ждем 5 секунд чтобы гарантированно выполнилось преобразование

                        DateTime tm;
                        tm = DateTime.Now;
                        while (!File.Exists(string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), release_name)) || (File.GetLastWriteTime(string.Format(@"{0}\{1}\{2}", m_fullpath, i.ToString(), release_name)).AddMilliseconds(100) > tm)) // пока не появился сконвертированный файл
                        {
                            System.Threading.Thread.Sleep(1000);// ждем секунду чтобы гарантированно выполнилось преобразование
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
                    // чистим все оставшиеся директориии от предыдущих раз
                    while (Directory.Exists(string.Format(@"{0}\{1}", m_fullpath, i.ToString())))
                    {
                        Directory.Delete(string.Format(@"{0}\{1}", m_fullpath, i.ToString()), true);
                        i++;
                    }


                }
                else
                {
                    // чистим все оставшиеся директориии от предыдущих раз
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

                    // вот тут надо будет перекодировать!!!
                    Process proc = new Process();
                    proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");
                    //proc.StartInfo.FileName = string.Format(@"{0}\{1}", "C:\\Program Files\\SSP\\InstallInfoChange", "fox622.exe ");
                    proc.StartInfo.Arguments = string.Format(@"{0}\{1}", m_fullpath, tofind_name) + " " + string.Format(@"{0}\{1}", m_fullpath, release_name);
                    proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";
                    proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                    proc.Start();

                    System.Threading.Thread.Sleep(5000);// ждем 5 секунд чтобы гарантированно выполнилось преобразование

                    DateTime tm;
                    tm = DateTime.Now;
                    while (!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)) || (File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name)).AddMilliseconds(100) > tm)) // пока не появился сконвертированный файл
                    {
                        System.Threading.Thread.Sleep(1000);// ждем секунду чтобы гарантированно выполнилось преобразование
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Win32Exception e)
            {
                if (e.NativeErrorCode == ERROR_FILE_NOT_FOUND)
                {
                    MessageBox.Show("Ошибка приложения. Проверьте путь доступа к файлу: " + e.Message, "Внимание!", MessageBoxButtons.OK);
                }

                else if (e.NativeErrorCode == ERROR_ACCESS_DENIED)
                {
                    MessageBox.Show("Ошибка приложения. Доступ к файлу запрещен: " + e.Message, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

        //    // забиваем на Архив
        //    //archive_folder_tofind = archive_sber_path;
        //    //if (!Directory.Exists(archive_folder_tofind))
        //    //{
        //    //    MessageBox.Show("По указанному пути в конфигурационном файле отсутствует директория для ведения архива. В эту папку файл сохранён не будет.", "Внимание", MessageBoxButtons.OK);
        //    //    archive_folder_tofind = "";
        //    //}


        //    try
        //    {
        //        // почистить каталог рекурсивно
        //        if (Directory.Exists(string.Format(@"{0}", m_fullpath)))                    
        //            Directory.Delete(string.Format(@"{0}", m_fullpath),true);
                
        //        // создать каталог на сервере
        //        Directory.CreateDirectory(string.Format(@"{0}", m_fullpath));
                
        //        for (i = 0; (i <= col_files); i++)
        //        {
        //            // TODO: вот тут и будем формировать новое имя файла - правильное.
        //            // release_name - сюда записать новое имя файла.

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
        //            // вот тут надо будет перекодировать!!!
        //            Process proc = new Process();
        //            proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");

        //            proc.StartInfo.Arguments = string.Format(@"{0}\{1}", m_fullpath, tofind_name) + " " + string.Format(@"{0}\{1}", m_fullpath, release_name);
        //            proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";
                    
        //            proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
        //            proc.Start();

        //            System.Threading.Thread.Sleep(5000);// ждем 5 секунд чтобы гарантированно выполнилось преобразование

        //            DateTime tm;
        //            tm = DateTime.Now;
        //            while (!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)) || (File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name)).AddMilliseconds(100) > tm)) // пока не появился сконвертированный файл
        //            {
        //                System.Threading.Thread.Sleep(1000);// ждем секунду чтобы гарантированно выполнилось преобразование
        //                tm = DateTime.Now;
        //            }

        //            // переименовать файл в правильное имя
        //            File.Move(string.Format(@"{0}\{1}", m_fullpath, release_name), string.Format(@"{0}\{1}", m_fullpath, makenewSberFileName() + makenewSberFileExt(i)));

        //            if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
        //                File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));

        //            if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)))
        //                File.Delete(string.Format(@"{0}\{1}", m_fullpath, release_name));

        //            // на Архив - можно забить уже
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
        //            MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
        //        }
        //    }
        //    catch (Win32Exception e)
        //    {
        //        if (e.NativeErrorCode == ERROR_FILE_NOT_FOUND)
        //        {
        //            MessageBox.Show("Ошибка приложения. Проверьте путь доступа к файлу: " + e.Message, "Внимание!", MessageBoxButtons.OK);
        //        }

        //        else if (e.NativeErrorCode == ERROR_ACCESS_DENIED)
        //        {
        //            MessageBox.Show("Ошибка приложения. Доступ к файлу запрещен: " + e.Message, "Внимание!", MessageBoxButtons.OK);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        if (DBFcon.State == System.Data.ConnectionState.Open)
        //        {
        //            DBFcon.Close();
        //            DBFcon.Dispose();
        //        }
        //        MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("По указанному пути в конфигурационном файле отсутствует директория для ведения архива. В эту папку файл сохранён не будет.", "Внимание", MessageBoxButtons.OK);
                    archive_folder_tofind = "";
                }

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                {
                    DialogResult rv = MessageBox.Show("По пути " + string.Format(@"{0}\{1}", m_fullpath, tofind_name) + ", указанном в конфигурационном файле, существует файл. Выгрузка прекращена, очистите папку вручную, если это необходимо.", "Внимание", MessageBoxButtons.OK);
                    return iCnt; // завершить программу выходом
                }
                
                CreateKtfomsToFind_DBF(bVFP, m_fullpath, tofind_name);

                // TODO: написать вставку строки в дбф для ктфомс
                DBFcon = new OleDbConnection();
                if (bVFP) DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                else DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                DBFcon.Open();

                //Decimal nOsp = GetOSP_Num();

                prbWritingDBF.Value = 0;

                int iDocCnt = 0; 
                if (DT_ktfoms_doc != null) iDocCnt = DT_ktfoms_doc.Rows.Count; ;

                prbWritingDBF.Maximum = 2*iDocCnt;// 2 раза чтобы update сделать потом еще столько же раз

                prbWritingDBF.Step = 1;

                // автоматом больше не пишем
                ////записываем в DBF возбужденные ИП
                //foreach (DataRow row in DT_ktfoms_reg.Rows)
                //{
                //    // TODO: даты нужны другие - для KTFOMS
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
                    // получить код соглашения - в дальнейшем он вобще будет параметром как pens_id будет pens_agr_code.
                    nAgreementID = Convert.ToInt32(DT_ktfoms_doc.Rows[0]["AGREEMENT_ID"]);
                    txtAgreementCode = GetAgreement_Code(nAgreementID);

                    //TODO: создать local_log LocalLogID
                    // 1 - cтатус Новый
                    // 1 - вид пакета Запрос
                    LLogID = CreateLLog(conGIBDD, 1, 1, txtAgreementCode, 0, "Пакет запросов.");

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
                            // отсчитать счетчиком
                            iCnt++;

                            // записать в лог
                            //WritePackLog(con, nPackID, "Обработан запрос # " + iCnt.ToString() + " ext_request_id = " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");
                            WriteLLog(conGIBDD,LLogID, "Обработан запрос # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");

                            
                        }
                        else
                        {
                            //WritePackLog(con, nPackID, "Ошибка! запрос # " + nID.ToString() + " обработать не удалось.\n");
                            WriteLLog(conGIBDD, LLogID, "Ошибка! запрос ext_request_id = " + nID.ToString() + " обработать не удалось.\n");
                            row["GOD"] = -1;
                        }
                        
                        //if (InsertRowToDBF(row, nOsp, 0, 1, DatZapr1_param, DatZapr2_param, tofind_name, true)) iCnt++;
                        prbWritingDBF.PerformStep();
                    }
                    
                    // записать количество в local_log
                    UpdateLLogCount(conGIBDD, LLogID, Convert.ToInt32(iCnt));
                }


                prbWritingDBF.PerformStep();
                DBFcon.Close();
                DBFcon.Dispose();

                //DataTable dt = GetDBFTable("SELECT NOMOSP, LITZDOLG, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FL_OKON, OSNOKON FROM TOFIND", "TOFIND1", string.Format(@"{0}\{1}", fullpath, "tofind.dbf"));
                //DBF.Save(dt, fullpath);


                if (!archive_folder_tofind.Equals(""))
                {
                    // функция сама с датами и копиями разберется
                    Copy(string.Format(@"{0}\{1}", m_fullpath, tofind_name), archive_folder_tofind);
                }
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }

            // если были запросы с ошибками
            if (nErrorPackID > 0)
            {
                // 2 - обработано
                UpdateLLogStatus(conGIBDD, nErrorPackID, 2);
                WriteLLog(conGIBDD, nErrorPackID, DateTime.Now + " Выгрузка пакета запросов завершена.\n");
            }

            // 2 - обработано
            UpdateLLogStatus(conGIBDD, LLogID, 2);
            WriteLLog(conGIBDD, LLogID, DateTime.Now + " пакет выгружен в файл: " + m_fullpath + "\\" + tofind_name + "\nВсего в файл выгружено запросов: " + iCnt.ToString() + "\n");

            // закомментировал потому что теперь пакеты внешние
            //// получить список пакетов
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
            //        WritePackLog(con, nRowPackID, DateTime.Now + " пакет выгружен в файл: " + m_fullpath + "\\" + tofind_name+ "\n");
            //        WritePackLog(con, nRowPackID, "Всего в файл выгружено запросов: " + iCnt.ToString() + "\n");
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
                    MessageBox.Show("По указанному пути в конфигурационном файле отсутствует директория для ведения архива. В эту папку файл сохранён не будет.", "Внимание", MessageBoxButtons.OK);
                    archive_folder_tofind = "";
                }
                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                {
                    
                    DialogResult rv = MessageBox.Show("По пути " + string.Format(@"{0}\{1}", m_fullpath, tofind_name) + ", указанном в конфигурационном файле, существует файл. Выгрузка прекращена, очистите папку вручную, если это необходимо.", "Внимание", MessageBoxButtons.OK);
                    return iCnt; // завершить программу выходом
                    
                }
                
                CreatePensToFind_DBF(bVFP, m_fullpath, tofind_name);
                

                // TODO: написать вставку строки в дбф для ктфомс
                DBFcon = new OleDbConnection();
                if (bVFP) DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
                else DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
                DBFcon.Open();

                //Decimal nOsp = GetOSP_Num();

                prbWritingDBF.Value = 0;
                if (DT_pens_doc != null)
                {
                    prbWritingDBF.Maximum = DT_pens_doc.Rows.Count * 2; //  в 2 раза больше чтобы потом еще update сделать
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
                    // получить код соглашения - в дальнейшем он вобще будет параметром как pens_id будет pens_agr_code.
                    nAgreementID = Convert.ToInt32(DT_pens_doc.Rows[0]["AGREEMENT_ID"]);
                    txtAgreementCode = GetAgreement_Code(nAgreementID);

                    //TODO: создать local_log LocalLogID
                    // 1 - cтатус Новый
                    // 1 - вид пакета Запрос
                    LLogID = CreateLLog(conGIBDD, 1, 1, txtAgreementCode, 0, "Пакет запросов.");
                    
                    foreach (DataRow row in DT_pens_doc.Rows)
                    {
                        // теперь не нужен этот параметр - лог пишм другой LLogID
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
                            // отсчитать счетчиком
                            iCnt++;

                            // TODO: записать в local_log
                            WriteLLog(conGIBDD, LLogID, "Обработан запрос # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");
                            
                            //// записать в лог
                            //WritePackLog(con, nPackID, "Обработан запрос # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");

                        }
                        else
                        {

                            // TODO: записать в local_log
                            WriteLLog(conGIBDD, LLogID, "Ошибка! запрос ext_request_id " + nID.ToString() + " обработать не удалось.\n");
                            //WritePackLog(con, nPackID, "Ошибка! запрос # " + nID.ToString() + " обработать не удалось.\n");
                            row["GOD"] = -1;
                        }

                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }

                    // записать количество в local_log
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
                //    // функция сама с датами и копиями разберется
                //    Copy(string.Format(@"{0}\{1}", m_fullpath, tofind_name), archive_folder_tofind);
                //}

                //release_name = endtofindname;
                // вот тут надо будет перекодировать!!!
                Process proc = new Process();
                proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");
                proc.StartInfo.Arguments = string.Format(@"{0}\{1}", m_fullpath, tofind_name) + " " + string.Format(@"{0}\{1}", m_fullpath, release_name);
                proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                proc.Start();


                System.Threading.Thread.Sleep(5000);// ждем 5 секунд чтобы гарантированно выполнилось преобразование.
                Int32 iCounter = 0;

                while (!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name))) 
                {
                    System.Threading.Thread.Sleep(1000);
                    iCounter++;
                    if (iCounter == 600)
                    {
                        // если прошло 10 минут
                        Exception ex = new Exception("Ошибка. Запросы в ОПФ не были сконвертированы в формат Fox 2.x и не переведены на статус ''Отправлено''.");
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if ((DBFcon != null) &&( DBFcon.State == System.Data.ConnectionState.Open))
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }

            // TODO: вместо списка пакетов будем использовать уже существующий LocalLogID

            // если были запросы с ошибками
            if (nErrorPackID > 0)
            {
                // 2 - обработано
                UpdateLLogStatus(conGIBDD, nErrorPackID, 2);
                WriteLLog(conGIBDD, nErrorPackID, DateTime.Now + " Выгрузка пакета запросов завершена.\n");
            }

            // 2 - обработано
            UpdateLLogStatus(conGIBDD, LLogID, 2);
            WriteLLog(conGIBDD, LLogID, DateTime.Now + " пакет выгружен в файл: " + m_fullpath + "\\" + tofind_name + "\nВсего в файл выгружено запросов: " + iCnt.ToString() + "\n");
            
            // закомментировал потому что теперь пакеты внешние

            //// получить список пакетов
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
            //        WritePackLog(con, nRowPackID, DateTime.Now + " пакет выгружен в файл: " + m_fullpath + "\\" + tofind_name + "\n");
            //        WritePackLog(con, nRowPackID, "Всего в файл выгружено запросов: " + iCnt.ToString() + "\n");
            //    }
            //}

            if (DT_pens_doc != null)
            {
                foreach (DataRow row in DT_pens_doc.Rows)// select только ради сортировки
                {
                    // обновить запрос и пакет
                    // UpdatePackRequest(row);

                    // сделать update строчек ext_request
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
                    MessageBox.Show("По указанному пути в конфигурационном файле отсутствует директория для ведения архива. В эту папку файл сохранён не будет.", "Внимание", MessageBoxButtons.OK);
                    archive_folder_tofind = "";
                }
                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                {
                    DialogResult rv = MessageBox.Show("По пути " + string.Format(@"{0}\{1}", m_fullpath, tofind_name) + ", указанном в конфигурационном файле, существует файл. Выгрузка прекращена, очистите папку вручную, если это необходимо.", "Внимание", MessageBoxButtons.OK);
                    return iCnt; // завершить программу выходом
                        
                }
                
                CreatePotdToFind_DBF(bVFP, m_fullpath, tofind_name);

                // TODO: написать вставку строки в дбф для ктфомс
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

                    // получить код соглашения - в дальнейшем он вобще будет параметром как pens_id будет pens_agr_code.
                    nAgreementID = Convert.ToInt32(DT_potd_doc.Rows[0]["AGREEMENT_ID"]);
                    txtAgreementCode = GetAgreement_Code(nAgreementID);

                    //TODO: создать local_log LocalLogID
                    // 1 - cтатус Новый
                    // 1 - вид пакета Запрос
                    LLogID = CreateLLog(conGIBDD, 1, 1, txtAgreementCode, 0, "Пакет запросов.");
                    

                    //записываем в DBF
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
                            // отсчитать счетчиком
                            iCnt++;

                            // записать в лог
                            //WritePackLog(con, nPackID, "Обработан запрос # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");
                            WriteLLog(conGIBDD, LLogID, "Обработан запрос # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");

                        }
                        else
                        {
                            //WritePackLog(con, nPackID, "Ошибка! запрос # " + nID.ToString() + " обработать не удалось.\n");
                            WriteLLog(conGIBDD, LLogID, "Ошибка! запрос ext_request_id " + nID.ToString() + " обработать не удалось.\n");
                            row["GOD"] = -1;
                        }

                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    // записать количество в local_log
                    UpdateLLogCount(conGIBDD, LLogID, Convert.ToInt32(iCnt));
                }


                prbWritingDBF.PerformStep();
                DBFcon.Close();
                DBFcon.Dispose();

                //DataTable dt = GetDBFTable("SELECT NOMOSP, LITZDOLG, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FL_OKON, OSNOKON FROM TOFIND", "TOFIND1", string.Format(@"{0}\{1}", fullpath, "tofind.dbf"));
                //DBF.Save(dt, fullpath);


                //if (!archive_folder_tofind.Equals(""))
                //{
                //    // функция сама с датами и копиями разберется
                //    Copy(string.Format(@"{0}\{1}", m_fullpath, tofind_name), archive_folder_tofind);
                //}

                // вот тут надо будет перекодировать!!!
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
                        // если прошло 10 минут
                        Exception ex = new Exception("Ошибка. Расширенные запросы в ОПФ не были сконвертированы в формат Fox 2.x и не переведены на статус ''Отправлено''.");
                        throw ex;
                    }   
                }
                
                //System.Threading.Thread.Sleep(5000);// ждем 5 секунд чтобы гарантированно выполнилось преобразование.

                //DateTime tm;
                //tm = DateTime.Now;
                //while (!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)) || (File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name)).AddMilliseconds(3000) > tm)) // пока не появился сконвертированный файл
                //{
                //    System.Threading.Thread.Sleep(1000);// ждем секунду чтобы гарантированно выполнилось преобразование.
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (DBFcon.State == System.Data.ConnectionState.Open)
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }

            // если были запросы с ошибками
            if (nErrorPackID > 0)
            {
                // 2 - обработано
                UpdateLLogStatus(conGIBDD, nErrorPackID, 2);
                WriteLLog(conGIBDD, nErrorPackID, DateTime.Now + " Выгрузка пакета запросов завершена.\n");
            }

            // 2 - обработано
            UpdateLLogStatus(conGIBDD, LLogID, 2);
            WriteLLog(conGIBDD, LLogID, DateTime.Now + " пакет выгружен в файл: " + m_fullpath + "\\" + tofind_name + "\nВсего в файл выгружено запросов: " + iCnt.ToString() + "\n");

            // закомментировал потому что теперь пакеты внешние


            //// получить список пакетов
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
            //        WritePackLog(con, nRowPackID, DateTime.Now + " пакет выгружен в файл: " + m_fullpath + "\\" + tofind_name + "\n");
            //        WritePackLog(con, nRowPackID, "Всего в файл выгружено запросов: " + iCnt.ToString() + "\n");
            //    }
            //}

            if (DT_potd_doc != null)
            {
                foreach (DataRow row in DT_potd_doc.Rows)// select только ради сортировки
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
                    MessageBox.Show("По указанному пути в конфигурационном файле отсутствует директория для ведения архива. В эту папку файл сохранён не будет.", "Внимание", MessageBoxButtons.OK);
                    archive_folder_tofind = "";
                }
                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                {
                    DialogResult rv = MessageBox.Show("По пути " + string.Format(@"{0}\{1}", m_fullpath, tofind_name) + ", указанном в конфигурационном файле, существует файл. Удалить его?", "Внимание", MessageBoxButtons.YesNo);
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
                //    // функция сама с датами и копиями разберется
                //    Copy(string.Format(@"{0}\{1}", m_fullpath, tofind_name), archive_folder_tofind);
                //}

                //release_name = endtofindname;
                // вот тут надо будет перекодировать!!!
                Process proc = new Process();
                proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");
                proc.StartInfo.Arguments = string.Format(@"{0}\{1}", m_fullpath, tofind_name) + " " + string.Format(@"{0}\{1}", m_fullpath, release_name);
                proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                proc.Start();


                System.Threading.Thread.Sleep(5000);// ждем 5 секунд чтобы гарантированно выполнилось преобразование.

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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if ((DBFcon != null) && (DBFcon.State == System.Data.ConnectionState.Open))
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return iCnt;
        }

        private Int64 FillDBF_Krc(bool bVFP, string m_fullpath, Int64 iCnt, string tablename)
        {
            // TODO: написать вставку строки в дбф для ктфомс
            DBFcon = new OleDbConnection();
            if (bVFP) DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + m_fullpath + ";Mode=ReadWrite;Collating Sequence=RUSSIAN");
            else DBFcon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", m_fullpath);
            DBFcon.Open();

            prbWritingDBF.Value = 0;
            prbWritingDBF.Maximum = DT_krc_reg.Rows.Count;
            prbWritingDBF.Step = 1;

            //записываем в DBF возбужденные ИП
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
                    MessageBox.Show("По указанному пути в конфигурационном файле отсутствует директория для ведения архива. В эту папку файл сохранён не будет.", "Внимание", MessageBoxButtons.OK);
                    archive_folder_tofind = "";
                }
                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                    File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name)); // удалить файл tofind - т.к. он все равно не итоговый

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name)))
                {
                    DialogResult rv = MessageBox.Show("По пути " + string.Format(@"{0}\{1}", m_fullpath, release_name) + ", указанном в конфигурационном файле, существует файл. Выгрузка прекращена, очистите папку вручную, если это необходимо.", "Внимание", MessageBoxButtons.OK);
                    return iCnt; // завершить программу выходом
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
                
                // Если вдруг почему-то 0, то 1 - МОСП
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
                    // получить код соглашения - в дальнейшем он вобще будет параметром как pens_id будет pens_agr_code.

                    // TODO: вот тут самое интересное - потому что логов должно быть много - по числу соглашений в списке!!!
                    nAgreementID = Convert.ToInt32(DT_doc_fiz.Rows[0]["AGREEMENT_ID"]);
                    txtAgreementCode = GetAgreement_Code(nAgreementID);

                    //TODO: создать local_log LocalLogID
                    // 1 - cтатус Новый
                    // 1 - вид пакета Запрос
                    LLogID = CreateLLog(conGIBDD, 1, 1, txtAgreementCode, 0, "Пакет запросов.");


                    foreach (DataRow row in DT_doc_fiz.Select("ZAPROS > 0", "FIOVK"))// select только ради сортировки
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
                            // отсчитать счетчиком
                            iCnt++;

                            // записать в лог - только для пакета с банка Возрождения
                            //WritePackLog(con, nPackID, "Обработан запрос # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");
                            WriteLLog(conGIBDD, LLogID, "Обработан запрос # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");

                        }
                        else
                        {
                            //WritePackLog(con, nPackID, 
                            WriteLLog(conGIBDD, LLogID, "Ошибка! запрос ext_request_id = " + nID.ToString() + " обработать не удалось.\n");
                            row["GOD"] = -1; // если вставка была неудачной (нет года рождения) - то не надо потом менять статус на отправлен
                            // а вот и надо менять - чтобы они не копились там пачками...
                        }

                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    // записать количество в local_log (пишем сразу итого - вдруг физиков вобще нет
                    UpdateLLogCount(conGIBDD, LLogID, Convert.ToInt32(iCnt)); // лог общий для всех и физ и юр.

                }



                if ((DT_doc_jur != null) && (DT_doc_jur.Rows.Count > 0))
                {
                    // получить код соглашения - в дальнейшем он вобще будет параметром как pens_id будет pens_agr_code.

                    // TODO: вот тут самое интересное - потому что логов должно быть много - по числу соглашений в списке!!!
                    // если выше был null в таблице юриков и не получен код соглашения
                    if (nAgreementID == 0)
                    {
                        nAgreementID = Convert.ToInt32(DT_doc_jur.Rows[0]["AGREEMENT_ID"]);
                        txtAgreementCode = GetAgreement_Code(nAgreementID);

                        //TODO: создать local_log LocalLogID
                        // 1 - cтатус Новый
                        // 1 - вид пакета Запрос
                        LLogID = CreateLLog(conGIBDD, 1, 1, txtAgreementCode, 0, "Пакет запросов.");
                    }

                    foreach (DataRow row in DT_doc_jur.Select("ZAPROS > 0", "FIOVK"))// select только ради сортировки
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
                            // отсчитать счетчиком
                            iCnt++;

                            // записать в лог - только для пакета с банка Возрождения
                            // WritePackLog(con, nPackID, "Обработан запрос # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");
                            WriteLLog(conGIBDD, LLogID, "Обработан запрос на юр. лицо # " + iCnt.ToString() + " ext_request_id = " + nID.ToString() + "\n");

                        }
                        else
                        {
                            // WritePackLog(con, nPackID,
                            WriteLLog(conGIBDD, LLogID, "Ошибка! запрос на юр. лицо ext_request_id = " + nID.ToString() + " обработать не удалось.\n");
                            row["GOD"] = -1; // если вставка была неудачной (нет года рождения) - то не надо потом менять статус на отправлен
                        }
                        prbWritingDBF.PerformStep();
                        prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    // записать количество в local_log
                    UpdateLLogCount(conGIBDD, LLogID, Convert.ToInt32(iCnt));
                }

                prbWritingDBF.PerformStep();
                DBFcon.Close();
                DBFcon.Dispose();

                // вот тут надо будет перекодировать!!!
                Process proc = new Process();
                proc.StartInfo.FileName = string.Format(@"{0}\{1}", System.Windows.Forms.Application.StartupPath, "fox622.exe ");
                proc.StartInfo.Arguments = string.Format(@"{0}\{1}", m_fullpath, tofind_name) + " " + string.Format(@"{0}\{1}", m_fullpath, release_name);
                proc.StartInfo.WorkingDirectory = "C:\\Program Files\\SSP\\InstallInfoChange";
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                proc.Start();

                System.Threading.Thread.Sleep(5000);// ждем 5 секунд чтобы гарантированно выполнилось преобразование.

                DateTime tm;
                tm = DateTime.Now;
                Int32 iCounter = 0;
                
                while ((!File.Exists(string.Format(@"{0}\{1}", m_fullpath, release_name))) || (File.GetLastWriteTime(string.Format(@"{0}\{1}", m_fullpath, release_name)).AddMilliseconds(100) > tm)) // пока не появился сконвертированный файл
                {
                    System.Threading.Thread.Sleep(1000);// ждем секунду чтобы гарантированно выполнилось преобразование.
                    tm = DateTime.Now;
                    iCounter++;
                    if (iCounter == 600)
                    {
                        // если прошло 10 минут
                        Exception ex = new Exception("Ошибка. Запросы не могут быть отправлены в Сбербанк. Слишком долго шла конвертация в формат Fox 2.x");
                        throw ex;
                    }
                }

                if (File.Exists(string.Format(@"{0}\{1}", m_fullpath, tofind_name)))
                    File.Delete(string.Format(@"{0}\{1}", m_fullpath, tofind_name));

                if (!archive_folder_tofind.Equals(""))
                {
                    Copy(string.Format(@"{0}\{1}", m_fullpath, release_name), archive_folder_tofind);
                }

                // если все хорошо - то надо сделать UPDATE ЗАПРОСА в бд установить СТАТУС - ОТПРАВЛЕН (10)
                // поскольку через 1 запрос - Банк возрождение делаем запросы во все банки через tofind.dbf
                // то делаем для всех UPDATE

                // по уму надо сделать обновление так:
                // отсортировать список DT_doc_fiz.Rows по параметру пакет
                // обновлять пакеты и запросы, вести флаг статуса пакета и после всех запросов обновлять пакет.

                // если были запросы с ошибками - установить статус обработано для их лога
                if (nErrorPackID > 0)
                {
                    // 2 - обработано
                    UpdateLLogStatus(conGIBDD, nErrorPackID, 2);
                    WriteLLog(conGIBDD, nErrorPackID, DateTime.Now + " Выгрузка пакета запросов завершена.\n");
                }

                // установить статсу обработано для лога запросов обычных
                // 2 - обработано
                UpdateLLogStatus(conGIBDD, LLogID, 2);
                WriteLLog(conGIBDD, LLogID, DateTime.Now + " пакет выгружен в файл: " + m_fullpath + "\\" + tofind_name + "\nВсего в файл выгружено запросов: " + iCnt.ToString() + "\n");

                // теперь пакет логов для банка возрождение сформирован, и его тупо надо скопировать для всех остальных КР. ОРГОВ

                decimal nOrg_id;
                decimal nAgr_id;

                // вот какая проблема - не нужно копировать 30 (банк возрождение). Он - вроде как первый.
                txtAgreementCode = "";

                // здесь делаем цикл про списку контрагентов-кред.оргов и копируем запись с логом

                foreach (string txtOrg_id in Legal_List)
                {
                    // по коду организации узнать номер согашения
                    nOrg_id = Convert.ToDecimal(txtOrg_id);
                    // получить Agreement_ID по номеру контрагента
                    nAgr_id = GetAgr_by_Org(nOrg_id);
                    txtAgreementCode = GetAgreement_Code(nAgr_id);

                    if(txtAgreementCode != "30") // сам 30 - Возрождение копировать не надо, т.к. он является эталонным.
                        CopyLLogParent(conGIBDD, LLogID, txtAgreementCode);
                }

                // закомментировал потому что теперь пакеты внешние
                //// получить список пакетов
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
                //        WritePackLog(con, nRowPackID, DateTime.Now + " пакет выгружен в файл: " + m_fullpath + "\\" + tofind_name + "\n");
                //        WritePackLog(con, nRowPackID, "Всего в файл выгружено запросов: " + iCnt.ToString() + "\n");
                //    }
                //}


                if (DT_doc_fiz != null)
                {
                    foreach (DataRow row in DT_doc_fiz.Rows)// select только ради сортировки
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

                // закомментировал - потому что теперь все логи во внешней таблице и не надо трогать пакеты ПК ОСП
                //// получить список пакетов
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
                //        WritePackLog(con, nRowPackID, DateTime.Now + " пакет выгружен в файл: " + m_fullpath + "\\" + tofind_name + "\n");
                //        WritePackLog(con, nRowPackID, "Всего в файл выгружено запросов: " + iCnt.ToString() + "\n");
                //    }
                //}

                if (DT_doc_jur != null)
                {
                    foreach (DataRow row in DT_doc_jur.Rows)// select только ради сортировки
                    {
                        //UpdatePackRequest(row);
                        //UpdateKredOrgRequest(row);
                        // сделать обновления запросов в кредитные организации
                        // суть проблемы в том - что мы сделали выборку только по одному из банков - а обновлять надо по всем.
                        // возникает вопрос - как получить список всех? нужен sql или какой-то алгоритм
                        // есть поле ext_request.req_id = DBF.zapros
                        // есть ext_request.agreement_code = 

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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Win32Exception e)
            {
                if (e.NativeErrorCode == ERROR_FILE_NOT_FOUND)
                {
                    MessageBox.Show("Ошибка приложения. Проверьте путь доступа к файлу: " + e.Message, "Внимание!", MessageBoxButtons.OK);
                }

                else if (e.NativeErrorCode == ERROR_ACCESS_DENIED)
                {
                    MessageBox.Show("Ошибка приложения. Доступ к файлу запрещен: " + e.Message, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if ((DBFcon != null) && (DBFcon.State == System.Data.ConnectionState.Open))
                {
                    DBFcon.Close();
                    DBFcon.Dispose();
                }
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            

            return iCnt;

        }


        private bool InsertKtfomsRowToDBF(DataRow row, DateTime dat1, DateTime dat2, bool bDoc, string tablename, ref decimal nErrorPackID)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);

            // TODO: неплохо бы передавать код соглашения в row
            int nAgreementID = Convert.ToInt32(row["AGREEMENT_ID"]);

            // все они пригодятся только если сообщение об ошибке писать нужно будет
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

                // получить дату рождения и год рождения
                if (!DateTime.TryParse(txtDD_R, out dtDate))
                {
                    dtDate = DateTime.MaxValue;
                    bNoYearBorn = true; // проблемы с датой рождения
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

                        //// чmобы влезало в 20 символов отрезаем начало '86/'
                        //txtNum_IP = cutEnd(Convert.ToString(row["ZAPROS"]).Trim().Substring(3), 20);
                        //m_cmd.CommandText += ", '" + txtNum_IP + "'";

                        // номип - число 5 знаков... =((
                        // придется анализировать строку и искать его

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

                        // надо вытащить из FIOVK отдельно Ф, И и О

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


                        // дату рождения мы еще раньше вытащили
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
                        txtResponse += "Запрос не был отправлен контрагентам, так как у должника - физ. лица не заполнена дата рождения.";
                    }

                    decimal nStatus = 15;
                    string txtZapros = Convert.ToString(row["ZAPROS"]).Trim();
                    Decimal nID = 0;
                    if (!Decimal.TryParse(txtZapros, out nID))
                    {
                        nID = -1;
                    }

                    txtResponse += " ZAPROS = " + txtZapros + "\n";
                    // TODO: получить номер запроса по nID

                    if (nID > 0)
                    {
                        
                        //InsertZaprosTo_PK_OSP(con, nID, txtResponse, DateTime.Now, nStatus, ktfoms_id, ref iRewriteState);
                        // TODO: тут надо написать новую функцию - через интерфейсные таблицы =)
                        // вытащить параметры - пакет, agent_agreement, agent_dept_code, agent_code, enity_name
                        txtAgreementCode = GetAgreement_Code(nAgreementID);
                        txtAgentCode = GetAgent_Code(nAgreementID);
                        txtAgentDeptCode = GetAgentDept_Code(nAgreementID);

                        string txtEntityName = GetLegal_Name(ktfoms_id);

                        // получить nAgent_id, nAgent_dept_id
                        decimal nAgent_id = GetAgent_ID(nAgreementID);
                        decimal nAgent_dept_id = GetAgentDept_ID(nAgreementID);

                        // если нового пакета еще не сделали - то делаем
                        if (nErrorPackID == 0)
                        {
                            // ставим статус обработан - 70
                            //nErrorPackID = ID_CreateDX_PACK_I(con, 70, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);
                            // TODO: вставить новый лог с ответами
                            // -1 Запрос не был отправлен, автоматически сформирован пакет ответов о неотправке запроса
                            nErrorPackID = CreateLLog(conGIBDD, 1, -1, txtAgreementCode, 0, "Пакет запросов, которые не были выгружены т.к. в них не заполнены обязательные поля.\n");
                            // WritePackLog(con, nErrorPackID, "Этот пакет создан для входящих ответов, которые автоматически созданы при выгрузке некорректных запросов (нет даты рождения или ИНН). Пакет пустой и служит индикатором того, что после обработки интерфейсных таблиц должен появиться еще такой пакет на статусе новый.");
                        }

                        InsertResponseIntTable(con, nID, txtResponse, DateTime.Now, nStatus, ktfoms_id, ref iRewriteState, nErrorPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName);

                        // пешем в лог что запрос не выгружен
                        WriteLLog(conGIBDD, nErrorPackID, txtResponse);
                        // сделать ++ для количества запросов в логе пакета
                        AppendLLogCount(conGIBDD, nErrorPackID, 1);
                    }
                    return false;
                }
        }
        catch (OleDbException ole_ex)
        {
            foreach (OleDbError err in ole_ex.Errors)
            {
                MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
            }
            return false;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            return false;
        }
        }

        private bool InsertPensRowToDBF(DataRow row, DateTime dat1, DateTime dat2, bool bDoc, string tablename, ref decimal nErrorPackID)
        {
            OleDbConnection conGIBDD;
            conGIBDD = new OleDbConnection(constrGIBDD);
            
            // TODO: неплохо бы передавать код соглашения в row
            int nAgreementID = Convert.ToInt32(row["AGREEMENT_ID"]);
            
            // все они пригодятся только если сообщение об ошибке писать нужно будет
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

            // получить дату рождения и год рождения
            if (!DateTime.TryParse(txtDateBornD, out dtDatrozhd))
            {
                bNoYearBorn = true; // проблемы с годом рождения
            }
            try
            {
                if (!bNoYearBorn)// если нет косяков по Году Рождения
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

                            // внимание!!! при обработке ответа из пенсионного надо брать не zapros а nomzap
                            m_cmd.CommandText += ", '" + cutEnd(Convert.ToString(row["NOMIP"]).Trim(), 40) + "'";

                            //// номип - число 5 знаков... =((
                            //// придется анализировать строку и искать его

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


                            // Здесь написано лучше чем ParceFIO

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
                        txtResponse += "Запрос не был отправлен контрагентам, так как у должника - физ. лица не заполнена дата рождения.";
                    }

                    decimal nStatus = 15;
                    string txtZapros = Convert.ToString(row["ZAPROS"]).Trim();
                    Decimal nID = 0;
                    if (!Decimal.TryParse(txtZapros, out nID))
                    {
                        nID = -1;
                    }

                    txtResponse += " ZAPROS = " + txtZapros + "\n";
                    // TODO: получить номер запроса по nID

                    if (nID > 0)
                    {
                        // TODO: тут надо написать новую функцию - через интерфейсные таблицы =)
                        // вытащить параметры - пакет, agent_agreement, agent_dept_code, agent_code, enity_name
                        txtAgreementCode = GetAgreement_Code(nAgreementID);
                        txtAgentCode = GetAgent_Code(nAgreementID);
                        txtAgentDeptCode = GetAgentDept_Code(nAgreementID);

                        string txtEntityName = GetLegal_Name(pens_id);

                        // получить nAgent_id, nAgent_dept_id
                        decimal nAgent_id = GetAgent_ID(nAgreementID);
                        decimal nAgent_dept_id = GetAgentDept_ID(nAgreementID);

                        // если нового пакета еще не сделали - то делаем
                        if (nErrorPackID == 0)
                        {
                            // TODO: вставить новый лог с ответами
                            nErrorPackID = CreateLLog(conGIBDD, 1, -1, txtAgreementCode, 0, "Пакет запросов, которые не были выгружены т.к. в них не заполнены обязательные поля.\n");
                            //nErrorPackID = ID_CreateDX_PACK_I(con, 70, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);
                            //WritePackLog(con, nErrorPackID, "Этот пакет создан для входящих ответов, которые автоматически созданы при выгрузке некорректных запросов (нет даты рождения или ИНН). Пакет пустой и служит индикатором того, что после обработки интерфейсных таблиц должен появиться еще такой пакет на статусе новый.");
                        }

                        // пешем в лог что запрос не выгружен
                        WriteLLog(conGIBDD, nErrorPackID, txtResponse);
                        // сделать ++ для количества запросов в логе пакета
                        AppendLLogCount(conGIBDD, nErrorPackID, 1);

                        // вставить в интерфейсную таблицу ответ, пакет указать nErrorPackID
                        InsertResponseIntTable(con, nID, txtResponse, DateTime.Now, nStatus, pens_id, ref iRewriteState, nErrorPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName);
                        
                        //InsertZaprosTo_PK_OSP(con, nID, txtResponse, DateTime.Now, nStatus, pens_id, ref iRewriteState);
                    }
                    return false;
                }
        }
        catch (Exception ex)
        {
            //if (DBFcon != null) DBFcon.Close();
            MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

            // TODO: неплохо бы передавать код соглашения в row
            int nAgreementID = Convert.ToInt32(row["AGREEMENT_ID"]);
            // все они пригодятся только если сообщение об ошибке писать нужно будет
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

            // получить дату рождения и год рождения
            if (!DateTime.TryParse(txtDateBornD, out dtDatrozhd))
            {
                bNoYearBorn = true; // проблемы с годом рождения
            }
            try
            {
                if (!bNoYearBorn)// если нет косяков по дате рождения
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

                                // номип - число 5 знаков... =((
                                // придется анализировать строку и искать его

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


                                // надо вытащить из FIOVK отдельно Ф, И и О

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
                                    //m_cmd.CommandText = "UPDATE ZAPROS SET STATUS = 'Ошибка в запросе', TEXT_ERROR = 'В карточке ИП не указана корректная дата рождения должника в формате ДД.ММ.ГГГГ', TEXT = 'В карточке ИП не указана корректная дата рождения должника ДД.ММ.ГГГГ'";
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
                        txtResponse += "Запрос не был отправлен контрагентам, так как у должника - физ. лица не заполнена дата рождения.";
                    }

                    decimal nStatus = 15;
                    string txtZapros = Convert.ToString(row["ZAPROS"]).Trim();
                    Decimal nID = 0;
                    if (!Decimal.TryParse(txtZapros, out nID))
                    {
                        nID = -1;
                    }
                    
                    txtResponse += " ZAPROS = " + txtZapros + "\n";
                        // TODO: получить номер запроса по nID

                    if (nID > 0)
                    {
                        // TODO: тут надо написать новую функцию - через интерфейсные таблицы =)
                        // вытащить параметры - пакет, agent_agreement, agent_dept_code, agent_code, enity_name
                        txtAgreementCode = GetAgreement_Code(nAgreementID);
                        txtAgentCode = GetAgent_Code(nAgreementID);
                        txtAgentDeptCode = GetAgentDept_Code(nAgreementID);

                        string txtEntityName = GetLegal_Name(potd_id);

                        // получить nAgent_id, nAgent_dept_id
                        decimal nAgent_id = GetAgent_ID(nAgreementID);
                        decimal nAgent_dept_id = GetAgentDept_ID(nAgreementID);

                        // если нового пакета еще не сделали - то делаем
                        if (nErrorPackID == 0)
                        {
                            //nErrorPackID = ID_CreateDX_PACK_I(con, 70, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);
                            //WritePackLog(con, nErrorPackID, "Этот пакет создан для входящих ответов, которые автоматически созданы при выгрузке некорректных запросов (нет даты рождения или ИНН). Пакет пустой и служит индикатором того, что после обработки интерфейсных таблиц должен появиться еще такой пакет на статусе новый.");
                            
                            // вставить новый лог с ответами
                            nErrorPackID = CreateLLog(conGIBDD, 1, -1, txtAgreementCode, 0, "Пакет запросов, которые не были выгружены т.к. в них не заполнены обязательные поля.\n");

                        }

                        InsertResponseIntTable(con, nID, txtResponse, DateTime.Now, nStatus, potd_id, ref iRewriteState, nErrorPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName);

                        // пишем в лог что запрос не выгружен
                        WriteLLog(conGIBDD, nErrorPackID, txtResponse);
                        // сделать ++ для количества запросов в логе пакета
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
                return false;
            }
            catch (Exception ex)
            {
                //if (DBFcon != null) DBFcon.Close();
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

                    // TODO: 40 - это очень мало!!! В ПКСП в IP 500 в ZAPROS 100
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
                    MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    // когда будет статус - ошибка - тут надо будет ставить запрос на статус ОШИБКА

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    // когда будет статус - ошибка - тут надо будет ставить запрос на статус ОШИБКА

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    // когда будет статус - ошибка - тут надо будет ставить запрос на статус ОШИБКА

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            
            return nPackID;

        }

        private bool SetSendlistDocumentStatus(Decimal nO_ID, Decimal nPackID, int nStatus)
        {
            // обновить на статус тот самый DOCUMENT, который является списком рассылки запроса O_ID и в пакете PACK_ID
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
                    // обновить на статус тот самый DOCUMENT, который является списком рассылки запроса O_ID и в пакете PACK_ID
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
                    // когда будет статус - ошибка - тут надо будет ставить запрос на статус ОШИБКА

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                return false;
            }
            return true;

        }

        private bool SetDocumentStatus(Decimal nID, int nStatus, int nSecondStatus, bool bOpposit)
        {
            // если нужно выражение != SECONDSTATUS - использовать уже готовую функцию.
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
                        // когда будет статус - ошибка - тут надо будет ставить запрос на статус ОШИБКА

                    }

                }
                catch (OleDbException ole_ex)
                {
                    foreach (OleDbError err in ole_ex.Errors)
                    {
                        MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    // когда будет статус - ошибка - тут надо будет ставить запрос на статус ОШИБКА

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    // когда будет статус - ошибка - тут надо будет ставить запрос на статус ОШИБКА

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    // когда будет статус - ошибка - тут надо будет ставить запрос на статус ОШИБКА

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    // когда будет статус - ошибка - тут надо будет ставить запрос на статус ОШИБКА

                }

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                return false;
            }
            return true;

        }


        private bool UpdateBankFizRequest(DataRow row)
        {
            Decimal nID = 0;
            string txtValue = "InfoChangeCredOrg"; // статус ОТПРАВЛЕН
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
                    // когда будет статус - ошибка - тут надо будет ставить запрос на статус ОШИБКА

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

            // надо будет не только этот пакет, row["pack_id"], но и аналогичные...
            // логика такова, что мне нужен не только этот sendlist и это этот пакет,
            // но и все остальные sendlist по этому запросу ID, которые отправлены контрагентам из LegalList
            // и все пакеты, в которые включены эти самые sendlist
            // то есть входными данными является req.id и Legal_list
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
            // обновить запрос из пакета
            decimal nID = 0;
            int iGod = 0;
            string txtGod, txtID, txtPackID, txtReq_id, txtAgreementCode;
            decimal nReq_id = 0;
            decimal nAgreement_code = 0;
            try{

                txtGod = Convert.ToString(row["GOD"]).Trim();

                // получить req_id
                txtReq_id = Convert.ToString(row["zapros"]).Trim();
                if (!Decimal.TryParse(txtReq_id, out nReq_id))
                {
                    nReq_id = 0;
                }

                if (!Int32.TryParse(txtGod, out iGod))
                {
                    iGod = 0;
                }

                // прогнать по списку все LegalList - получиьт mvv_agreement_code
                foreach (string txtOrg_id in Legal_List)
                {
                    decimal nOrg_id = Convert.ToDecimal(txtOrg_id);
                    nAgreement_code =  GetAgr_by_Org(nOrg_id);
                    string txtAgrCode = GetAgreement_Code(nAgreement_code);
                    
                    // обновить запись в ext_request по 2-м параметрам
                    
                    // теперь все обновляем - чтобы они не накапливались как невыгруженные, ведь мы вставили на них ответ
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                return false;
            }
            return true;
                        
        }


        private bool UpdateExtRequestRow(DataRow row)
        {
            // обновить запрос из пакета
            decimal nID = 0;
            decimal nReqID = 0;
            int iGod = 0;
            string txtGod, txtID, txtPackID, txtReqID;

            txtGod = Convert.ToString(row["GOD"]).Trim();
            txtID = Convert.ToString(row["ext_request_id"]).Trim();
            txtReqID = Convert.ToString(row["zapros"]).Trim();

            // надо будет не только этот пакет, row["pack_id"], но и аналогичные...
            // логика такова, что мне нужен не только этот sendlist и это этот пакет,
            // но и все остальные sendlist по этому запросу ID, которые отправлены контрагентам из LegalList
            // и все пакеты, в которые включены эти самые sendlist
            // то есть входными данными является req.id и Legal_list
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
                // 15 - выгружено с ошибкой
                SetDocumentStatus(nReqID, 15);
            }
            //else
            //{
            //    return SetExtReqProcessed(nID, 1);
            //}

            
            // теперь все обновляем как выгруженные, ведь мы вставляем ответ и если оставить его - то он будет копиться
            return SetExtReqProcessed(nID, 1);
        }

        private bool UpdatePackRequest(DataRow row)
        {
            // обновить запрос из пакета
            decimal nPackID = 0;
            decimal nID = 0;
            int iGod = 0;
            string txtGod, txtID, txtPackID;

            txtGod = Convert.ToString(row["GOD"]).Trim();
            txtID = Convert.ToString(row["ZAPROS"]).Trim();
            txtPackID = Convert.ToString(row["pack_id"]).Trim();

            // надо будет не только этот пакет, row["pack_id"], но и аналогичные...
            // логика такова, что мне нужен не только этот sendlist и это этот пакет,
            // но и все остальные sendlist по этому запросу ID, которые отправлены контрагентам из LegalList
            // и все пакеты, в которые включены эти самые sendlist
            // то есть входными данными является req.id и Legal_list
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
                    //  если не был вставлен запрос, то пакет и список рассылки перевести на статус 71 - ошибка электронной отправки
                    // если в пакете не все было с ошибкой - то нельзя ставить 71
                    // есть еще статус 70 - отправлен частично.
                    // как узнать что надо ставить 70 - а не 71 - если есть хоть 1 правильный
                    // то есть изначально 71  а если вдруг было 71 и вдруг пришел хороший запрос в пакете - то 70
                    nStatus = 71;
                }

                    foreach (string txtOrg_id in Legal_List)
                    {
                        bUpdatedPack = true;
                        bUpdated = true;
                        decimal nOrg_id = Convert.ToDecimal(txtOrg_id);
                        
                        // узнать что это sendlist, где nID с контрагентом txtOrg_id
                        decimal nPackID = GetPackIdFromSendlistByOrgId(nID, nOrg_id);
                        
                        // если запрос в пакете
                        if (nPackID > 0)
                            {
                                // ставим пакету статус - отправлен, причем если надо чтобы в случае хоть 1 пакета с ошибкой - была ошибка отправки пакета - то с OppositStatus
                                if (!SetDocumentStatus(nPackID, Convert.ToInt32(nStatus), 23, false))
                                {
                                    // если обновить не удалось - то наверное уже другой статус стоял и делаем так:
                                    // если стоял такой же статус, то частично, а иначе - оставить как был
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                return false;
            }
            return true;
        }


        private bool UpdatePackRequest(int nGod, decimal nID, decimal nPackID, decimal nStatus)
        {
            // обновить запрос из пакета, используя передачу параметров через переменные (а не DATAROW)
            

            bool bUpdated = true;
            bool bUpdatedPack = true;
            try
            {
                if (nGod == -1)
                {
                    //  если не был вставлен запрос, то пакет и список рассылки перевести на статус 71 - ошибка электронной отправки
                    // или 70 - частично отправлен
                    nStatus = 71;
                }

                    // установить статус пакета
                    if(nPackID != 0){
                            // если статуса вобще не было - надо бы его выставить все-же
                            // узнать на каком статусе пакет будет и написать функцию для установки если статус какой нужно, а не неравный ему
                            // 23 - отправка сторонней программой
                            if (!SetDocumentStatus(nPackID, Convert.ToInt32(nStatus), 23, false))
                            {
                                // если обновить не удалось - то наверное уже другой статус стоял и делаем так:
                                // если стоял такой же статус, то частично, а иначе - оставить как был
                                if (!SetDocumentStatus(nPackID, 70, Convert.ToInt32(nStatus)))
                                {
                                    bUpdatedPack = false;
                                }
                            }

                            
                        // установить статус запроса-пункта списка рассылки в пакете
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                return false;
            }
            return true;
        }

        
        private bool UpdateRequest(DataRow row)
        {
            
            Decimal nID = 0;
            int nStatus = 10; // статус ОТПРАВЛЕН
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
                        // когда будет статус - ошибка - тут надо будет ставить запрос на статус ОШИБКА

                    }
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                    return false;
                }
                return true;
        }

        // вставка строчки запроса в Крединую Организацию в DBF файл
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
            
            int nAgreementID = Convert.ToInt32(row["AGREEMENT_ID"]); // тут будет только 30 (Возрождение) - надо потом из списка вытащить
            // все они пригодятся только если сообщение об ошибке писать нужно будет
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
                    nLitzDolg = 1;// юр. лицо
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
                            txtInnOrg = "00" + txtInnOrg; // у юр. лиц слева добавляем 00, до 12 символов
                        }
                    }
                    txtInnOrg = cutEnd(txtInnOrg.Trim(), 12);

                    // вычислить vid_d и inn_d, если все плохо то не вставлять
                    if (bRegTable && (nLitzDolg == 1))
                    {
                        
                        if (txtInnOrg.Length < 10)
                        {
                            bNoInnJur = true;
                        }
                        // TODO: проверить по маске что тут 10, 11, 12 цифр
                    }

                    if (nLitzDolg == 2)// если физ. лицо
                    {
                        // TODO: решить вопрос с годом
                        if (!Int32.TryParse(Convert.ToString(row["GOD"]), out nBirthYear))
                        {
                            nBirthYear = 0;
                        }

                        txtDateBornD = Convert.ToString(row["DATROZHD"]).Trim();
                        dtDatrozhd = DateTime.MaxValue;

                        // получить дату рождения и год рождения
                        if (DateTime.TryParse(txtDateBornD, out dtDate))
                        {
                            if (nBirthYear == 0) nBirthYear = dtDate.Year; // если не был установлен г.р., то берем из даты
                            dtDatrozhd = dtDate;
                        }
                        else
                        {
                            dtDatrozhd = DateTime.MaxValue;
                            dtDate = DateTime.MaxValue;
                            if (nBirthYear == 0)
                            {
                                bNoYearBorn = true; // проблемы с годом рождения
                            }
                        }

                        
                    }

                    if (!bNoInnJur && !bNoYearBorn)// если нет косяков ни по ИНН ни по Году Рождения
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

                        // TODO: тут надо брать из GOD
                        m_cmd.CommandText += ", " + Convert.ToString(nBirthYear);

                        txtNomspi = Convert.ToString(row["NOMSPI"]).Trim();
                        if (!Int32.TryParse(txtNomspi, out nNomspi))
                        {
                            nNomspi = 0;
                            txtNomspi = "0";
                        }

                        m_cmd.CommandText += ", " + Convert.ToString(nNomspi);

                        txtNOMIP = Convert.ToString(row["IPNO_NUM"]).Trim(); // это номер ИП - десятичный (можно считать его порядковым)
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

                        //txtOsnokon = cutEnd(Convert.ToString(row["OSNOKON"]).Trim(), 250); - с конца 2010 года больше никакого основания окончания!!!
                        txtOsnokon = "";
                        m_cmd.CommandText += ", '" + txtOsnokon + "'";

                        m_cmd.Parameters.Add(new OleDbParameter("DATROZHD", OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "DATROZHD", System.Data.DataRowVersion.Original, dtDatrozhd));
                        m_cmd.CommandText += ", ?";

                        m_cmd.CommandText += ')';
                        m_cmd.ExecuteNonQuery();
                        m_cmd.Dispose();

                        

                        //if (Convert.ToInt32(row["ID_DBTRCLS"]) == 95)
                        // если это ИП (ID_DBTRCLS = 95) - то вставить в DBF еще и запрос по Юр. лицу
                            if ((Convert.ToInt32(row["ID_DBTRCLS"]) == 95) && (txtInnOrg.Length >= 10) && (bOKON == 0))
                            {
                                m_cmd = new OleDbCommand();
                                m_cmd.Connection = DBFcon;

                                m_cmd.CommandText = "INSERT INTO " + tofind_name + " (NOMOSP, LITZDOLG, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FL_OKON, OSNOKON, DATROZHD) VALUES (";

                                m_cmd.CommandText += Convert.ToString(nOsp);

                                nLitzDolg = 1;// юр. лицо
                                m_cmd.CommandText += ", " + Convert.ToString(nLitzDolg);

                                m_cmd.CommandText += ", '" + txtNameDolg + "'";

                                m_cmd.CommandText += ", '" + txtZapros + "'";

                                m_cmd.CommandText += ", " + Convert.ToString(nYear);

                                m_cmd.CommandText += ", " + nNomspi.ToString();

                                // номип - число 5 знаков... =((
                                // придется анализировать строку и искать его

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
                        // else - // тут надо будет менять статус запроса на Ошибка. Запрос не был отправлен (нет года рождения).
                        return true;

                    }
                    else // если есть косяки по году рождения или ИНН
                    {
                        string txtResponse = "";
                        if (bNoInnJur)
                        {
                            txtResponse += "Запрос не был отправлен контрагентам, так как некорректно заполнено поле ИНН для должника - юр. лица.";
                        }

                        if (bNoYearBorn)
                        {
                            txtResponse += "Запрос не был отправлен контрагентам, так как у должника - физ. лица не заполнены ни год ни дата рождения.";
                        }

                        decimal nStatus = 15;
                        string txtID = Convert.ToString(row["ZAPROS"]).Trim();
                        Decimal nID = 0;
                        if(!Decimal.TryParse(txtID, out nID)){
                            nID = -1;
                        }

                        txtResponse += " ZAPROS = " + txtID + "\n";
                        // TODO: получить номер запроса по nID
                        

                        if(nID > 0){
                            // вставляем ответ для каждого контрагента - то есть для всех банков
                            foreach (string txtOrg_id in Legal_List)
                            {
                                // по коду организации узнать номер согашения
                                decimal nOrg_id = Convert.ToDecimal(txtOrg_id);
                                // получить Agreement_ID по номеру контрагента
                                decimal nAgr_id = GetAgr_by_Org(nOrg_id);
                                
                                // получить nAgent_id, nAgent_dept_id
                                decimal nAgent_id = GetAgent_ID(nAgr_id);
                                decimal nAgent_dept_id = GetAgentDept_ID(nAgr_id);
                                
                                // Вставка ответа для кред. организации с сообщением об ошибке в интерфейсную таблицу

                                // вытащить параметры - пакет, agent_agreement, agent_dept_code, agent_code, enity_name
                                txtAgreementCode = GetAgreement_Code(nAgr_id);
                                txtAgentCode = GetAgent_Code(nAgr_id);
                                txtAgentDeptCode = GetAgentDept_Code(nAgr_id);

                                
                                // если нового пакета еще не сделали - то делаем
                                if (nErrorPackID == 0)
                                {
                                    //nErrorPackID = ID_CreateDX_PACK_I(con, 70, nAgent_id, nAgent_dept_id, nAgr_id, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);
                                    //WritePackLog(con, nErrorPackID, "Этот пакет создан для входящих ответов, которые автоматически созданы при выгрузке некорректных запросов (нет даты рождения или ИНН).Пакет пустой и служит индикатором того, что после обработки интерфейсных таблиц должен появиться еще такой пакет на статусе новый.");
                                    // TODO: вставить новый лог с ответами
                                    nErrorPackID = CreateLLog(conGIBDD, 1, -1, txtAgreementCode, 0, "Этот пакет создан для входящих ответов, которые автоматически созданы при выгрузке некорректных запросов (нет даты рождения или ИНН).\n");
                                }
                                
                                string txtEntityName = GetLegal_Name(nOrg_id);

                                // тут вставляется ОТВЕТ - причем отрицательный

                                InsertResponseIntTable(con, nID, txtResponse, DateTime.Now, nStatus, nOrg_id, ref iRewriteState, nErrorPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName);

                                // пешем в лог что запрос не выгружен
                                WriteLLog(conGIBDD, nErrorPackID, txtResponse);
                                // сделать ++ для количества запросов в логе пакета
                                AppendLLogCount(conGIBDD, nErrorPackID, 1);

                                // теперь нужно не забыть поправить запрос, что он уже выгружен...
                                // но по факту обновление ведь идет потом, так что там нужно придумать особый порядок

                                // тут вставляется ОТВЕТ - причем отрицательный
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
                        MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    //if (DBFcon != null) DBFcon.Close();
                    MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

                    // номип - число 5 знаков... =((
                    // придется анализировать строку и искать его

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
                        MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                    }
                    if (m_cmd != null) m_cmd.Dispose();
                    return false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

            // перекинуть все пакеты на подготовить к отправке
            // для сторонней отправки
            string txtUpdateSql = "";
            decimal nOrg_id = 0;



            // !закомментировал все UPDATE - т.к. это все очень долго и пусть это планировщик делает
            
            //// обработать Сбербанк
            //nOrg_id = 86200999999005;
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 23 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 3 and req_type.sndl_contr_id = " + nOrg_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// для автоматической отправки
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 11 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 2 and req_type.sndl_contr_id = " + nOrg_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// Обработать все что в Списке
            //foreach (string txtOrg_id in Legal_List)
            //{
            //    if (Decimal.TryParse(txtOrg_id, out nOrg_id))
            //    {
            //        txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 23 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 3 and req_type.sndl_contr_id = " + txtOrg_id + ")";
            //        UpdateSqlExecute(con, txtUpdateSql);

            //        // для автоматической отправки
            //        txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 11 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 2 and req_type.sndl_contr_id = " + txtOrg_id + ")";
            //        UpdateSqlExecute(con, txtUpdateSql);
            //    }
            //}

            //DT_doc = GetDataTableFromFB("SELECT DISTINCT UPPER(a.NAME_D) as FIOVK, a.NUM_IP as ZAPROS, a.VIDD_KEY as LITZDOLG, a.DATE_BORN_D as GOD, c.PRIMARY_SITE as NOMSPI, a.NUM_IP as NOMIP, a.SUM_ as SUMMA, a.WHY as VIDVZISK, a.INND as INNORG, b.DATE_DOC as DATZAPR, a.ADR_D as ADDR,a.TEXT_PP as OSNOKON, a.PK as FK_IP, a.PK_ID as FK_ID, a.uscode, a.FIO_SPI FROM IP a left join s_users c on (a.uscode=c.uscode) LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.DATE_IP_OUT is null and b.KOD = 1011 and (b.DATE_DOC >= '" + DatZapr1.ToShortDateString() + "' AND b.DATE_DOC <= '" + DatZapr2.ToShortDateString() + "') and a.ssd is null and a.ssv is null", "TOFIND");

            // т.к. запрос в банки делаем для всех, то убрал фильтр по типу должника
            // DT_doc_jur = GetDataTableFromFB("select 1 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 31 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 1 or (ip.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT  where ncc_parent_id = 1)))", "TOFIND");
            // DT_doc_jur = GetDataTableFromFB("select 1 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 31 and zapr_d.docstatusid = 2 and ip_d.docstatusid = 9 and (ip.ID_DBTRCLS = 1 or (ip.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 1)))", "TOFIND");
            //DT_doc_jur = GetDataTableFromFB("select pack.id as pack_id, 1 LITZDOLG, d_req.id ZAPROS, req.IPNO_NUM, req.div, req.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, d_req.doc_date DATZAPR,  req.ID_DBTR_ADR ADDR, req.ID_DBTR_BORN DATROZHD, req.ID_DBTRCLS,  req.DBTR_BORN_YEAR GOD, req.ID_DEBTSUM SUMMA, req.ID_DBTR_INN INNORG,  d_req.doc_number, req.ID_DEBTCLS_NAME VIDVZISK from dx_pack_o packo left join dx_pack pack on pack.id = packo.id join sendlist sl on pack.id = sl.dx_pack_id join o_ip req on sl.sendlist_o_id = req.id join document d_req on req.id = d_req.id join document ip_d on d_req.parent_id = ip_d.id join document dpack on pack.id = dpack.id join SPI on req.IP_EXEC_PRIST = spi.SUSER_ID where dpack.docstatusid = 23  and pack.agreement_id = 30 and packo.has_been_sent is null and d_req.docstatusid != 19 and d_req.docstatusid != 15 and (req.ID_DBTRCLS = 1 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 1)))", "TOFIND");
            DT_doc_jur = GetDataTableFromFB("select 30 agreement_id, ext_request_id,  pack_id,  1 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID where mvv_agreement_code = 30 and processed = 0 and (req.ID_DBTRCLS = 1 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 1)))", "TOFIND");

            // насколько корректно брать запросы только по одному из соглашений? -
            // в существующей архитектуре - это абсолютно корректно, т.к. у нас только 1 файл с запросами
            // НО! этот узкий момент неообходимо описать в документации - что все региональные кредитные выгружаются на основе agreement_id = 30 (БАНК ВОЗРОЖДЕНИЕ)

            // DT_doc_fiz = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 31 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_NSI_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            //DT_doc_fiz = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 31 and zapr_d.docstatusid = 2 and ip_d.docstatusid = 9 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            //DT_doc_fiz = GetDataTableFromFB("select pack.id as pack_id, 2 LITZDOLG, d_req.id ZAPROS, req.IPNO_NUM, req.div, req.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, d_req.doc_date DATZAPR,  req.ID_DBTR_ADR ADDR, req.ID_DBTR_BORN DATROZHD, req.ID_DBTRCLS,  req.DBTR_BORN_YEAR GOD, req.ID_DEBTSUM SUMMA, req.ID_DBTR_INN INNORG,  d_req.doc_number, req.ID_DEBTCLS_NAME VIDVZISK from dx_pack_o packo left join dx_pack pack on pack.id = packo.id join sendlist sl on pack.id = sl.dx_pack_id join o_ip req on sl.sendlist_o_id = req.id join document d_req on req.id = d_req.id join document ip_d on d_req.parent_id = ip_d.id join document dpack on pack.id = dpack.id join SPI on req.IP_EXEC_PRIST = spi.SUSER_ID where dpack.docstatusid = 23  and pack.agreement_id = 30 and packo.has_been_sent is null  and d_req.docstatusid != 19 and d_req.docstatusid != 15  and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            DT_doc_fiz = GetDataTableFromFB("select 30 agreement_id, ext_request_id,  pack_id,  2 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID where mvv_agreement_code = 30 and processed = 0 and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            
            
            // 30 - только Возрождение

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

            // тут типа true - bVFP
            // нужно как-то понять как из reg отфильтровать тех кого не надо уже писать в DBF
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
            
            // и еще - установить progress-bar в ноль и посчитать длину
            prbWritingDBF.Value = 0;
            prbWritingDBF.Step = 1;

            if (DT_doc != null)
            {
                prbWritingDBF.Maximum = DT_doc.Rows.Count;
                // последний параметр - надо ли фильтровать по ИНН
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
            // Получить названия организаций-адресатов
            Legal_Name_List = (String[])Legal_List.Clone();
            Legal_Сonv_List = (String[])Legal_List.Clone();
            //int code = 0;
            Decimal code = 0;
            //for (int i = 1; i < Legal_Name_List.Length; i++)
            for (int i = 0; i < Legal_Name_List.Length; i++)
            {
                //code = Convert.ToInt32((Legal_List[i]).Trim());
                code = Convert.ToDecimal((Legal_List[i]).Trim());
                Legal_Name_List[i] = GetLegal_Name(code);
                Legal_Сonv_List[i] = GetLegal_Conv(code);
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
                    case 1026: infc[t] = 'А'; break;
                    case 1027: infc[t] = 'Б'; break;
                    case 8218: infc[t] = 'В'; break;
                    case 1107: infc[t] = 'Г'; break;
                    case 8222: infc[t] = 'Д'; break;
                    case 8230: infc[t] = 'Е'; break;
                    case 1088: infc[t] = 'Ё'; break;
                    case 8224: infc[t] = 'Ж'; break;
                    case 8225: infc[t] = 'З'; break;
                    case 8364: infc[t] = 'И'; break;
                    case 8240: infc[t] = 'Й'; break;
                    case 1033: infc[t] = 'К'; break;
                    case 8249: infc[t] = 'Л'; break;
                    case 1034: infc[t] = 'М'; break;
                    case 1036: infc[t] = 'Н'; break;
                    case 1035: infc[t] = 'О'; break;
                    case 1039: infc[t] = 'П'; break;
                    case 1106: infc[t] = 'Р'; break;
                    case 8216: infc[t] = 'С'; break;
                    case 8217: infc[t] = 'Т'; break;
                    case 8220: infc[t] = 'У'; break;
                    case 8221: infc[t] = 'Ф'; break;
                    case 8226: infc[t] = 'Х'; break;
                    case 8211: infc[t] = 'Ц'; break;
                    case 8212: infc[t] = 'Ч'; break;
                    case 152: infc[t] = 'Ш'; break;
                    case 8482: infc[t] = 'Щ'; break;
                    case 1113: infc[t] = 'Ъ'; break;
                    case 8250: infc[t] = 'Ы'; break;
                    case 1114: infc[t] = 'Ь'; break;
                    case 1116: infc[t] = 'Э'; break;
                    case 1115: infc[t] = 'Ю'; break;
                    case 1119: infc[t] = 'Я'; break;

                    case 160: infc[t] = 'а'; break;
                    case 1038: infc[t] = 'б'; break;
                    case 1118: infc[t] = 'в'; break;
                    case 1032: infc[t] = 'г'; break;
                    case 164: infc[t] = 'д'; break;
                    case 1168: infc[t] = 'е'; break;
                    case 1089: infc[t] = 'ё'; break;
                    case 166: infc[t] = 'ж'; break;
                    case 167: infc[t] = 'з'; break;
                    case 1025: infc[t] = 'и'; break;
                    case 169: infc[t] = 'й'; break;
                    case 1028: infc[t] = 'к'; break;
                    case 171: infc[t] = 'л'; break;
                    case 172: infc[t] = 'м'; break;
                    case 173: infc[t] = 'н'; break;
                    case 174: infc[t] = 'о'; break;
                    case 1031: infc[t] = 'п'; break;
                    case 1072: infc[t] = 'р'; break;
                    case 1073: infc[t] = 'с'; break;
                    case 1074: infc[t] = 'т'; break;
                    case 1075: infc[t] = 'у'; break;
                    case 1076: infc[t] = 'ф'; break;
                    case 1077: infc[t] = 'х'; break;
                    case 1078: infc[t] = 'ц'; break;
                    case 1079: infc[t] = 'ч'; break;
                    case 1080: infc[t] = 'ш'; break;
                    case 1081: infc[t] = 'щ'; break;
                    case 1082: infc[t] = 'ъ'; break;
                    case 1083: infc[t] = 'ы'; break;
                    case 1084: infc[t] = 'ь'; break;
                    case 1085: infc[t] = 'э'; break;
                    case 1086: infc[t] = 'ю'; break;
                    case 1087: infc[t] = 'я'; break;

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

                // вставить MVV_RESPONSE

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

                //SetDocumentStatus(nID, 19);// установить статус Получен ответ для запроса


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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }
        
        // отличается наличием параметров, которые на самом деле нафик тут не нужны, но по аналогии с insert сделано
        private bool UpdateZaprosIn_PK_OSP(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id)
        {

            OleDbCommand cmdInsMVV_RESPONSE;
            OleDbTransaction tran = null;

            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                // вставить MVV_RESPONSE

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

                //SetDocumentStatus(nID, 19);// установить статус Получен ответ для запроса


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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }


        // создать индекс на таблице I_ID
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

                // если флаг удалять учтенные стоит
                if (flDeleteUsed)
                {

                    // удалить все что FL_USE = 1

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
                
                // если флаг удалять старые стоит
                if (flDeleteOld)
                {
                    // Удаление перед сверкой информации о тех ИД, которые были выданы 1 год и 10 дней назад (10 дней на вступление в силу)
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

                // вставить MVV_DATA_RESPONSE

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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

                // вставить MVV_DATA_RESPONSE

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

                //SetDocumentStatus(nID, 19);// установить статус Получен ответ для запроса


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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }

            return false;
        }

        // Инт. таблицы
        // заглушка чтобы работала функция без указания DX_PACK_ID
        private bool InsertResponseIntTable(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id, ref int iRewriteState)
        {
            // string txtAgentCode, string txtAgentDeptCode, string txtAgentAgreementCode, string txtEntityName = пустые тут, вобще-то надо все найти
            return InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, entt_id, ref iRewriteState, 0, " ", " ", " ", " ");

        }

        // Инт. таблицы - тут не надо мучаться с пакетами
        // это с iRewriteState
        // отличается наличием параметров Статус и КодКонтрагента, пишет MVV_EXTERNAL_RESPONSE и out iRewriteState
        private bool InsertResponseIntTable(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id, ref int iRewriteState, decimal nDX_PACK_ID, string txtAgentCode, string txtAgentDeptCode, string txtAgentAgreementCode, string txtEntityName)
        {
            // что такое entt_id - похоже что это код ПФ из справочника Legal
            // вставить EXT_INPUT_HEADER < - > EXT_RESPONSE (связь один к одному)
            OleDbCommand cmd, cmdEXT_INPUT_HEADER, cmdCheckAnsw, cmdEXT_RESPONSE, cmdPackDocs, cmdDocNumber;
            Decimal newID, prevID;
            OleDbTransaction tran = null;
            decimal nAgreementID = 0;
            decimal nAgent_dept_id = 0;
            decimal nAgent_id = 0;

            //iRewriteState = 
            //1 - обычный режим - запрашивать реакцию у пользователя 
            //2 - дописать
            //3 - перезаписать
            //4 - пропустить
            //20 - дописать все
            //21 - перезаписать все
            //22 - пропустить все, которые найдены

            try
            {
                // все - больше не нужно ID, только Code которые в параметрах функции передаем
                // TODO: убрать если потом будет ненужно (всегда 0 - пакет передавать будем)
                // если известно на какой пакет отвечаем, то получаем параметры
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

                // проверить что не было загружено ответа на этот запрос
                // select id from MVV_EXTERNAL_RESPONSE ext left join MVV_RESPONSE resp on ext.id = resp.id left join DOCUMENT doc on ext.id = doc.id where doc.parent_id =:ID and resp.entity_id = :ENTITY_ID
                cmdCheckAnsw = new OleDbCommand("select first 1 ext.id from MVV_EXTERNAL_RESPONSE ext join MVV_RESPONSE resp on ext.id = resp.id join DOCUMENT doc on ext.id = doc.id where doc.parent_id =:ID and resp.entity_id = :ENTITY_ID", con, tran);
                cmdCheckAnsw.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));
                cmdCheckAnsw.Parameters.Add(new OleDbParameter(":ENTITY_ID", Convert.ToDecimal(entt_id)));
                prevID = Convert.ToDecimal(cmdCheckAnsw.ExecuteScalar());

                // тут надо анализировать iRewriteState

                // если запрос неотвеченный - больше не проверяем - грузим в таблицы до упора
                // - пусть пристав потом сам с оператором МВВ разбирается что там за ответ
                //if (prevID <= 0)
                //{

                    // получить новый ключ
                    cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
                    newID = Convert.ToDecimal(cmd.ExecuteScalar());

                    // вставить DOCUMENT
                    cmdEXT_INPUT_HEADER = new OleDbCommand();
                    cmdEXT_INPUT_HEADER.Connection = con;
                    cmdEXT_INPUT_HEADER.Transaction = tran;
                    cmdEXT_INPUT_HEADER.CommandText = "insert into EXT_INPUT_HEADER (ID, METAOBJECTNAME, PROCEED, PACK_NUMBER, EXTERNAL_KEY, AGENT_CODE, AGENT_DEPT_CODE, AGENT_AGREEMENT_CODE, DATE_IMPORT)";
                    cmdEXT_INPUT_HEADER.CommandText += " VALUES (:ID ,'EXT_RESPONSE', 0, :PACK_NUMBER, :EXTERNAL_KEY, :AGENT_CODE, :AGENT_DEPT_CODE, :AGENT_AGREEMENT_CODE, :DATE_IMPORT)";

                    cmdEXT_INPUT_HEADER.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));


                    // 20120828 обрезаю внешний ключ до 8 с конца
                    string txtExtPack = Convert.ToString(nDX_PACK_ID);
                    if (txtExtPack.Length > 8)
                    {
                        txtExtPack = txtExtPack.Substring(txtExtPack.Length - 8, 8);
                    }
                    decimal nExtPack = 0;
                    Decimal.TryParse(txtExtPack, out nExtPack);
                    
                    cmdEXT_INPUT_HEADER.Parameters.Add(new OleDbParameter(":PACK_NUMBER", nExtPack));
                    
                    // переписать - сделать номер пакета через генератор и передавать его, если передали пустой, то тогда делать генерацию
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

                    // вставить MVV_RESPONSE
                    // в 14 релизе добавились еще параметры EXAD_AGENT_ID, EXAD_DEPT_ID, OUTER_AGREEMENT_ID
                    // а в 68-й сборке они куда-то пропали и неизвестно где они

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

                    // TODO: OUTER_AGREEMENT_ID, OUTER_AGREEMENT_NAME - соглашение

                    if (cmdEXT_RESPONSE.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to EXT_RESPONSE table id = " + nID.ToString());
                        throw ex;
                    }

                    tran.Commit();
                    con.Close();

                    // проверить - надо ли это делать, т.к. возможно что система сама потом обновит статус приобработке инт. таблицы
                    // система сама не делает - надо ставить самому
                    SetDocumentStatus(nID, 19);// установить статус Получен ответ для запроса

                    //SetDocumentStatus(nID, 15);// установить статус Обработано с ошибкой
                    
                    return true;
                //}
                //// TODO: решить что делать, если уже есть ответ на этот запрос...
                //else
                //{
                //    tran.Rollback();
                //    con.Close();

                //    // предлагаю просто дописывать - и не придумывать никаких вопросов.
                //    // TODO: кстати - вопрос правильно ли так делать - потому что есть реальная проблема
                //    // - если в ext_response.proceed = 1, то обрабатываться эта строчка больше не будет
                //    // а мы делаем update именно ее, тогда надо и proceed ставить 0
                //    // возможно, статус запроса надо будет поменять..
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            
            // проверить подключение - а то может статься что не закрыли
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

                // вставить MVV_RESPONSE

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

                //SetDocumentStatus(nID, 19);// установить статус Получен ответ для запроса


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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }


        // заглушка чтобы работала функция без указания DX_PACK_ID
        private bool InsertZaprosTo_PK_OSP(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id, ref int iRewriteState)
        {
            
            return InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, entt_id, ref iRewriteState, 0);
        }

        // это с iRewriteState
        // отличается наличием параметров Статус и КодКонтрагента, пишет MVV_EXTERNAL_RESPONSE и out iRewriteState
        // разобраться с пакетом
        private bool InsertZaprosTo_PK_OSP(OleDbConnection con, decimal nID, string txtOtvet, DateTime dtDatOtv, decimal nStatus, decimal entt_id, ref int iRewriteState, decimal nDX_PACK_ID)
        {

            OleDbCommand cmd, cmdMVV_I, cmdCheckAnsw, cmdInsDoc, cmdInsMVV_RESPONSE, cmdInsMVV_EXTERNAL_RESPONSE, cmdPackDocs;
            Decimal newID, prevID;
            OleDbTransaction tran = null;
            decimal nAgreementID = 0;
            decimal nAgent_dept_id = 0;
            decimal nAgent_id = 0;

            //iRewriteState = 
            //1 - обычный режим - запрашивать реакцию у пользователя 
            //2 - дописать
            //3 - перезаписать
            //4 - пропустить
            //20 - дописать все
            //21 - перезаписать все
            //22 - пропустить все, которые найдены

            try
            {

                // если известно на какой пакет отвечаем, то получаем параметры
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

                // проверить что не было загружено ответа на этот запрос
                // select id from MVV_EXTERNAL_RESPONSE ext left join MVV_RESPONSE resp on ext.id = resp.id left join DOCUMENT doc on ext.id = doc.id where doc.parent_id =:ID and resp.entity_id = :ENTITY_ID
                cmdCheckAnsw = new OleDbCommand("select first 1 ext.id from MVV_EXTERNAL_RESPONSE ext join MVV_RESPONSE resp on ext.id = resp.id join DOCUMENT doc on ext.id = doc.id where doc.parent_id =:ID and resp.entity_id = :ENTITY_ID", con, tran);
                cmdCheckAnsw.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));
                cmdCheckAnsw.Parameters.Add(new OleDbParameter(":ENTITY_ID", Convert.ToDecimal(entt_id)));
                prevID = Convert.ToDecimal(cmdCheckAnsw.ExecuteScalar());

                // тут надо анализировать iRewriteState

                if (prevID <= 0)
                {

                    // получить новый ключ
                    cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
                    newID = Convert.ToDecimal(cmd.ExecuteScalar());

                    // вставить DOCUMENT
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

                    // TODO: вставить MVV_I c 14 релиза

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


                    // вставить MVV_RESPONSE
                    // в 14 релизе добавились еще параметры EXAD_AGENT_ID, EXAD_DEPT_ID, OUTER_AGREEMENT_ID
                    // а в 68-й сборке они куда-то пропали и неизвестно где они

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

                    // TODO: OUTER_AGREEMENT_ID, OUTER_AGREEMENT_NAME - соглашение

                    if (cmdInsMVV_RESPONSE.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to MVV_RESPOSNSE table id = " + nID.ToString());
                        throw ex;

                    }

                    
                    // вставить MVV_EXTERNAL_RESPONSE, вставить параметр DX_PACK_ID


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

                    // если вставляем в пакет, до пишем в промежуточную таблицу
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

                    SetDocumentStatus(nID, 19);// установить статус Получен ответ для запроса
                    //SetDocumentStatus(nID, 15);// установить статус Обработано с ошибкой
                    
                    return true;
                }
                else
                {
                    tran.Rollback();
                    con.Close();

                    // тут если все как обычно, то запускаем диалог frmRewriteDialog.ShowForm(), 
                    // а если iRewriteState > 1, то разбираемся
                    if (iRewriteState == 1)
                    {
                        frmRewriteDialog frmRwD = new frmRewriteDialog();
                        iRewriteState = frmRwD.ShowForm();
                    }

                    // пытаемся перзаписать, дописать, пропустить
                    switch (iRewriteState)
                    {
                        //2 - дописать
                        case (2):
                            if (AppendZaprosIn_PK_OSP(con, prevID, txtOtvet, dtDatOtv, nStatus, entt_id))
                            {
                                iRewriteState = 1;
                                return true;
                            }
                            
                            break;

                        //3 - перезаписать
                        case (3):
                            
                            if (UpdateZaprosIn_PK_OSP(con, prevID, txtOtvet, dtDatOtv, nStatus, entt_id))
                            {
                                iRewriteState = 1;
                                return true;
                            }
                            break;
                        
                        // 4 - пропустить
                        case (4):
                            iRewriteState = 1;
                            break;

                        //20 - дописать все
                        case (20):
                            if (AppendZaprosIn_PK_OSP(con, prevID, txtOtvet, dtDatOtv, nStatus, entt_id))
                            {
                                return true;
                            }
                            break;
                        //21 - перезаписать все
                        case (21):
                            if (UpdateZaprosIn_PK_OSP(con, prevID, txtOtvet, dtDatOtv, nStatus, entt_id))
                            {
                                return true;
                            }
                            break;
                            //22 - пропустить все
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }

        // создать входящий пакет DX_PACK_I
        private decimal ID_CreateDX_PACK_I(OleDbConnection con, decimal nStatus, decimal nAGENT_ID, decimal nAGENT_DEPT_ID, decimal nAGREEMENT_ID, string txtPLAIN_LOG, string txtAgent_code, string txtAgreement_code, string txtAgent_dept_code)
        {
            // 1- новый
            // 70 - обработан
            // 71 - обработан с ошибками
            decimal nID = 0;
            OleDbCommand cmd, cmdInsDoc, cmdDX_PACK, cmdDX_PACK_I, cmdPACK_LOGS;
            OleDbTransaction tran = null;

            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                // TODO: вставить DX_PACK_I

                // получить новый ключ
                cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
                nID = Convert.ToDecimal(cmd.ExecuteScalar());



                // вставить DOCUMENT
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
                // добавил в 163 сборке 15 релиза
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

                // узнать параметры ИП по ИД

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

                // получить новый ключ
                cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
                newID = Convert.ToDecimal(cmd.ExecuteScalar());

                // вставить DOCUMENT
                cmdInsDoc = new OleDbCommand();
                cmdInsDoc.Connection = con;
                cmdInsDoc.Transaction = tran;
                cmdInsDoc.CommandText = "insert into DOCUMENT (ID, METAOBJECTNAME, DOCSTATUSID, DOCUMENTCLASSID, CREATE_DATE, SUSER_ID)";
                cmdInsDoc.CommandText += " VALUES (:ID, 'I_IP_OTHER', :DOCSTATUSID, :DOCUMENTCLASSID, :CREATE_DATE, :SUSER_ID)";

                cmdInsDoc.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));

                //cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(1)));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(nStatus)));

                cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCUMENTCLASSID", Convert.ToDecimal(11))); // класс документооборота для объекта I - Ыходящий документ
                //cmdInsDoc.Parameters.Add(new OleDbParameter(":PARENT_ID", Convert.ToDecimal(nID)));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":CREATE_DATE", DateTime.Now));
                cmdInsDoc.Parameters.Add(new OleDbParameter(":SUSER_ID", Convert.ToDecimal(nUserID)));
                //cmdInsDoc.Parameters.Add(new OleDbParameter(":AMOUNT", Convert.ToDouble(nAmount)));


                if (cmdInsDoc.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to document table parent_id = " + newID.ToString());
                    throw ex;
                }

                // вставить I

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


                // вставить I_IP


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
                cmdInsI_IP_OTHER.Parameters.Add(new OleDbParameter(":INDOC_TYPE_NAME", Convert.ToString("Сопроводительное письмо")));
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                if (con != null)
                {
                    con.Close();
                }
            }
            return -1;
        }
        
        // добавить депозитный документ о погашении долга 
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
                // самое лучшее решение - это выбрать contr_id из i_id и передать сюда

                dsIP_params = new DataSet();
                dtIP_params = dsIP_params.Tables.Add("IP_params");
                newID = 0;
                prevID = 0;
                id_dbtr = 0;

                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();

                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                // узнать параметры ИП по ИД

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
                    txtContrName = Convert.ToString(dsIP_params.Tables[0].Rows[0]["id_dbtr_name"]); // теперь отправителем будет сам должник
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

                // получить новый ключ
                cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
                newID = Convert.ToDecimal(cmd.ExecuteScalar());

                // вставить DOCUMENT
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

                // вставить I
                // - Отправитель 	I.CONTR_NAME
                // - Адрес отправителя I.ADR
                

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


                    // вставить I_IP


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


                    // вставить I_DEPOSIT 

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

                    // вставить I_OP_CS

                    cmdInsI_OP_CS  = new OleDbCommand();
                    cmdInsI_OP_CS.Connection = con;
                    cmdInsI_OP_CS.Transaction = tran;
                    cmdInsI_OP_CS.CommandText = "insert into I_OP_CS (ID, CHANGEDBT_REASON_ID, CHANGEDBT_REASON_DESCR, I_OP_CS_CHANGESUM)";
                    cmdInsI_OP_CS.CommandText += "  VALUES (:ID, 3, 'Оплата штрафа в ГИБДД', :I_OP_CS_CHANGESUM)";
                    cmdInsI_OP_CS.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                    cmdInsI_OP_CS.Parameters.Add(new OleDbParameter(":I_OP_CS_CHANGESUM", Convert.ToDouble(nAmount)));

                    if (cmdInsI_OP_CS.ExecuteNonQuery() == -1)
                    {
                        Exception ex = new Exception("Error inserting new row to I_OP_CS  table id = " + newID.ToString());
                        throw ex;

                    }


                    // вставить I_OP_CS_ENDDBT

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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

            //с помощью специальной функции, анализируя дату и номер ид найти ИП на статусе в исполнении

            // в какой таблице лучше искать? join doc_ip_doc, doc_ip, document
            // where (d.docstatusid ! = -1) and (d.docstatusid ! = 7) and (d.docstatusid ! = 10) = !(Удален, Отказан, Окончен)
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
            int iRewriteState = 1; // обычный режим перезаписи ответов на запрос (запрашивать действия у пользователя)
            decimal nAgreementID = 0;
            decimal nAgent_dept_id = 0;
            decimal nAgent_id = 0;
            decimal nDx_pack_id = 0;
            decimal nNewPackID = 0;

            string txtAgreementCode = "";
            string txtAgentCode = "";
            string txtAgentDeptCode = "";
            string txtEntityName = "";
            bool bNotIntTablesResp = false; // если ответ на запрос, сделанный без интерфейсных таблиц.

            if (cbxOrg.SelectedValue != null)
            {
                // нужно в базу ПК ОСП добавить контрагентов, их ID внести в список
                org = Convert.ToDecimal(cbxOrg.SelectedValue);

                openFileDialog1.Filter = "DBF файлы(*.dbf)|*.dbf";
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
                                m_cmd.CommandText = "SELECT * FROM NOFIND ORDER BY ZAPROS";// упорядочили по полю ZAPROS чтобы работать с несколькими ответами на 1 запрос
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


                                // если файл с ответами не пустой, то по первому ответу определить
                                // этb ответы на запросы из 14 релиза или старые (без пакетов)
                                if (tbl.Rows.Count > 0)
                                {
                                    decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["ZAPROS"]);

                                    if (FindSendlist(nFirstID, org)) // указываем параметр org - контрагент из списка рассылки, которому была направлена копия
                                    {
                                        // значит это новый запрос
                                        // получить параметры: соглашение, контрагент, подразделение
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
                                        nAgreementID = GetAgr_by_Org(org); // номер соглашения
                                        nAgent_id = GetAgent_ID(nAgreementID);
                                        nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                    }

                                    txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                    txtEntityName = GetLegal_Name(org);

                                    // нужно создать новый входящий пакет
                                    //nNewPackID = ID_CreateDX_PACK_I(con, 1, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);

                                    // TODO: показать форму выбора запроса, к которому крепим ответ
                                    frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD, "FL_NOFIND");
                                    nParentID = LogList.ShowForm();

                                    // если не было выбрано пропустить загрузку ответа
                                    if (nParentID != -1)
                                    {
                                        // 1 - Новый
                                        // 4 - Ответ отрицательный
                                        nNewPackID = CreateLLog(conGIBDD, 1, 4, txtAgreementCode, nParentID, "Пакет ответов из " + txtEntityName + ".");



                                        // записать в лог пакета дату и начало обработки
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                        //WritePackLog(con, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");


                                        foreach (DataRow row in tbl.Rows)
                                        {
                                            //m_cmd = new OleDbCommand();
                                            //m_cmd.Connection = con;
                                            //m_cmd.Transaction = tran;
                                            // вот тут-то надо срочно написать функцию создания нового ответа
                                            // Необходимо по коду (ZAPROS)  найти документ в базе DOCUMENT.ID
                                            nStatus = 7; // нет данных
                                            txtID = Convert.ToString(row["ZAPROS"]);
                                            if (!Decimal.TryParse(txtID, out nID))
                                            {
                                                nID = 0;
                                            }
                                            if (FindZapros(nID))
                                            {
                                                // значить начинаем вставлять в базу структуры данных ответа
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

                                                    // проверить дату запроса
                                                    txtDatZap = Convert.ToString(row["DATZPR2"]);
                                                    if (!DateTime.TryParse(txtDatZap, out dtDatZap))
                                                    {
                                                        dtDatZap = DateTime.MaxValue;
                                                    }

                                                    bNotIntTablesResp = false; // теперь все ответы точно из интерфейсных таблиц
                                                    //if (dtDatZap < dtIntTablesDeplmntDate)
                                                    //{
                                                    //    bNotIntTablesResp = true;
                                                    //}
                                                    //else
                                                    //{
                                                    //    bNotIntTablesResp = false;
                                                    //}


                                                    string txtOtvet;

                                                    // вот сюда и воткнуть ФИО и год рождения, адрес
                                                    string txtResLine = Convert.ToString(row["FIO"]).TrimEnd();
                                                    if (row["GODR"] != System.DBNull.Value)
                                                    {
                                                        txtResLine += " (" + Convert.ToInt32(row["GODR"]).ToString() + " г.р.)";
                                                    }
                                                    txtResLine += " " + Convert.ToString(row["ADRES"]).TrimEnd();

                                                    txtOtvet = "в соответствии с " + PKOSP_GetOrgConvention(org);
                                                    txtOtvet += " получен ответ: ";

                                                    txtOtvet += "Ответ из " + GetLegal_Name(org) + ". Нет данных о должнике " + txtResLine + ". Дата ответа: " + dtDatOtv.ToShortDateString();

                                                    // iRewriteState
                                                    // 1 - обычный режим - запрашивать реакцию у пользователя 
                                                    // 2 - дописать все
                                                    // 3 - перезаписать все4 - пропустить все, которые найдены
                                                    if (bNotIntTablesResp)
                                                    {
                                                        if (InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID))
                                                        {
                                                            iCnt++;
                                                            WritePackLog(con, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");
                                                            prbWritingDBF.PerformStep();
                                                            prbWritingDBF.Refresh();
                                                            System.Windows.Forms.Application.DoEvents();
                                                        }
                                                        else
                                                        {
                                                            // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                            WritePackLog(con, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                            nStatus = 15; // ошибка
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                        {
                                                            iCnt++;
                                                            // WritePackLog(con, nNewPackID, "Обработан ответ # " + nID.ToString() + "\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");

                                                            prbWritingDBF.PerformStep();
                                                            prbWritingDBF.Refresh();
                                                            System.Windows.Forms.Application.DoEvents();
                                                        }
                                                        else
                                                        {
                                                            // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                            //WritePackLog(con, nNewPackID, "Ошибка! Ответ # " + nID.ToString() + " обработать не удалось.\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                            nStatus = 15; // ошибка
                                                        }
                                                    }

                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                                                    if (nNewPackID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "Ошибка! Загрузка пакета ответов экстренно прервалась.\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Значение счетчика = " + iCnt.ToString() + "\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");

                                                        if (nID > 0)
                                                        {
                                                            WriteLLog(conGIBDD, nNewPackID, "ID запроса = " + nID.ToString() + "\n");
                                                        }
                                                    }

                                                }
                                            }
                                            else
                                            {
                                                // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                if (nNewPackID > 0)
                                                {
                                                    WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос ID = " + nID.ToString() + " не удалось загрузить т.к. не обнаружен запрос-родитель.\n");
                                                }
                                            }
                                        }
                                        //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                        //WritePackLog(con, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                        // установить количество обработанных запросов
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                        // обновить статус лога-ответа
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                        // обновить статус лога-родителя
                                        // сразу меняем статус, т.к. фильтр все равно по флагам
                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - загружен ответ
                                        // обновить флаг что обработан NOFIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_NOFIND");


                                        // закомментировать т.к. теперь никакого пакета не создается
                                        //// если все ок, то нужно поменять статус пакета
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
                                        nAgreementID = GetAgr_by_Org(org); // номер соглашения
                                        nAgent_id = GetAgent_ID(nAgreementID);
                                        nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                    }


                                    txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                    txtEntityName = GetLegal_Name(org);

                                    // создать лог отрицательных ответов и записать туда что 0 в пакете ответов
                                    frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD, "FL_NOFIND");
                                    nParentID = LogList.ShowForm();

                                    // если не было выбрано пропустить загрузку ответа
                                    if (nParentID != -1)
                                    {
                                        // 1 - Новый
                                        // 4 - Ответ отрицательный
                                        nNewPackID = CreateLLog(conGIBDD, 1, 4, txtAgreementCode, nParentID, "Пакет ответов из " + txtEntityName + ".");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");
                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                        // установить количество обработанных запросов
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                        // обновить статус лога-ответа
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                        // обновить статус лога-родителя
                                        // сразу меняем статус, т.к. фильтр все равно по флагам
                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - загружен ответ
                                        // обновить флаг что обработан NOFIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_NOFIND");
                                    }
                                }
   
                                    
                                MessageBox.Show("Обработано ответов: " + iCnt.ToString() + ".\nСейчас будет сформирован реестр ответов.", "Сообщение", MessageBoxButtons.OK);
#region "REESTR"
                                //**********Формирование**реестра**nofind************
                                //для разделения по приставам. 

                                //список всех приставов
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
                                // пишу обработчик для HTML
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
                                    report.AddToReport("Реестр ответов на запросы СП-И о наличии ден. средств из кредитной организации");
                                    report.AddToReport(GetLegal_Name(org) + " от " + Convert.ToDateTime(tbl.Rows[0]["DATOTV"]).ToShortDateString() + "<br>");
                                    //report.AddToReport("За период с " + Convert.ToDateTime(tbl.Rows[0]["DATZPR1"]).ToShortDateString() + " по " + Convert.ToDateTime(tbl.Rows[0]["DATZPR2"]).ToShortDateString() + "<br>");
                                    report.AddToReport("Нет данных о наличии счетов у должников<br>");

                                    spi = Convert.ToInt32(drspi["NOMSPI"]);
                            
                                    report.AddToReport("СП-И: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "<br>");
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
                                //    //      пример для Ворда
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

                                //        par.Range.Text += "Реестр ответов на запросы СП-И о наличии ден. средств из банка";
                                //        par.Range.Text += GetLegal_Name(org) + " от " + Convert.ToDateTime(tbl.Rows[0]["DATOTV"]).ToShortDateString() + "\n";
                                //        par.Range.Text += "За период с " + Convert.ToDateTime(tbl.Rows[0]["DATZPR1"]).ToShortDateString() + " по " + Convert.ToDateTime(tbl.Rows[0]["DATZPR2"]).ToShortDateString() + "\n";
                                //        par.Range.Text += "Нет данных о наличии счетов у должников\n";

                                //        spi = Convert.ToInt32(drspi["NOMSPI"]);

                                //        //par.Range.Text += "СП-И: " + GetSpiName3(Convert.ToInt32(drspi["NOMSPI"])) + "\n";
                                //        par.Range.Text += "СП-И: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";
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
                                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                                }
                                //return false;
                            }
                            catch (Exception ex)
                            {
                                //if (DBFcon != null) DBFcon.Close();
                                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                                //return false;
                            }
                            //return true;
                        }
                        # endregion
                        #region "E_TOFIND"
                        else if(openFileDialog1.FileName.ToLower().Contains("e_tofind.dbf"))
                        {
                            // это файл с ошибками
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
                                // сделать выборку по таблице tofind
                                m_cmd.CommandText = "select  distinct  NOMOSP, LITZDOLG, FIOVK, FIOVK, ZAPROS, GOD, NOMSPI, NOMIP, SUMMA, VIDVZISK, INNORG, DATZAPR, ADDR, FLZPRSPI, DATZAPR1, DATZAPR2, FL_OKON, OSNOKON, OSNOKON from E_TOFIND ORDER BY ZAPROS";// упорядочили по полю ZAPROS чтобы работать с несколькими ответами на 1 запрос - тут не надо по логике, но кто знает что там в ответе
                                //m_cmd.CommandText = "SELECT * FROM FIND";// упорядочили по полю ZAPROS чтобы работать с несколькими ответами на 1 запрос

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
                                nStatus = 15; // Ошибка

                                nAgreementID = 0;
                                nAgent_dept_id = 0;
                                nAgent_id = 0;
                                nDx_pack_id = 0;
                                nNewPackID = 0;

                                txtAgreementCode = "";
                                txtAgentCode = "";
                                txtAgentDeptCode = "";
                                txtEntityName = "";


                                // если файл с ответами не пустой, то по первому ответу определить
                                // этb ответы на запросы из 14 релиза или старые (без пакетов)
                                if (tbl.Rows.Count > 0)
                                {
                                    decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["ZAPROS"]);
                                    if (FindSendlist(nFirstID, org)) // указываем параметр org - контрагент из списка рассылки, которому была направлена копия
                                    {
                                        // значит это новый запрос
                                        // получить параметры: соглашение, контрагент, подразделение
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
                                        nAgreementID = GetAgr_by_Org(org); // номер соглашения
                                        nAgent_id = GetAgent_ID(nAgreementID);
                                        nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                    }


                                    txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                    txtEntityName = GetLegal_Name(org);



                                    // нужно создать новый входящий пакет
                                    // nNewPackID = ID_CreateDX_PACK_I(con, 1, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);
                                    // TODO: показать форму выбора запроса, к которому крепим ответ
                                    frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD, "FL_E_TOFIND");
                                    nParentID = LogList.ShowForm();

                                    if (nParentID != -1)
                                    {
                                        // 1 - Новый
                                        // 5 - Отказано в обработке запроса
                                        nNewPackID = CreateLLog(conGIBDD, 1, 5, txtAgreementCode, nParentID, "Пакет ответов из " + txtEntityName + ".");

                                        // записать в лог пакета дату и начало обработки
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                        //WritePackLog(con, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");

                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");


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
                                                // значить начинаем вставлять в базу структуры данных ответа
                                                try
                                                {
                                                    string txtDatOtv = "";
                                                    DateTime dtDatOtv;

                                                    txtDatOtv = Convert.ToString(row["DATZAPR2"]); // это дата выгрузки запроса - пусть она и будет датой ответа о непринятии запроса
                                                    if (!DateTime.TryParse(txtDatOtv, out dtDatOtv))
                                                    {
                                                        dtDatOtv = DateTime.MaxValue;
                                                    }

                                                    string txtDatZap = "";
                                                    DateTime dtDatZap;


                                                    // проверить дату запроса
                                                    txtDatZap = Convert.ToString(row["DATZAPR"]);
                                                    if (!DateTime.TryParse(txtDatZap, out dtDatZap))
                                                    {
                                                        dtDatZap = DateTime.MaxValue;
                                                    }

                                                    bNotIntTablesResp = false; // теперь все ответы точно из интерфейсных таблиц
                                                    //if (dtDatZap < dtIntTablesDeplmntDate)
                                                    //{
                                                    //    bNotIntTablesResp = true;
                                                    //}
                                                    //else
                                                    //{
                                                    //    bNotIntTablesResp = false;
                                                    //}

                                                    string txtOtvet;
                                                    txtOtvet = "в соответствии с " + PKOSP_GetOrgConvention(org);
                                                    txtOtvet += " получен ответ: ";

                                                    txtOtvet += "Ответ из " + GetLegal_Name(org);
                                                    txtOtvet += ". Запрос не принят в обработку. Ошибка в данных запроса. Необходимо проверить кооректность реквизитов ИП. Для физ. лиц: ФИО должника (полное), дата рождения должника. Для юр. лиц: ИНН, наименование должника.\n.";

                                                    //if (InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, org))
                                                    if (bNotIntTablesResp)
                                                    {
                                                        if (InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID))
                                                        {
                                                            iCnt++;
                                                            WritePackLog(con, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");
                                                            prbWritingDBF.PerformStep();
                                                            prbWritingDBF.Refresh();
                                                            System.Windows.Forms.Application.DoEvents();
                                                        }
                                                        else
                                                        {
                                                            // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                            WritePackLog(con, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                        {
                                                            iCnt++;
                                                            // WritePackLog(con, nNewPackID, "Обработан ответ # " + nID.ToString() + "\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");

                                                            prbWritingDBF.PerformStep();
                                                            prbWritingDBF.Refresh();
                                                            System.Windows.Forms.Application.DoEvents();
                                                        }
                                                        else
                                                        {
                                                            // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                            // WritePackLog(con, nNewPackID, "Ошибка! Ответ # " + nID.ToString() + " обработать не удалось.\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                        }
                                                    }

                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                                                    if (nNewPackID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "Ошибка! Загрузка пакета ответов экстренно прервалась.\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Значение счетчика = " + iCnt.ToString() + "\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");

                                                        if (nID > 0)
                                                        {
                                                            WriteLLog(conGIBDD, nNewPackID, "ID запроса = " + nID.ToString() + "\n");
                                                        }
                                                    }

                                                }
                                            }
                                            else
                                            {
                                                // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                if (nNewPackID > 0)
                                                {
                                                    WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос ID = " + nID.ToString() + " не удалось загрузить т.к. не обнаружен запрос-родитель.\n");
                                                }
                                            }

                                        }
                                        //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                        //WritePackLog(con, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                        // установить количество обработанных запросов
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);


                                        // обновить статус лога-ответа
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);


                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - загружен ответ

                                        // обновить флаг что обработан E_TOFIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_E_TOFIND");

                                        // закомментировать т.к. теперь никакого пакета не создается
                                        //// если все ок, то нужно поменять статус пакета
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
                                        nAgreementID = GetAgr_by_Org(org); // номер соглашения
                                        nAgent_id = GetAgent_ID(nAgreementID);
                                        nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                    }


                                    txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                    txtEntityName = GetLegal_Name(org);

                                    // создать лог отрицательных ответов и записать туда что 0 в пакете ответов
                                    frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD, "FL_E_TOFIND");
                                    nParentID = LogList.ShowForm();

                                    // если не было выбрано пропустить загрузку ответа
                                    if (nParentID != -1)
                                    {
                                        // 1 - Новый
                                        // 5 - Ответ не взяты в обработку
                                        nNewPackID = CreateLLog(conGIBDD, 1, 5, txtAgreementCode, nParentID, "Пакет ответов из " + txtEntityName + ".");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");
                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                        // установить количество обработанных запросов
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                        // обновить статус лога-ответа
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                        // обновить статус лога-родителя
                                        // сразу меняем статус, т.к. фильтр все равно по флагам
                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - загружен ответ
                                        // обновить флаг что обработан E_TOFIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_E_TOFIND");
                                    }
                                }

                                MessageBox.Show("Обработано ответов: " + iCnt.ToString() + ".\nСейчас будет сформирован реестр ответов.", "Сообщение", MessageBoxButtons.OK);

                                //**********Формирование**реестра**find************
                                //Надо вспомнить запись в Ворд + подсчитывая количество строк 
                                //для разделения по приставам. 

                                //список всех приставов
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
                                // пишу обработчик для HTML
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
                                    report.AddToReport("Реестр не принятых в обработку запросов СП-И о наличии ден. средств из банка<br>");
                                    report.AddToReport("" + GetLegal_Name(org) + " от " + DateTime.Today.ToShortDateString() + "<br>");
                                    //report.AddToReport("За период с " + Convert.ToDateTime(ds.Tables["E_TOFIND"].Rows[0]["DATZAPR1"]).ToShortDateString() + " по " + Convert.ToDateTime(ds.Tables["E_TOFIND"].Rows[0]["DATZAPR1"]).ToShortDateString() + "<br>");

                                    spi = Convert.ToInt32(drspi["NOMSPI"]);

                                    report.AddToReport("СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "<br>");
                                                                        
                                    report.AddToReport("</h3>");

                                    foreach (DataRow row in tbl.Rows)
                                    {
                                        report.AddToReport("<p>");
                                        if (spi == Convert.ToInt32(row["NOMSPI"]))
                                        {
                                            string txtResLine = GetIPNum(con, Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIOVK"]).TrimEnd();
                                            if (Convert.ToInt32(row["LITZDOLG"]) == 2)
                                            {
                                                txtResLine += " (" + Convert.ToInt32(row["GOD"]).ToString() + " г.р.)";
                                                //txtResLine += " (" + Convert.ToInt32(row["GODR"]).ToString() + " г.р.)";
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
                                //    // для Ворда

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

                                //        par.Range.Text += "Реестр не принятых в обработку запросов СП-И о наличии ден. средств из банка\n";
                                //        par.Range.Text += GetLegal_Name(org) + " от " + DateTime.Today.ToShortDateString() + "\n";
                                //        par.Range.Text += "За период с " + Convert.ToDateTime(tbl.Rows[0]["DATZAPR1"]).ToShortDateString() + " по " + Convert.ToDateTime(tbl.Rows[0]["DATZAPR2"]).ToShortDateString() + "\n";

                                //        spi = Convert.ToInt32(drspi["NOMSPI"]);

                                //        sch_line = 0;
                                //        if (fl_fst == 1)
                                //        {
                                //            sch_line = 1;
                                //            fl_fst = 0;
                                //        }
                                //        par.Range.Text += PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";

                                //        //par.Range.Text += "НОМЕР ИП             ДОЛЖНИК                            АДРЕС                      СЧЕТ       ОСТАТОК\n";
                                //        par.Range.Text += "НОМЕР ИП             ДОЛЖНИК      ГОД РОЖДЕНИЯ\n";

                                //        sch_line += 10;

                                //        foreach (DataRow row in tbl.Rows)
                                //        {
                                //            if (spi == Convert.ToInt32(row["NOMSPI"]))
                                //            {
                                //                string txtResLine = GetIPNum(Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIOVK"]).TrimEnd();
                                //                if (Convert.ToInt32(row["LITZDOLG"]) == 2)
                                //                {
                                //                    txtResLine += " (" + Convert.ToInt32(row["GOD"]).ToString() + " г.р.)";
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
                                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                                }
                                //return false;
                            }
                            catch (Exception ex)
                            {
                                //if (DBFcon != null) DBFcon.Close();
                                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                                //m_cmd.CommandText = "SELECT FIL, PRIZ, LITZDOLG, FIO, ADRES, NOMLS, PRIZS, OSTAT, RSCHET, OSTSCH, NOMOSP, ZAPROS FROM FIND ORDER BY ZAPROS";// упорядочили по полю ZAPROS чтобы работать с несколькими ответами на 1 запрос
                                m_cmd.CommandText = "select distinct  fil, priz, litzdolg, fio, godr, adres, nomls, prizs, ostat, rschet, ostsch, nomosp, zapros, nomspi, nomip, datotv, flzprspi, datzpr1, datzpr2 from FIND ORDER BY ZAPROS";// упорядочили по полю ZAPROS чтобы работать с несколькими ответами на 1 запрос
                                //m_cmd.CommandText = "SELECT * FROM FIND";// упорядочили по полю ZAPROS чтобы работать с несколькими ответами на 1 запрос

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
                                nStatus = 20; // ответ получен

                                // Сделать преамбулу txtPreamb
                                string txtPreamb = "в соответствии с " + PKOSP_GetOrgConvention(org);
                                txtPreamb += " получен ответ: ";
                                // txtPreamb += "Ответ из " + GetLegal_Name(org);

                                // Сделать общую строчку txtCommonRespText
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

                                // если файл с ответами не пустой, то по первому ответу определить
                                // этb ответы на запросы из 14 релиза или старые (без пакетов)
                                if (tbl.Rows.Count > 0)
                                {
                                    decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["ZAPROS"]);
                                    if (FindSendlist(nFirstID, org)) // указываем параметр org - контрагент из списка рассылки, которому была направлена копия
                                    {
                                        // значит это новый запрос
                                        // получить параметры: соглашение, контрагент, подразделение
                                        DataTable dtParams = GetPackParams(con, nFirstID, org);
                                        if ((dtParams != null) && (dtParams.Rows.Count > 0))
                                        {
                                            nAgreementID = Convert.ToDecimal(dtParams.Rows[0]["agreement_id"]);
                                            nAgent_dept_id = Convert.ToDecimal(dtParams.Rows[0]["agent_dept_id"]);
                                            nAgent_id = Convert.ToDecimal(dtParams.Rows[0]["agent_id"]);
                                            nDx_pack_id = Convert.ToDecimal(dtParams.Rows[0]["dx_pack_id"]);
                                        }
                                    }

                                    // нужно создать новый входящий пакет
                                    if (nAgreementID == 0)
                                    {
                                        nAgreementID = GetAgr_by_Org(org); // номер соглашения
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
                                        // 1 - Новый
                                        // 3 - Ответ положительный
                                        nNewPackID = CreateLLog(conGIBDD, 1, 3, txtAgreementCode, nParentID, "Пакет ответов из " + txtEntityName + ".");


                                        // записать в лог пакета дату и начало обработки
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                        //WritePackLog(con, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");


                                        for (i = 0; i < tbl.Rows.Count; i++)
                                        {
                                            DataRow row = tbl.Rows[i];

                                            // Сделать общую строчку txtCommonRespText
                                            // Сделать текущую строчку txtCurrRowRespText
                                            // Если это не последняя строчка и след. строка содержит продолжение ответа, то 
                                            // txtCommonRespText += txtCurrRowRespText;
                                            // Иначе - txtTotalRespText = txtPreamb + txtResponseHeader + txtCommonRespText;
                                            // Записать txtTotalRespText
                                            // в результате избегаем UPDATE для записи многострочных ответов.



                                            txtID = Convert.ToString(row["ZAPROS"]);
                                            if (!Decimal.TryParse(txtID, out nID))
                                            {
                                                nID = 0;
                                            }

                                            // проверить - был есть ли такой запрос вобще FindZapros(nID)
                                            // проверить, это старый или новый запрос - FindSendlist(nID)

                                            if (FindZapros(nID))
                                            {
                                                // значить начинаем вставлять в базу структуры данных ответа
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


                                                    // проверить дату запроса
                                                    txtDatZap = Convert.ToString(row["DATZPR2"]);
                                                    if (!DateTime.TryParse(txtDatZap, out dtDatZap))
                                                    {
                                                        dtDatZap = DateTime.MaxValue;
                                                    }


                                                    bNotIntTablesResp = false; // теперь все ответы точно из интерфейсных таблиц
                                                    //if (dtDatZap < dtIntTablesDeplmntDate)
                                                    //{
                                                    //    bNotIntTablesResp = true;
                                                    //}
                                                    //else
                                                    //{
                                                    //    bNotIntTablesResp = false;
                                                    //}

                                                    // Сделать текущую строчку txtCurrRowRespText
                                                    string txtCurrRowRespText = "";

                                                    // вот сюда и воткнуть ФИО и год рождения, адрес
                                                    txtCurrRowRespText += Convert.ToString(row["FIO"]).TrimEnd();
                                                    if (Convert.ToInt32(row["LITZDOLG"]) == 2)
                                                    {
                                                        txtCurrRowRespText += " (" + Convert.ToInt32(row["GODR"]).ToString() + " г.р.)";
                                                    }
                                                    txtCurrRowRespText += " " + Convert.ToString(row["ADRES"]).TrimEnd() + " ";

                                                    string priz = Convert.ToString(row["PRIZ"]).TrimEnd();
                                                    if (priz.Length > 0) txtCurrRowRespText += Convert.ToString(row["PRIZ"]).TrimEnd();

                                                    if ((row.Table.Columns.Contains("NOMLS")) && (row.Table.Columns.Contains("OSTAT")) && (Convert.ToString(row["NOMLS"]).TrimEnd() != ""))
                                                    {
                                                        string txtLs = Convert.ToString(row["NOMLS"]).TrimEnd();
                                                        txtCurrRowRespText += "л/с: " + txtLs + " остаток = " + Convert.ToDecimal(row["OSTAT"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtLs);
                                                    }

                                                    if ((row.Table.Columns.Contains("RSCHET")) && (row.Table.Columns.Contains("OSTSCH")) && (Convert.ToString(row["RSCHET"]).TrimEnd() != ""))
                                                    {
                                                        string txtRs = Convert.ToString(row["RSCHET"]).TrimEnd();
                                                        txtCurrRowRespText += "р/с: " + txtRs + " остаток = " + Convert.ToDecimal(row["OSTSCH"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtRs);
                                                    }

                                                    // Если это не последняя строчка и след. строка содержит продолжение ответа, то 
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

                                                            // вставить в базу - уточнить Rewrite State
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
                                                                    //WritePackLog(con, nNewPackID, "Обработан ответ # " + nID.ToString() + "\n");
                                                                    WriteLLog(conGIBDD, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");
                                                                    prbWritingDBF.PerformStep();
                                                                    prbWritingDBF.Refresh();
                                                                    System.Windows.Forms.Application.DoEvents();
                                                                }
                                                                else
                                                                {
                                                                    //WritePackLog(con, nNewPackID, "Ошибка! Ответ # " + nID.ToString() + " обработать не удалось.\n");
                                                                    WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                                }
                                                            }
                                                            // приводим строчку в исходное состояние
                                                            txtCommonRespText = "";
                                                        }

                                                    }
                                                    else
                                                    {
                                                        // значит это последняя строчка.
                                                        bMoreTanOne = false;

                                                        txtCommonRespText += txtCurrRowRespText;

                                                        txtCommonRespText = txtPreamb + " " + txtCommonRespText;

                                                        // вставить в базу - уточнить Rewrite State
                                                        if (bNotIntTablesResp)
                                                        {
                                                            if (InsertZaprosTo_PK_OSP(con, nID, txtCommonRespText, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID))
                                                            {
                                                                iCnt++;
                                                                // WritePackLog(con, nNewPackID, "Обработан ответ # " + nID.ToString() + "\n");
                                                                WriteLLog(conGIBDD, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");
                                                                prbWritingDBF.PerformStep();
                                                                prbWritingDBF.Refresh();
                                                                System.Windows.Forms.Application.DoEvents();
                                                            }
                                                            else
                                                            {
                                                                // WritePackLog(con, nNewPackID,
                                                                WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (InsertResponseIntTable(con, nID, txtCommonRespText, dtDatOtv, nStatus, org, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                            {
                                                                iCnt++;
                                                                // WritePackLog(con, nNewPackID, "Обработан ответ # " + nID.ToString() + "\n");
                                                                WriteLLog(conGIBDD, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");
                                                                prbWritingDBF.PerformStep();
                                                                prbWritingDBF.Refresh();
                                                                System.Windows.Forms.Application.DoEvents();
                                                            }
                                                            else
                                                            {
                                                                // WritePackLog(con, nNewPackID, "Ошибка! Ответ # " + nID.ToString() + " обработать не удалось.\n");
                                                                WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                            }
                                                        }
                                                        // приводим строчку в исходное состояние
                                                        txtCommonRespText = "";
                                                    }

                                                    // теперь надо помыслить чего -делать-то еще надо :D - надо же какой тупой комментарий

                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                                                    if (nNewPackID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "Ошибка! Загрузка пакета ответов экстренно прервалась.\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Значение счетчика = " + iCnt.ToString() + "\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");

                                                        if (nID > 0)
                                                        {
                                                            WriteLLog(conGIBDD, nNewPackID, "ID запроса = " + nID.ToString() + "\n");
                                                        }
                                                    }

                                                }
                                            }
                                            else
                                            {
                                                // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                if (nNewPackID > 0)
                                                {
                                                    WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос ID = " + nID.ToString() + " не удалось загрузить т.к. не обнаружен запрос-родитель.\n");
                                                }
                                            }

                                        }

                                        //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        //WritePackLog(con, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                        //WritePackLog(con, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                        // установить количество обработанных запросов
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);


                                        // обновить статус лога-ответа
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - загружен ответ

                                        // обновить флаг что обработан FIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_FIND");

                                        //TODO: после этого надо еще и флаг обновить - что загружен

                                        // закомментировать т.к. теперь никакого пакета не создается   
                                        //// если все ок, то нужно поменять статус пакета
                                        //if(nNewPackID >0){
                                        //    SetDocumentStatus(nNewPackID, 70);
                                        //}
                                    }
                                }
                                else
                                {

                                    if (nAgreementID == 0)
                                    {
                                        nAgreementID = GetAgr_by_Org(org); // номер соглашения
                                        nAgent_id = GetAgent_ID(nAgreementID);
                                        nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                    }


                                    txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                    txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                    txtEntityName = GetLegal_Name(org);

                                    // создать лог отрицательных ответов и записать туда что 0 в пакете ответов
                                    frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD, "FL_FIND");
                                    nParentID = LogList.ShowForm();

                                    // если не было выбрано пропустить загрузку ответа
                                    if (nParentID != -1)
                                    {
                                        // 1 - Новый
                                        // 3 - Ответ положительный
                                        nNewPackID = CreateLLog(conGIBDD, 1, 3, txtAgreementCode, nParentID, "Пакет ответов из " + txtEntityName + ".");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");
                                        WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                        WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                        WriteLLog(conGIBDD, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                        // установить количество обработанных запросов
                                        UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                        // обновить статус лога-ответа
                                        UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                        // обновить статус лога-родителя
                                        // сразу меняем статус, т.к. фильтр все равно по флагам
                                        UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - загружен ответ
                                        // обновить флаг что обработан FIND
                                        UpdateLLogFlag(conGIBDD, nNewPackID, 1, "FL_FIND");
                                    }


                                }

                                //tran.Commit();
                                //con.Close();
                                MessageBox.Show("Обработано ответов: " + iCnt.ToString() + ".\nСейчас будет сформирован реестр ответов.", "Сообщение", MessageBoxButtons.OK);

                                //**********Формирование**реестра**find************
                                //Надо вспомнить запись в Ворд + подсчитывая количество строк 
                                //для разделения по приставам. 

                                //список всех приставов
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
                                // пишу обработчик для HTML
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
                                    report.AddToReport("Реестр ответов на запросы СП-И о наличии ден. средств из банка<br>");
                                    report.AddToReport("" + GetLegal_Name(org) + " от " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATOTV"]).ToShortDateString() + "<br>");
                                    report.AddToReport("За период с " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATZPR1"]).ToShortDateString() + " по " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATZPR2"]).ToShortDateString() + "<br>");

                                    spi = Convert.ToInt32(drspi["NOMSPI"]);

                                    report.AddToReport("СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "<br>");
                                    report.AddToReport("</h3>");
                                    
                                                                        
                                    

                                    foreach (DataRow row in tbl.Rows)
                                    {
                                        report.AddToReport("<p>");
                                        if (spi == Convert.ToInt32(row["NOMSPI"]))
                                        {
                                            string txtResponse = "";
                                            if ((row.Table.Columns.Contains("NOMLS")) && (row.Table.Columns.Contains("OSTAT")) && (Convert.ToString(row["NOMLS"]).TrimEnd() != ""))
                                            {
                                                //txtResponse += "л/с: " + Convert.ToString(row["NOMLS"]).TrimEnd() + " остаток = " + Money_ToStr(Convert.ToDecimal(row["OSTAT"])).TrimEnd();
                                                string txtLs = Convert.ToString(row["NOMLS"]).TrimEnd();
                                                txtResponse += "л/с: " + txtLs + " остаток = " + Convert.ToDecimal(row["OSTAT"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtLs);
                                            }

                                            if ((row.Table.Columns.Contains("RSCHET")) && (row.Table.Columns.Contains("OSTSCH")) && (Convert.ToString(row["RSCHET"]).TrimEnd() != ""))
                                            {
                                                //txtResponse += "; р/с: " + Convert.ToString(row["RSCHET"]).TrimEnd() + " остаток = " + Money_ToStr(Convert.ToDecimal(row["OSTSCH"])).TrimEnd();
                                                string txtRs = Convert.ToString(row["RSCHET"]).TrimEnd();
                                                txtResponse += "р/с: " + txtRs + " остаток = " + Convert.ToDecimal(row["OSTSCH"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtRs);
                                            }

                                            string txtResLine = GetIPNum(con, Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIO"]).TrimEnd();
                                            if (Convert.ToInt32(row["LITZDOLG"]) == 2)
                                            {

                                                txtResLine += " (" + Convert.ToInt32(row["GODR"]).ToString() + " г.р.)";
                                                //txtResLine += " (" + Convert.ToString(row["GODR"]).TrimEnd('0').TrimEnd(',') + " г.р.)";
                                                //txtResLine += " (" + Convert.ToInt32(row["GODR"]).ToString() + " г.р.)";
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

                                //    // для Ворда
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

                                //    //    par.Range.Text += "Реестр ответов на запросы СП-И о наличии денежных средств из банка\n";
                                //    //    par.Range.Text += GetLegal_Name(org) + " от " + Convert.ToDateTime(tbl.Rows[0]["DATOTV"]).ToShortDateString() + "\n";
                                //    //    par.Range.Text += "За период с " + Convert.ToDateTime(tbl.Rows[0]["DATZPR1"]).ToShortDateString() + " по " + Convert.ToDateTime(tbl.Rows[0]["DATZPR2"]).ToShortDateString() + "\n";

                                //    //    spi = Convert.ToInt32(drspi["NOMSPI"]);

                                //    //    sch_line = 0;
                                //    //    if (fl_fst == 1)
                                //    //    {
                                //    //        sch_line = 1;
                                //    //        fl_fst = 0;
                                //    //    }
                                //    //    par.Range.Text += PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";

                                //    //    //par.Range.Text += "НОМЕР ИП             ДОЛЖНИК                            АДРЕС                      СЧЕТ       ОСТАТОК\n";
                                //    //    par.Range.Text += "НОМЕР ИП             ДОЛЖНИК      ГОД РОЖДЕНИЯ          АДРЕС                      СЧЕТ       ОСТАТОК\n";

                                //    //    sch_line += 10;

                                //    //    foreach (DataRow row in tbl.Rows)
                                //    //    {
                                //    //        if (spi == Convert.ToInt32(row["NOMSPI"]))
                                //    //        {
                                //    //            string txtResponse = "";
                                //    //            if ((row.Table.Columns.Contains("NOMLS")) && (row.Table.Columns.Contains("OSTAT")) && (Convert.ToString(row["NOMLS"]).TrimEnd() != ""))
                                //    //            {
                                //    //                //txtResponse += "л/с: " + Convert.ToString(row["NOMLS"]).TrimEnd() + " остаток = " + Money_ToStr(Convert.ToDecimal(row["OSTAT"])).TrimEnd();
                                //    //                string txtLs = Convert.ToString(row["NOMLS"]).TrimEnd();
                                //    //                txtResponse += "л/с: " + txtLs + " остаток = " + Convert.ToDecimal(row["OSTAT"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtLs);
                                //    //            }

                                //    //            if ((row.Table.Columns.Contains("RSCHET")) && (row.Table.Columns.Contains("OSTSCH")) && (Convert.ToString(row["RSCHET"]).TrimEnd() != ""))
                                //    //            {
                                //    //                //txtResponse += "; р/с: " + Convert.ToString(row["RSCHET"]).TrimEnd() + " остаток = " + Money_ToStr(Convert.ToDecimal(row["OSTSCH"])).TrimEnd();
                                //    //                string txtRs = Convert.ToString(row["RSCHET"]).TrimEnd();
                                //    //                txtResponse += "р/с: " + txtRs + " остаток = " + Convert.ToDecimal(row["OSTSCH"]).ToString("F2").Replace(',', '.') + " " + getValuteByCod(txtRs);
                                //    //            }

                                //    //            string txtResLine = GetIPNum(Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIO"]).TrimEnd();
                                //    //            if (Convert.ToInt32(row["LITZDOLG"]) == 2)
                                //    //            {
                                //    //                txtResLine += " (" + Convert.ToInt32(row["GODR"]).ToString() + " г.р.)";
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
                                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                                }
                                if (nNewPackID > 0)
                                {
                                    WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответы не удалось загрузить без ошибок.\n");
                                    // обновить статус лога-ответа
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ответ загружен с ошибкой
                                }
                                //return false;
                            }
                            catch (Exception ex)
                            {
                                //if (DBFcon != null) DBFcon.Close();
                                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                                if (nNewPackID > 0)
                                {
                                    WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответы не удалось загрузить без ошибок.\n");
                                    // обновить статус лога-ответа
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ответ загружен с ошибкой
                                }
                                //return false;
                            }
                            //return true;
                        }
                        # endregion
                    }

                }
            }
            else MessageBox.Show("Ошибка приложения. Выберите организацию из списка.", "Внимание!", MessageBoxButtons.OK);

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
                        txtRes = "руб.";
                        break;

                    case "840":
                        txtRes = "долл.";
                        break;

                    case "978":
                        txtRes = "евро";
                        break;

                    case "826":
                        txtRes = "фунт стерл.";
                        break;

                    case "392":
                        txtRes = "яп. иена";
                        break;

                    case "756":
                        txtRes = "швейц. франк";
                        break;

                    default:
                        txtRes = "валюта с кодом " + txtCod;
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

            // Все закомментировал - поскольку UPDATE много времени занимает
            //string txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 23 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 3 and req_type.sndl_contr_id = " + ktfoms_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// для автоматической отправки
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 11 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 2 and req_type.sndl_contr_id = " + ktfoms_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            // автоматом больше не выгребаем
            DT_ktfoms_reg = null;

            // DT_ktfoms_doc = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 11 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_NSI_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            // DT_ktfoms_doc = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 11 and zapr_d.docstatusid = 2 and ip_d.docstatusid = 9 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            
            // вариант для 14 релиза
            // ключевой параметр pack.agreement_id = 30 - номер соглашения
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

            // поменять имя таблицы на osp_xxx.dbf
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
            int sch_line = 1; // счетчик строк, как минимум одна, даже если она пустая
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
                sch_line++; // если был перевод строки из-за длины
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

            
            openFileDialog1.Filter = "DBF файлы(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();
            int iRewriteState = 1; // обычный режим перезаписи ответов на запрос (запрашивать действия у пользователя)
            bool bNotIntTablesResp = false; // если ответ на запрос, сделанный без интерфейсных таблиц.
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
                            decimal nDx_pack_id = 0; //  интересно - зачем он нужен
                            string txtAgreementCode = "";
                            string txtAgentCode = "";
                            string txtAgentDeptCode = "";


                            // если файл с ответами не пустой, то по первому ответу определить
                            // этb ответы на запросы из 14 релиза или старые (без пакетов)
                            if (tbl.Rows.Count > 0)
                            {
                                decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["ZAPROS"]);
                                if (FindSendlist(nFirstID, ktfoms_id))
                                {
                                    // значит это новый запрос
                                    // получить параметры: соглашение, контрагент, подразделение
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

                                // если соглашение не выбрано, то устанавливаем все вместе
                                if (nAgreementID == 0)
                                {
                                    //GetAgr_by_Org - тоже неплохой вариант
                                    nAgreementID = 120;
                                    nAgent_id = GetAgent_ID(nAgreementID);
                                    nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                }

                                txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                string txtEntityName = GetLegal_Name(ktfoms_id);

                                // нужно создать новый входящий пакет
                                //nNewPackID = ID_CreateDX_PACK_I(con, 1, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);

                                // TODO: показать форму выбора запроса, к которому крепим ответ
                                frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD);
                                nParentID = LogList.ShowForm();

                                if (nParentID != -1)
                                {
                                    // 1 - Новый
                                    // 2 - Ответ простой
                                    nNewPackID = CreateLLog(conGIBDD, 1, 2, txtAgreementCode, nParentID, "Пакет ответов из " + txtEntityName + ".");

                                    // записать в лог пакета дату и начало обработки
                                    WritePackLog(con, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                    WritePackLog(con, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");

                                    foreach (DataRow row in tbl.Rows)
                                    {
                                        txtID = Convert.ToString(row["ZAPROS"]);
                                        if (!Decimal.TryParse(txtID, out nID))
                                        {
                                            nID = 0;
                                        }

                                        if (FindZapros(nID))
                                        {
                                            // значить начинаем вставлять в базу структуры данных ответа
                                            try
                                            {
                                                string txtDatZap = "";
                                                DateTime dtDatOtv, dtDatZap;


                                                // проверить дату запроса
                                                txtDatZap = Convert.ToString(row["DATZAPR"]);
                                                if (!DateTime.TryParse(txtDatZap, out dtDatZap))
                                                {
                                                    dtDatZap = DateTime.MaxValue;
                                                }

                                                bNotIntTablesResp = false; // теперь все ответы точно из интерфейсных таблиц
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

                                                // TODO: здесь нужно 
                                                txtResponse = "в соответствии с " + PKOSP_GetOrgConvention(ktfoms_id);
                                                txtResponse += " получен ответ: ";

                                                if (priz.ToUpper().Equals("T"))
                                                {
                                                    txtResponse += "Полное наименование места работы: " + Convert.ToString(row["NAMELONG"]).TrimEnd() + ".\n";
                                                    txtResponse += "ФИО руководителя: " + Convert.ToString(row["FIO_BOSS"]).TrimEnd() + ".\n";
                                                    txtResponse += "Телефон руководителя: " + Convert.ToString(row["TEL_BOSS"]).TrimEnd() + ".\n";
                                                    txtResponse += "Адрес работодателя: " + Convert.ToString(row["ADR_PR"]).TrimEnd() + ".\n";
                                                    txtResponse += "Адрес должника: " + Convert.ToString(row["ADRES"]).TrimEnd() + ".\n";
                                                    txtResponse += "Тип договора страхования: " + Convert.ToString(row["TYPE_DOG"]).TrimEnd() + ".\n";
                                                    txtResponse += "Номер договора страхования: " + Convert.ToString(row["N_DOG"]).TrimEnd() + ".\n";
                                                    nStatus = 20;

                                                }
                                                else
                                                {
                                                    //txtResponse = "нет данных о должнике при запросе за период с " + dat1.ToShortDateString() + " по " + dat2.ToShortDateString();
                                                    txtResponse += "нет данных о должнике по запросу от " + Convert.ToDateTime(row["DATZAPR"]).ToShortDateString();
                                                    nStatus = 7;
                                                }

                                                txtOtvet = txtResponse;

                                                // добавить в качестве параметра ID контрагента
                                                if (bNotIntTablesResp)
                                                {
                                                    if (InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, ktfoms_id, ref iRewriteState, nNewPackID))
                                                    {
                                                        iCnt++;
                                                        WritePackLog(con, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");
                                                        prbWritingDBF.PerformStep();
                                                        prbWritingDBF.Refresh();
                                                        System.Windows.Forms.Application.DoEvents();
                                                    }
                                                    else
                                                    {
                                                        // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                        WritePackLog(con, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                    }

                                                }
                                                else
                                                {
                                                    if (InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, ktfoms_id, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                    {
                                                        iCnt++;
                                                        //WritePackLog(con, nNewPackID, "Обработан ответ # " + nID.ToString() + "\n");
                                                        WritePackLog(con, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");
                                                        prbWritingDBF.PerformStep();
                                                        prbWritingDBF.Refresh();
                                                        System.Windows.Forms.Application.DoEvents();
                                                    }
                                                    else
                                                    {
                                                        // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                        //WritePackLog(con, nNewPackID, "Ошибка! Ответ # " + nID.ToString() + " обработать не удалось.\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                                                if (nNewPackID > 0)
                                                {
                                                    WriteLLog(conGIBDD, nNewPackID, "Ошибка! Загрузка пакета ответов экстренно прервалась.\n");
                                                    WriteLLog(conGIBDD, nNewPackID, "Значение счетчика = " + iCnt.ToString() + "\n");
                                                    WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");
                                                    if (nID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "ID запроса = " + nID.ToString() + "\n");
                                                    }
                                                }

                                            }
                                        }
                                        else
                                        {
                                            // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                            // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                            if (nNewPackID > 0)
                                            {
                                                WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос № " + nID.ToString() + " не удалось загрузить т.к. не обнаружен запрос-родитель.\n");
                                            }
                                        }

                                    }
                                    //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                    //WritePackLog(con, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                    //WritePackLog(con, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");
                                    WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                    WritePackLog(con, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                    WritePackLog(con, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                    // установить количество обработанных запросов
                                    UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                    // обновить статус лога-ответа
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                    // обновить статус лога-родителя
                                    UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - загружен ответ

                                    // нет больше пакета ПК ОСП и статуса ему менять не надо
                                    //// если все ок, то нужно поменять статус пакета
                                    //if (nNewPackID > 0)
                                    //{
                                    //    SetDocumentStatus(nNewPackID, 70);
                                    //}
                                }
                            }

                            MessageBox.Show("Обработано ответов: " + iCnt.ToString() + ".\n Сейчас будет сформирован реестр ответов.", "Сообщение", MessageBoxButtons.OK);   
                            

                            //**********тНПЛХПНБЮМХЕ**ПЕЕЯРПЮ**ktfoms************
                            //мЮДН БЯОНЛМХРЭ ГЮОХЯЭ Б бНПД + ОНДЯВХРШБЮЪ ЙНКХВЕЯРБН ЯРПНЙ 
                            //ДКЪ ПЮГДЕКЕМХЪ ОН ОПХЯРЮБЮЛ. 

                            //ЯОХЯНЙ БЯЕУ ОПХЯРЮБНБ
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
                            // пишу обработчик для HTML
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
                                report.AddToReport( "Реестр ответов на запросы СП-И в КТФОМС<br>");
                                report.AddToReport("Ответ из КТФОМС от " + Convert.ToDateTime(tbl.Rows[0]["DATZAPR"]).ToShortDateString() + "<br>");
                                spi = Convert.ToInt32(drspi["NOMSPI"]);
                                report.AddToReport("СП-И: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "<br>");
                                                                
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
                                                txtResponse = "нет данных";
                                            else
                                                txtResponse = Convert.ToString(row["TYPE_DOG"]).TrimEnd() + ", " + Convert.ToString(row["NAMELONG"]).TrimEnd();

                                            if ((Convert.ToString(row["FIO_BOSS"]).Trim()) != "")
                                                txtResponse += ", ФИО руководителя: " + Convert.ToString(row["FIO_BOSS"]).TrimEnd();

                                            if ((Convert.ToString(row["TEL_BOSS"]).Trim()) != "")
                                                txtResponse += ", телефон руководителя: " + Convert.ToString(row["TEL_BOSS"]).TrimEnd();

                                            if ((Convert.ToString(row["ADR_PR"]).Trim()) != "")
                                                txtResponse += ", адрес работодателя: " + Convert.ToString(row["ADR_PR"]).TrimEnd();

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

                            //        par.Range.Text += "Реестр ответов на запросы СП-И в КТФОМС";
                            //        //par.Range.Text += "КТФОМС от " + DateTime.Today.ToShortDateString() + "\n";
                            //        par.Range.Text += "Ответ из КТФОМС от " + Convert.ToDateTime(tbl.Rows[0]["DATZAPR"]).ToShortDateString() + "\n";

                            //        // убрал т.к. dat1 и dat2 никакого отношения к реальным датам запроса и ответа не имеют
                            //        // par.Range.Text += "За период с " + dat1.ToShortDateString() + " по " + dat2.ToShortDateString() + "\n";

                            //        spi = Convert.ToInt32(drspi["NOMSPI"]);
                            //        par.Range.Text += "СП-И: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";
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
                            //                        txtResponse = "нет данных";
                            //                    else
                            //                        txtResponse = Convert.ToString(row["TYPE_DOG"]).TrimEnd() + ", " + Convert.ToString(row["NAMELONG"]).TrimEnd();

                            //                    if ((Convert.ToString(row["FIO_BOSS"]).Trim()) != "")
                            //                        txtResponse += ", ФИО руководителя: " + Convert.ToString(row["FIO_BOSS"]).TrimEnd();

                            //                    if ((Convert.ToString(row["TEL_BOSS"]).Trim()) != "")
                            //                        txtResponse += ", телефон руководителя: " + Convert.ToString(row["TEL_BOSS"]).TrimEnd();

                            //                    if ((Convert.ToString(row["ADR_PR"]).Trim()) != "")
                            //                        txtResponse += ", адрес работодателя: " + Convert.ToString(row["ADR_PR"]).TrimEnd();

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
                                MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                            }
                            if (nNewPackID > 0)
                            {
                                WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответы не удалось загрузить без ошибок.\n");
                                // обновить статус лога-ответа
                                UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ответ загружен с ошибкой
                            }
                            //return false;
                        }
                        catch (Exception ex)
                        {
                            //if (DBFcon != null) DBFcon.Close();
                            MessageBox.Show("Ошибка при работе с данными. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                            WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответы не удалось загрузить без ошибок.\n");
                            // обновить статус лога-ответа
                            UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ответ загружен с ошибкой
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


            openFileDialog1.Filter = "DBF файлы(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();
            int iRewriteState = 1; // обычный режим перезаписи ответов на запрос (запрашивать действия у пользователя)
            DataTable tbl = null;
            decimal nStatus = 0;
            DataSet ds = null;
            bool bVFP_DBASE_local = false;
            OleDbConnection DbaseCon;
            bool bEx = false;
            string txtFileDir;
            
            bool bNotIntTablesResp = false; // если ответ на запрос, сделанный без интерфейсных таблиц.

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
                                MessageBox.Show("Ошибка при работе с данными. Будет предпринята повторная попытка обработать файл. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                            }
                            bVFP_DBASE_local = true; // попробовать обработать через DBASE
                        }

                        if (bVFP_DBASE_local)
                            {
                                try
                                {
                                    // если имя файла больше 8 символов - то копировать и обработат меньшее
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
                                        MessageBox.Show("Ошибка при работе с данными. Файл обработать не удалось. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                                    }
                                    bEx = true;
                                }
                                bVFP_DBASE_local = false;
                            }

                        # region "ОБРАБОТКА ФАЙЛА"
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

                            // если файл с ответами не пустой, то по первому ответу определить
                            // этb ответы на запросы из 14 релиза или старые (без пакетов)
                            if (tbl.Rows.Count > 0)
                            {
                                // найти был ли запрос
                                decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["NOMZAP"]);
                                if (FindSendlist(nFirstID, pens_id)) // указываем параметр org - контрагент из списка рассылки, которому была направлена копия
                                {
                                    // значит это новый запрос
                                    // получить параметры: соглашение, контрагент, подразделение
                                    // TODO: этот кусок убрать - пареметры получать по Agreement_Code или Agreement_ID
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

                                // это основной вариант получения информации - по Agreement_ID
                                if (nAgreementID == 0)
                                {
                                    //GetAgr_by_Org - тоже неплохой вариант
                                    nAgreementID = 100;
                                    nAgent_id = GetAgent_ID(nAgreementID);
                                    nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                }


                                txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                string txtEntityName = GetLegal_Name(pens_id);

                                // Больше не нужно создавать пакет - просто создаем свой лог
                                // нужно создать новый входящий пакет
                                //nNewPackID = ID_CreateDX_PACK_I(con, 1, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);

                                // TODO: показать форму выбора запроса, к которому крепим ответ
                                // форма - таблица с возможностью поиска по дате
                                // входные параметры - txtAgreementCode
                                frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD);
                                nParentID = LogList.ShowForm();


                                if (nParentID != -1)
                                {
                                    // 1 - Новый
                                    // 2 - Ответ простой
                                    nNewPackID = CreateLLog(conGIBDD, 1, 2, txtAgreementCode, nParentID, "Пакет ответов из " + txtEntityName + ".");

                                    // записать в лог пакета дату и начало обработки

                                    //WritePackLog(con, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                    //WritePackLog(con, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");

                                    WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                    WriteLLog(conGIBDD, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");


                                    foreach (DataRow row in tbl.Rows)
                                    {
                                        txtID = Convert.ToString(row["NOMZAP"]);
                                        if (!Decimal.TryParse(txtID, out nID))
                                        {
                                            nID = 0;
                                        }
                                        // если был запрос
                                        if (FindZapros(nID))
                                        {
                                            // значить начинаем вставлять в базу структуры данных ответа
                                            try
                                            {

                                                string txtDatZap = "";
                                                DateTime dtDatOtv;



                                                // нет DATOTV в файле, будем просто текущую дату писать
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

                                                txtResponse = "в соответствии с " + PKOSP_GetOrgConvention(pens_id);
                                                txtResponse += " получен ответ: ";

                                                if (priz == 1)
                                                {
                                                    txtResponse += "Должник является получателем пенсии.\n";
                                                    txtResponse += "Адрес: " + Convert.ToString(row["ADRES"]).TrimEnd() + "\n";
                                                    txtResponse += "Cумма дохода, на которую можно обратить взыскание: " + Convert.ToString(row["SUMMA"]).TrimEnd() + ". " + Convert.ToString(row["KOMMENT"]).TrimEnd() + "\n";
                                                    nStatus = 20; // ответ получен

                                                }
                                                else
                                                {
                                                    nStatus = 7; // нет данных
                                                    if (priz == 0)
                                                    {
                                                        txtResponse += "нет данных о должнике по запросу от " + dtDatZap.ToShortDateString();
                                                    }
                                                    else
                                                    {
                                                        txtResponse += "нет данных о должнике по запросу от " + dtDatZap.ToShortDateString() + " " + Convert.ToString(row["SUMMA"]).TrimEnd();
                                                    }
                                                }
                                                txtOtvet = txtResponse;

                                                // TODO: тут надо написать новую функцию - через интерфейсные таблицы =)
                                                // вытащить параметры - пакет, agent_agreement, agent_dept_code, agent_code, enity_name

                                                bNotIntTablesResp = false; // теперь все ответы точно из интерфейсных таблиц
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
                                                        WritePackLog(con, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");
                                                        prbWritingDBF.PerformStep();
                                                        prbWritingDBF.Refresh();
                                                        System.Windows.Forms.Application.DoEvents();
                                                    }
                                                    else
                                                    {
                                                        // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                        WritePackLog(con, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                        nStatus = 15; // Ошибка
                                                    }
                                                }
                                                else
                                                {

                                                    if (InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, pens_id, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                    {
                                                        iCnt++;
                                                        //WritePackLog(con, nNewPackID, "Обработан ответ # " + nID.ToString() + "\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");
                                                        prbWritingDBF.PerformStep();
                                                        prbWritingDBF.Refresh();
                                                        System.Windows.Forms.Application.DoEvents();
                                                    }
                                                    else
                                                    {
                                                        // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                        // WritePackLog(con, nNewPackID, "Ошибка! Ответ # " + nID.ToString() + " обработать не удалось.\n");
                                                        WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                        nStatus = 15; // Ошибка
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                                                if (nNewPackID > 0)
                                                {
                                                    WriteLLog(conGIBDD, nNewPackID, "Ошибка! Загрузка пакета ответов экстренно прервалась.\n");
                                                    WriteLLog(conGIBDD, nNewPackID, "Значение счетчика = " + iCnt.ToString() + "\n");
                                                    WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");
                                                    if (nID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "ID запроса = " + nID.ToString() + "\n");
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                            if (nNewPackID > 0)
                                            {
                                                WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос ID = " + nID.ToString() + " не удалось загрузить т.к. не обнаружен запрос-родитель.\n");
                                            }
                                        }
                                    }

                                    //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                    //WritePackLog(con, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                    //WritePackLog(con, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                    WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                    WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                    WriteLLog(conGIBDD, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                    // установить количество обработанных запросов
                                    UpdateLLogCount(conGIBDD, nNewPackID, iCnt);


                                    // обновить статус лога-ответа
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                    // обновить статус лога-родителя
                                    // для ОСП в ПТЗ - ответа будет 2, поэтому:
                                    //  - если текущий статус = 2 (Обработан), то установить статус 12 (Обработана часть ответов)
                                    //  - если текущий статус = 12, то установить статус 10 (Загружен ответ)
                                    
                                    // если пакет грузим без привязки к родителю - то ничего делать не надо
                                    if (nParentID > 0)
                                    {
                                        decimal nOldStatus = GetLLogStatus(conGIBDD, nParentID);
                                        if (nOldStatus == 2) UpdateLLogParentStatus(conGIBDD, nNewPackID, 12); // 12 (Обработана часть ответов)
                                        else if (nOldStatus == 12) UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - загружен ответ
                                    }

                                    // для всех остальных ОСП будет только это:
                                    // UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - загружен ответ

                                    // убрал т.к. теперь никакого пакета не создается
                                    //// если все ок, то нужно поменять статус пакета
                                    //if (nNewPackID > 0)
                                    //{
                                    //    SetDocumentStatus(nNewPackID, 70);
                                    //}
                                }
                            }

                            MessageBox.Show("Обработано ответов: " + iCnt.ToString() + ".\n Сейчас будет сформирован реестр ответов.", "Сообщение", MessageBoxButtons.OK);

                            //**********Формирование**реестра**pens************
                            //Надо вспомнить запись в Ворд + подсчитывая количество строк 
                            //для разделения по приставам. 

                            //список всех приставов с положительными ответами выжать из таблицы tbl
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
                            // пишу обработчик для HTML
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
                                report.AddToReport("Реестр ответов на запросы СП-И о персонифицированных данных в ПФР<br />");
                                report.AddToReport("Ответы из ПФР от " + DateTime.Today.ToShortDateString() + "<br />");
                                //report.AddToReport("За период с " + dat1.ToShortDateString() + " по " + dat2.ToShortDateString() + "<br />");
                                spi = Convert.ToInt32(drspi["NOMSPI"]);
                                report.AddToReport("СП-И: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "<br />");
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
                                            report.AddToReport( "Должник является получателем пенсии. Cумма дохода, на которую можно обратить взыскание: " + Convert.ToString(row["SUMMA"]).TrimEnd() + "<br>");
                                            fl_no_answer = false;
                                        }
                                       
                                            
                                    }
                                    report.AddToReport("</p>");
                                }
                                
                                // если ничего положительного в ответах нет, то так и пишем
                                if (fl_no_answer)
                                {
                                    report.AddToReport( "Нет положительных ответов по запросам о наличии пенсии у должников.");                       
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
                            //    //      пример для Ворда

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

                            //        par.Range.Text += "Реестр ответов по запросам СП-И в ПФР о наличии пенсии\n";
                            //        par.Range.Text += "На запрос в ПФР от " + DateTime.Today.ToShortDateString() + "\n";
                            //        //par.Range.Text += "За период с " + dat1.ToShortDateString() + " по " + dat2.ToShortDateString() + "\n";

                            //        spi = Convert.ToInt32(drspi["NOMSPI"]);

                            //        par.Range.Text += "СП-И: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";

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
                            //                    par.Range.Text += "Должник является получателем пенсии. Cумма дохода, на которую можно обратить взыскание: " + Convert.ToString(row["SUMMA"]).TrimEnd() + "\n";
                            //                    sch_line += 5;
                            //                }

                            //                //if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                            //                //string priz = Convert.ToString(row["PRIZ"]).TrimEnd();
                            //                //if (priz.ToUpper().Equals("T"))
                            //            }
                            //        }
                            //        // если ничего положительного в ответах нет, то так и пишем
                            //        if (sch_line == 6)
                            //        {
                            //            par.Range.Text += "Нет положительных ответов по запросам о наличии пенсии у должников.";
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
                                MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                            }

                            if (nNewPackID > 0)
                            {
                                WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответы не удалось загрузить без ошибок.\n");
                                // обновить статус лога-ответа
                                UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ответ загружен с ошибкой
                            }

                            //return false;
                        }
                        catch (Exception ex)
                        {
                            //if (DBFcon != null) DBFcon.Close();
                            MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);

                            if (nNewPackID > 0)
                            {
                                WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответы не удалось загрузить без ошибок.\n");
                                // обновить статус лога-ответа
                                UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ответ загружен с ошибкой
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
            
            // Принудительно готовим к отправке все пакеты - которые на статусе НОВЫЙ. Для каждого regl.sending_mode свой статус

            //!!! ВСЕ ЗАКОММЕНТИРОВАЛ - чтобы не было проблем со временем работы - UPDATE очень долго делается

            ///// Внимание - для работы с интерфейсными таблицами отправка сторонней программой
            ///// не позволяет запуститься выгрузке сведений о запросе в интерфейсную таблицу, поэтому надо убирать этот вариант
            //// d.docstatusid = 23 это статус отправка сторонней программой 
            //string txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 23 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 3 and req_type.sndl_contr_id = " + pens_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// для автоматической отправки
            ////  d.docstatusid = 11 - это статус автоматическая отправка
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 11 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 2 and req_type.sndl_contr_id = " + pens_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// есть еще статус 21 - ручная отправка
            //// для автоматической отправки
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 21 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 1 and req_type.sndl_contr_id = " + pens_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            DT_pens_reg = null;
            //DT_pens_doc = GetDataTableFromFB("SELECT DISTINCT c.PRIMARY_SITE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD, UPPER(a.name_d) as FIOVK, a.DATE_BORN_D as DATROZHD, a.SUM_ as SUMVZ, a.ADR_D as ADDR, a.ADR_D as ADDR, a.PK as FK_IP, a.PK_ID as FK_ID, a.uscode, a.FIO_SPI FROM IP a left join s_users c on (a.uscode=c.uscode) LEFT JOIN DOCUMENT b ON b.FK = a.PK  WHERE a.SISP_KEY = '/1/3/' and a.DATE_IP_OUT is null and b.KOD != 1006 and (a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' and a.ssd is null and a.ssv is null", "TOFIND");
            //DT_pens_doc = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 15 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_NSI_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            //DT_pens_doc = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 15 and zapr_d.docstatusid = 2 and ip_d.docstatusid = 9 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            
            //// выборка пакетов напрямую из БД
            //DT_pens_doc = GetDataTableFromFB("select pack.id as pack_id, 2 LITZDOLG, d_req.id ZAPROS, req.IPNO_NUM, req.div, req.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, d_req.doc_date DATZAPR,  req.ID_DBTR_ADR ADDR, req.ID_DBTR_BORN DATROZHD, req.ID_DBTRCLS,  req.DBTR_BORN_YEAR GOD, req.ID_DEBTSUM SUMMA, req.ID_DBTR_INN INNORG,  d_req.doc_number, req.ID_DEBTCLS_NAME VIDVZISK from dx_pack_o packo left join dx_pack pack on pack.id = packo.id join sendlist sl on pack.id = sl.dx_pack_id join o_ip req on sl.sendlist_o_id = req.id join document d_req on req.id = d_req.id join document ip_d on d_req.parent_id = ip_d.id join document dpack on pack.id = dpack.id join SPI on req.IP_EXEC_PRIST = spi.SUSER_ID where dpack.docstatusid = 23  and pack.agreement_id = 100 and packo.has_been_sent is null  and d_req.docstatusid != 19 and d_req.docstatusid != 15  and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");

            // выборка пакетов из Интерфейсных Таблиц (ИТ)
            // ext_request_id - нужен чтобы потом быстро сделать update
            DT_pens_doc = GetDataTableFromFB("select 100 agreement_id, ext_request_id,  pack_id,  2 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID where mvv_agreement_code = 100 and processed = 0 and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            
            // 100 - номер соглашения с ПФ

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
                MessageBox.Show("Exception Thrown: " + e.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("Ошибка приложения. Файл не найден! Message: " + exc.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка приложения. Файл не найден! Message: " + exc.ToString(), "Внимание!", MessageBoxButtons.OK);
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

            openFileDialog1.Filter = "DBF файлы(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();
            decimal nStatus = 0;
            int iRewriteState = 1; // обычный режим перезаписи ответов на запрос (запрашивать действия у пользователя)
            bool bVFP_DBASE_local = false;
            DataTable tbl = null;
            DataSet ds = null;
            OleDbConnection DbaseCon;
            bool bEx = false;
            string txtFileDir;
            bool bNotIntTablesResp = false; // если ответ на запрос, сделанный без интерфейсных таблиц.

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
                                    MessageBox.Show("Ошибка при работе с данными. Будет предпринята повторная попытка обработать файл. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
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
                                    // если имя файла больше 8 символов - то копировать и обработат меньшее
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
                                        MessageBox.Show("Ошибка при работе с данными. Файл повторно обработать не удалось. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                                    }
                                    bEx = true;
                                }
                                bVFP_DBASE_local = false;
                            }


                        
                    
                        # region "ОБРАБОТКА ФАЙЛА"
                        // если файл открылся без Exception
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


                                    // если файл с ответами не пустой, то по первому ответу определить
                                    // этb ответы на запросы из 14 релиза или старые (без пакетов)
                                    if (tbl.Rows.Count > 0)
                                    {
                                        decimal nFirstID = Convert.ToDecimal(tbl.Rows[0]["ZAPROS"]);
                                        if (FindSendlist(nFirstID, potd_id)) // указываем параметр org - контрагент из списка рассылки, которому была направлена копия
                                        {
                                            // значит это новый запрос
                                            // получить параметры: соглашение, контрагент, подразделение
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
                                            //GetAgr_by_Org - тоже неплохой вариант
                                            nAgreementID = 110;
                                            nAgent_id = GetAgent_ID(nAgreementID);
                                            nAgent_dept_id = GetAgentDept_ID(nAgreementID);
                                        }


                                        txtAgreementCode = GetAgreement_Code(Convert.ToInt32(nAgreementID));
                                        txtAgentCode = GetAgent_Code(Convert.ToInt32(nAgreementID));
                                        txtAgentDeptCode = GetAgentDept_Code(Convert.ToInt32(nAgreementID));

                                        string txtEntityName = GetLegal_Name(potd_id);

                                        // нужно создать новый входящий пакет
                                        // nNewPackID = ID_CreateDX_PACK_I(con, 1, nAgent_id, nAgent_dept_id, nAgreementID, "", txtAgentCode, txtAgreementCode, txtAgentDeptCode);

                                        // TODO: показать форму выбора запроса, к которому крепим ответ
                                        frmLogList LogList = new frmLogList(con, txtAgreementCode, constrGIBDD);
                                        nParentID = LogList.ShowForm();

                                        if (nParentID != -1)
                                        {
                                            // 1 - Новый
                                            // 2 - Ответ простой
                                            nNewPackID = CreateLLog(conGIBDD, 1, 2, txtAgreementCode, nParentID, "Пакет ответов из " + txtEntityName + ".");

                                            // записать в лог пакета дату и начало обработки
                                            //WritePackLog(con, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                            //WritePackLog(con, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");
                                            WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " начало обработки ответа.\n");
                                            WriteLLog(conGIBDD, nNewPackID, "Обрабатывается файл: " + openFileDialog1.FileName + "\n");


                                            foreach (DataRow row in tbl.Rows)
                                            {
                                                txtID = Convert.ToString(row["ZAPROS"]);
                                                if (!Decimal.TryParse(txtID, out nID))
                                                {
                                                    nID = 0;
                                                }
                                                if (FindZapros(nID))
                                                {
                                                    // значить начинаем вставлять в базу структуры данных ответа
                                                    try
                                                    {
                                                        string txtDatZap = "";
                                                        DateTime dtDatOtv, dtDatZap;


                                                        // проверить дату запроса
                                                        txtDatZap = Convert.ToString(row["DATZAP"]);
                                                        if (!DateTime.TryParse(txtDatZap, out dtDatZap))
                                                        {
                                                            dtDatZap = DateTime.MaxValue;
                                                        }

                                                        bNotIntTablesResp = false; // теперь все ответы точно из интерфейсных таблиц
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

                                                        txtResponse = "в соответствии с " + PKOSP_GetOrgConvention(potd_id);
                                                        txtResponse += " получен ответ: ";

                                                        if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                                                        {
                                                            txtResponse += "нет данных по должнику";
                                                            nStatus = 7; // нет данных
                                                        }
                                                        else
                                                        {
                                                            txtResponse += "Адрес: " + Convert.ToString(row["ADRES"]).TrimEnd() + ".\n";
                                                            txtResponse += "Наименование страхователя: " + Convert.ToString(row["NAMEORG"]).TrimEnd() + ".\n";
                                                            txtResponse += "Местонахождение страхователя: " + Convert.ToString(row["ADRORG"]).TrimEnd() + ".\n";
                                                            txtResponse += "Дата начала периода работы: " + Convert.ToString(row["DATST"]).TrimEnd() + ".\n";
                                                            txtResponse += "Дата окончания периода работы: " + Convert.ToString(row["DATFN"]).TrimEnd() + ".\n";
                                                            txtResponse += "Комментарий: " + Convert.ToString(row["KOMMENT"]).TrimEnd() + ".\n";
                                                            nStatus = 20; // получен ответ
                                                        }

                                                        txtOtvet = txtResponse;

                                                        if (bNotIntTablesResp)
                                                        {
                                                            if (InsertZaprosTo_PK_OSP(con, nID, txtOtvet, dtDatOtv, nStatus, potd_id, ref iRewriteState, nNewPackID))
                                                            {
                                                                iCnt++;
                                                                WritePackLog(con, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");
                                                                prbWritingDBF.PerformStep();
                                                                prbWritingDBF.Refresh();
                                                                System.Windows.Forms.Application.DoEvents();
                                                            }
                                                            else
                                                            {
                                                                // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                                nStatus = 15; // ошибка
                                                                WritePackLog(con, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (InsertResponseIntTable(con, nID, txtOtvet, dtDatOtv, nStatus, potd_id, ref iRewriteState, nNewPackID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName))
                                                            {
                                                                iCnt++;
                                                                //WritePackLog(con, nNewPackID, "Обработан ответ # " + nID.ToString() + "\n");
                                                                WriteLLog(conGIBDD, nNewPackID, "Обработан ответ на запрос # " + nID.ToString() + "\n");

                                                                prbWritingDBF.PerformStep();
                                                                prbWritingDBF.Refresh();
                                                                System.Windows.Forms.Application.DoEvents();
                                                            }
                                                            else
                                                            {
                                                                // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                                nStatus = 15; // ошибка
                                                                //WritePackLog(con, nNewPackID, "Ошибка! Ответ # " + nID.ToString() + " обработать не удалось.\n");
                                                                WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос # " + nID.ToString() + " обработать не удалось.\n");
                                                            }
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                                                        if (nNewPackID > 0)
                                                        {
                                                            WriteLLog(conGIBDD, nNewPackID, "Ошибка! Загрузка пакета ответов экстренно прервалась.\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "Значение счетчика = " + iCnt.ToString() + "\n");
                                                            WriteLLog(conGIBDD, nNewPackID, "Exception: " + ex.ToString() + "\n");
                                                            if (nID > 0)
                                                            {
                                                                WriteLLog(conGIBDD, nNewPackID, "ID запроса = " + nID.ToString() + "\n");
                                                            }
                                                        }

                                                    }
                                                }
                                                else
                                                {
                                                    // ответ не удалось загрузить, надо бы это как-то в реестре отметить
                                                    //WritePackLog(con, nNewPackID, "Ошибка! Ответ # " + nID.ToString() + " обработать не удалось.\n");
                                                    if (nNewPackID > 0)
                                                    {
                                                        WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответ на запрос ID = " + nID.ToString() + " не удалось загрузить т.к. не обнаружен запрос-родитель.\n");
                                                    }
                                                }

                                            }
                                            //WritePackLog(con, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                            //WritePackLog(con, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                            //WritePackLog(con, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");
                                            WriteLLog(conGIBDD, nNewPackID, "+++++++++++++++++++++++++++++++++++++++\n");
                                            WriteLLog(conGIBDD, nNewPackID, DateTime.Now + " завершение обработки ответа.\n");
                                            WriteLLog(conGIBDD, nNewPackID, "Обработано записей: " + iCnt.ToString() + "\n");

                                            // установить количество обработанных запросов
                                            UpdateLLogCount(conGIBDD, nNewPackID, iCnt);

                                            // обновить статус лога-ответа
                                            UpdateLLogStatus(conGIBDD, nNewPackID, 2);

                                            // обновить статус лога-родителя
                                            UpdateLLogParentStatus(conGIBDD, nNewPackID, 10); // 10 - загружен ответ

                                            //// если все ок, то нужно поменять статус пакета
                                            //if (nNewPackID > 0)
                                            //{
                                            //    SetDocumentStatus(nNewPackID, 70);
                                            //}
                                        }
                                    }

                                    MessageBox.Show("Обработано ответов: " + iCnt.ToString() + ".\n Сейчас будет сформирован реестр ответов.", "Сообщение", MessageBoxButtons.OK);

                                    //**********Формирование**реестра**pens************
                                    //Надо вспомнить запись в Ворд + подсчитывая количество строк 
                                    //для разделения по приставам. 

                                    //список всех приставов
                                    DataTable dtspi = ds.Tables.Add("SPI");

                                    // вытащить всех СПИ из таблицы ответов.
                                    // в базу данных не полезу - выжму из DataTable
                                    // получить список пакетов
                                    string[] cols = new string[] { "NOMSPI" };
                                    dtspi = SelectDistinct(tbl, cols);

                                    // пишу обработчик для HTML
                                    prbWritingDBF.Value = 0;
                                    prbWritingDBF.Maximum = dtspi.Rows.Count;
                                    prbWritingDBF.Step = 1;
                                    Int32 spi = 0;

                                    ReportMaker report = new ReportMaker();
                                    report.StartReport();
                                    foreach (DataRow drspi in dtspi.Rows)
                                    {
                                        report.AddToReport("<h3>");
                                        report.AddToReport("Реестр ответов на запросы СП-И о персонифицированных данных в ПФР<br />");
                                        report.AddToReport("Ответы из ПФР от " + DateTime.Today.ToShortDateString() + "<br />");
                                        // report.AddToReport("За период с " + dat1.ToShortDateString() + " по " + dat2.ToShortDateString() + "<br />");
                                        spi = Convert.ToInt32(drspi["NOMSPI"]);
                                        report.AddToReport("СП-И: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "<br />");
                                        report.AddToReport("</h3>");

                                        foreach (DataRow row in tbl.Rows)
                                        {
                                            if (spi == Convert.ToInt32(row["NOMSPI"]))
                                            {
                                                report.AddToReport("<br />");
                                                report.AddToReport(Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FNAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["NAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["SNAMEDOL"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd() + "<br />");
                                                if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                                                    report.AddToReport("нет данных по должнику<br />");
                                                else
                                                {
                                                    report.AddToReport("Наименование страхователя: " + Convert.ToString(row["NAMEORG"]).TrimEnd() + ".<br />");
                                                    report.AddToReport("Местонахождение страхователя: " + Convert.ToString(row["ADRORG"]).TrimEnd() + ".<br />");

                                                    try
                                                    {
                                                        report.AddToReport("Дата начала периода работы: " + Convert.ToDateTime(row["DATST"]).ToShortDateString() + ".<br />");
                                                        report.AddToReport("Дата окончания периода работы: " + Convert.ToDateTime(row["DATFN"]).ToShortDateString() + ".<br />");
                                                    }
                                                    catch
                                                    {
                                                        // это левый catch - на самом деле надо DateTime парсить нормально

                                                    }
                                                    //par.Range.Text += "Дата начала периода работы: " + Convert.ToDateTime(row["DATST"]).ToShortDateString() + ".";
                                                    //par.Range.Text += "Дата окончания периода работы: " + Convert.ToDateTime(row["DATFN"]).ToShortDateString() + ".";

                                                    report.AddToReport("Комментарий: " + Convert.ToString(row["KOMMENT"]).TrimEnd() + "<br />");
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
                                    //                                //      пример для Ворда

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

                                    //                                    par.Range.Text += "Реестр ответов на запросы СП-И о персонифицированных данных в ПФР\n";
                                    //                                    par.Range.Text += "Ответы из ПФР от " + DateTime.Today.ToShortDateString() + "\n";
                                    //                                    par.Range.Text += "За период с " + dat1.ToShortDateString() + " по " + dat2.ToShortDateString() + "\n";

                                    //                                    spi = Convert.ToInt32(drspi["NOMSPI"]);

                                    //                                    sch_line = 0;
                                    //                                    if (fl_fst == 1)
                                    //                                    {
                                    //                                        sch_line = 1;
                                    //                                        fl_fst = 0;
                                    //                                    }
                                    //                                    //par.Range.Text += " ";
                                    //                                    par.Range.Text += "СП-И: " + PK_OSP_GetSPI_Name(Convert.ToInt32(drspi["NOMSPI"])) + "\n";
                                    //                                    //par.Range.Text += GetOSP_Name();
                                    //                                    sch_line += 8;

                                    //                                    foreach (DataRow row in tbl.Rows)
                                    //                                    {
                                    //                                        if (spi == Convert.ToInt32(row["NOMSPI"]))
                                    //                                        {
                                    //                                            par.Range.Text += Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FNAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["NAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["SNAMEDOL"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd();
                                    //                                            if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                                    //                                                par.Range.Text += "нет данных по должнику\n";
                                    //                                            else
                                    //                                            {
                                    //                                                par.Range.Text += "Наименование страхователя: " + Convert.ToString(row["NAMEORG"]).TrimEnd() + ".";
                                    //                                                par.Range.Text += "Местонахождение страхователя: " + Convert.ToString(row["ADRORG"]).TrimEnd() + ".";

                                    //                                                try 
                                    //                                                {
                                    //                                                    par.Range.Text += "Дата начала периода работы: " + Convert.ToDateTime(row["DATST"]).ToShortDateString() + ".";
                                    //                                                    par.Range.Text += "Дата окончания периода работы: " + Convert.ToDateTime(row["DATFN"]).ToShortDateString() + ".";                                                
                                    //                                                }
                                    //                                                catch
                                    //                                                {

                                    //                                                }
                                    //                                                //par.Range.Text += "Дата начала периода работы: " + Convert.ToDateTime(row["DATST"]).ToShortDateString() + ".";
                                    //                                                //par.Range.Text += "Дата окончания периода работы: " + Convert.ToDateTime(row["DATFN"]).ToShortDateString() + ".";

                                    //                                                par.Range.Text += "Комментарий: " + Convert.ToString(row["KOMMENT"]).TrimEnd() + "\n";
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
                                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                                }
                                
                                if (nNewPackID > 0)
                                {
                                    WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответы не удалось загрузить без ошибок.\n");
                                    // обновить статус лога-ответа
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ответ загружен с ошибкой
                                }

                                //return false;
                            }
                            catch (Exception ex)
                            {
                                //if (DBFcon != null) DBFcon.Close();
                                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);

                                if (nNewPackID > 0)
                                {
                                    WriteLLog(conGIBDD, nNewPackID, "Ошибка! Ответы не удалось загрузить без ошибок.\n");
                                    // обновить статус лога-ответа
                                    UpdateLLogStatus(conGIBDD, nNewPackID, 11); // ответ загружен с ошибкой
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

            // Все закомментировал - потому как update занимает слишком много времени
            
            ///// Внимание - для работы с интерфейсными таблицами отправка сторонней программой
            ///// не позволяет запуститься выгрузке сведений о запросе в интерфейсную таблицу, поэтому надо убирать этот вариант
            //// d.docstatusid = 23 это статус отправка сторонней программой 
            //string txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 23 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 3 and req_type.sndl_contr_id = " + potd_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// для автоматической отправки
            ////  d.docstatusid = 11 - это статус автоматическая отправка
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 11 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 2 and req_type.sndl_contr_id = " + potd_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);

            //// есть еще статус 21 - ручная отправка
            //// для автоматической отправки
            //txtUpdateSql = "UPDATE DOCUMENT d SET d.docstatusid = 21 WHERE d.docstatusid = 1 and d.METAOBJECTNAME = 'DX_PACK_O' and d.id IN (select d.id from sendlist_dbt_request_type req_type left join DX_PACK pk on pk.agreement_id = req_type.outer_agreement_id left join dx_pack_o pk_o on pk_o.id = pk.id left join dx_mvv_exchange_reglament regl on pk_o.EXCHANGE_REGLAMENT_ID = regl.id left join document d on d.id = pk.id where d.docstatusid = 1 and regl.sending_mode = 1 and req_type.sndl_contr_id = " + potd_id.ToString() + ")";
            //UpdateSqlExecute(con, txtUpdateSql);
                             

            //select 2 LITZDOLG, zapr_d.id ZAPROS, ip.div, ip.ID_DBTR_NAME FIOVK, ip.IPNO NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 15 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_NSI_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))
            //ReadPOTDData(DatZapr1_potd, DatZapr2_potd);
            //DT_potd_doc = GetDataTableFromFB("select 2 LITZDOLG, zapr_d.id ZAPROS, ip.IPNO_NUM, ip.div, ip.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, zapr_d.doc_date DATZAPR, ip.ID_DBTR_ADR ADDR, ip.ID_DBTR_BORN DATROZHD, ip.ID_DBTRCLS, ip.DBTR_BORN_YEAR GOD, ip.ID_DEBTSUM SUMMA, ip.ID_DBTR_INN INNORG, zapr_d.doc_number, ip.ID_DEBTCLS_NAME VIDVZISK from O_IP_REQ_IP req left join document zapr_d on req.id = zapr_d.id left join document ip_d on zapr_d.parent_id = ip_d.id left join o_ip ip on zapr_d.id = ip.id left join SPI on ip.IP_EXEC_PRIST = spi.SUSER_ID where req.o_ip_req_dbt_type = 206 and zapr_d.docstatusid = 2 and (ip.ID_DBTRCLS = 2 or (ip.ID_DBTRCLS in (select ncc_id from V_NSI_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");
            //DT_potd_doc = GetDataTableFromFB("select pack.id as pack_id, 2 LITZDOLG, d_req.id ZAPROS, req.IPNO_NUM, req.div, req.ID_DBTR_NAME FIOVK, ip_d.doc_number NOMIP, spi.SPI_ZONENUM NOMSPI, d_req.doc_date DATZAPR,  req.ID_DBTR_ADR ADDR, req.ID_DBTR_BORN DATROZHD, req.ID_DBTRCLS,  req.DBTR_BORN_YEAR GOD, req.ID_DEBTSUM SUMMA, req.ID_DBTR_INN INNORG,  d_req.doc_number, req.ID_DEBTCLS_NAME VIDVZISK from dx_pack_o packo left join dx_pack pack on pack.id = packo.id join sendlist sl on pack.id = sl.dx_pack_id join o_ip req on sl.sendlist_o_id = req.id join document d_req on req.id = d_req.id join document ip_d on d_req.parent_id = ip_d.id join document dpack on pack.id = dpack.id join SPI on req.IP_EXEC_PRIST = spi.SUSER_ID where dpack.docstatusid = 23  and pack.agreement_id = 110 and packo.has_been_sent is null  and d_req.docstatusid != 19 and d_req.docstatusid != 15   and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND");

            // выборка пакетов из Интерфейсных Таблиц (ИТ)
            // ext_request_id - нужен чтобы потом быстро сделать update
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
            //    DialogResult rv = MessageBox.Show("По пути " + string.Format(@"{0}\{1}", fullpath, tablename_n) + ", указанном в конфигурационном файле, существует файл. Удалить его?", "Внимание", MessageBoxButtons.YesNo);
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
            /***Выборка всех незаконченых ИП по штрафам ГИБДД***/
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
            txtResult = nMoney.ToString("N2").Replace(".", " руб. ");
            txtResult = txtResult.Replace(",", " руб. ") + " коп.";

            return txtResult;
        }

        private string Money_ToStr(double nMoney)
        {
            string txtResult = "";
            txtResult = nMoney.ToString("N2").Replace(".", " руб. ");
            txtResult = txtResult.Replace(",", " руб. ") + " коп.";

            return txtResult;
        }




        private void b_loadgibd_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "DBF файлы(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();
            OleDbConnection conGIBDD;

            if (res == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {
                    
                    //ChangeByte(openFileDialog1.FileName, 0x65, 30);
                    string tablename = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - 4);
                    tablename = tablename.Substring(tablename.LastIndexOf("\\") + 1);

                    // вычленить из имени файла дату исх. документа (реестра)
                    string txtDateIsh = tablename.Substring(6, 2) + '.' + tablename.Substring(4, 2) + '.' + tablename.Substring(0, 4);

                    DateTime dtDateIsh;
                    if (!DateTime.TryParse(txtDateIsh, out dtDateIsh)) {
                        dtDateIsh = DateTime.MinValue;
                    }

                    // вычленить из имени файла исх номер документа
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

                        // проверить не есть ли в списке уже загруженных реестр с исходящим номером txtIshNumber
                        // получить список всех загруженных реестров
                        ArrayList alLoadedReestrs = GetLoadedReestrs(conGIBDD);

                        if (alLoadedReestrs.Contains(txtIshNumber))
                        {
                            Exception ex = new Exception("Ошибка. Зафиксирована попытка повторно загрузить реестр " + txtIshNumber + ".");
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

                            // надо открыть подключение к б.д. и вставить в него сведения
                            // ключ - полный номер протокола
                            // дополнительное поле - сверка прошла успешно
                            // вставлять номер, дату, сумму, фамилию, фио, дату оплаты, номер реестра, дату реестра

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
                            m_cmd.Parameters.Add(new OleDbParameter(":DATID", dtDatID)); // Дата ИД
                            m_cmd.Parameters.Add(new OleDbParameter(":SUMM", nSum));     // Сумма ИД
                            m_cmd.Parameters.Add(new OleDbParameter(":SUMM_DOC", nSumDoc));   //Сумма оплаты
                            m_cmd.Parameters.Add(new OleDbParameter(":FIO_D", txtFioD));
                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_DOC", dtDateDoc)); // Дата оплаты
                            

                            m_cmd.Parameters.Add(new OleDbParameter(":ISH_NUMBER", txtIshNumber));   // исх номер
                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_ISH", dtDateIsh));   // дата исх 
                            m_cmd.Parameters.Add(new OleDbParameter(":FL_USE", Convert.ToInt32(0)));   // флаг - учтен/неучтен
                            m_cmd.Parameters.Add(new OleDbParameter(":NUM_DOC", txtNumDoc));   // номер квитанции
                            m_cmd.Parameters.Add(new OleDbParameter(":BORN_D", dtBornD));   // дата рождения должника
                            m_cmd.Parameters.Add(new OleDbParameter(":DATE_VH", DateTime.Today));   // дата загрузки реестра
                            
                            
                            
                            int result = m_cmd.ExecuteNonQuery();

                            if (result != -1)
                            {
                                iCnt++;
                                prbWritingDBF.PerformStep();
                            }
                        }
                        tran.Commit();
                        conGIBDD.Close();
                        MessageBox.Show("Данные успешно загружены.", "Сообщение", MessageBoxButtons.OK);
   
                    }
                    catch (OleDbException ole_ex)
                    {
                        foreach (OleDbError err in ole_ex.Errors)
                        {
                            MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                        }
                        //return false;
                    }
                    catch (Exception ex)
                    {
                        //if (DBFcon != null) DBFcon.Close();
                        MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
            //DT_reg = GetDataTableFromFB("SELECT DISTINCT a.NAME_D as FIOVK, a.NUM_IP as ZAPROS, a.VIDD_KEY as LITZDOLG, a.DATE_BORN_D as GOD, a.USCODE as NOMSPI, a.NUM_IP as NOMIP, a.SUM_ as SUMMA, a.WHY as VIDVZISK, a.INND as INNORG, a.DATE_IP_IN as DATZAPR, a.ADR_D as ADDR,a.TEXT_PP as OSNOKON, a.PK as FK_IP, a.PK_ID as FK_ID FROM IP a WHERE a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "' AND NUM_IP NOT LIKE '%СД' AND a.NUM_IP NOT LIKE '%СВ' AND NUM_IP NOT LIKE '%!%' ORDER BY a.NAME_D", "TOFIND", IsolationLevel.ReadCommitted);
            //DT_okon = GetDataTableFromFB("SELECT DISTINCT a.NAME_D as FIOVK, a.NUM_IP as ZAPROS, a.VIDD_KEY as LITZDOLG, a.DATE_BORN_D as GOD, a.USCODE as NOMSPI, a.NUM_IP as NOMIP, a.SUM_ as SUMMA, a.WHY as VIDVZISK, a.INND as INNORG, a.DATE_IP_OUT as DATZAPR, a.ADR_D as ADDR,a.TEXT_PP as OSNOKON FROM IP a WHERE a.DATE_IP_OUT is not null AND NUM_IP NOT LIKE '%!%' and DATE_IP_OUT <= '" + dat2.ToShortDateString() + "' and DATE_IP_OUT >= '" + dat1.ToShortDateString() + "'", "TOFIND", IsolationLevel.ReadCommitted);
            //end = DateTime.Now;
            //lblTime2.Text = Convert.ToString(((TimeSpan)(end - start)).TotalMilliseconds);

            //start = DateTime.Now;
            //DT_ktfoms_reg = GetDataTableFromFB(" SELECT DISTINCT a.USCODE as NOMSPI, a.NUM_IP as ZAPROS, a.sdc as NOMOTD, a.name_d as FIOVK, a.DATE_BORN_D as DATROZHD, a.ADR_D as ADDR, a.PK as FK_IP, a.PK_ID as FK_ID FROM IP a LEFT JOIN DOCUMENT b ON b.FK = a.PK WHERE (a.DATE_IP_IN >= '" + dat1.ToShortDateString() + "' AND a.DATE_IP_IN <= '" + dat2.ToShortDateString() + "') and a.VIDD_KEY LIKE '/1/%' AND a.NUM_IP NOT LIKE '%СД' AND a.NUM_IP NOT LIKE '%СВ' AND a.NUM_IP NOT LIKE '%!%'", "TOFIND", IsolationLevel.ReadCommitted);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + odbc_err.Message + "Native Error: " + odbc_err.NativeError + "Source: " + odbc_err.Source + "SQL State   : " + odbc_err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }

            MessageBox.Show("Загружено строк: " + myTbl.Rows.Count.ToString(), "Внимание!", MessageBoxButtons.OK);
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
            openFileDialog1.Filter = "DBF файлы(*.dbf)|*.dbf";
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
                        MessageBox.Show("Обработано ответов: " + iCnt.ToString() + "\n. Сейчас будет сформирован реестр ответов.", "Сообщение", MessageBoxButtons.OK);

                    }
                    catch (OleDbException ole_ex)
                    {
                        foreach (OleDbError err in ole_ex.Errors)
                        {
                            MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                        }
                        //return false;
                    }
                    catch (Exception ex)
                    {
                        //if (DBFcon != null) DBFcon.Close();
                        MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                        //return false;
                    }
                    //return true;

                }
            }
        }

        private void b_delgibdd_Click(object sender, EventArgs e)
        {
            /***Загрузка запросов для тех ИП по которым поступили платежи в ГИБДД***/
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }

        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            /***Загрузка запросов для тех ИП по которым поступили платежи в ГИБДД***/
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }


            prbWritingDBF.Value = 0;
            //prbWritingDBF.Maximum = DT_gibd_reg.Rows.Count;
            prbWritingDBF.Maximum = DT_gibd_rst.Rows.Count;
            prbWritingDBF.Step = 1;


            try
            {
                string txtGibdName = "орган ГИБДД";
                string txtGibdConv = "соглашением об обмене информацией";

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

                    //par.Range.Text = "Привет!";
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

                            par.Range.Text += "РЕЕСТР ИСПОЛНИТЕЛЬНЫХ ДОКУМЕНТОВ, ДОЛГ ПО КОТОРЫМ ОПЛАЧЕН.";
                            par.Range.Text += "СФОРМИРОВАН НА ОСНОВЕ ДАННЫХ, ПОЛУЧЕННЫХ ИЗ ГИБДД.\n";
                            par.Range.Text += "ДАТА ФОРМИРОВАНИЯ " + DatZapr.ToShortDateString() + "\n";

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
                            par.Range.Text += "НОМЕР ИД       ДОЛЖНИК     ВЗЫСКАТЕЛЬ     НОМЕР ИП       ДАТА ВНЕСЕНИЯ В БАЗУ ГИБДД      ДАТА ЗАГРУЗКИ";
                            //par.Range.Text += GetOSP_Name();
                            sch_line += 8;
                        }
                        if (spi == Convert.ToInt32(row["USCODE"]))
                        {
                            // тупняк какой-то svn глючит!
                            string txtResponse = Convert.ToString(row["NOMID"]) /*+ "  " + Money_ToStr(Convert.ToDecimal(row["summ"]))*/ + "  " + Convert.ToString(row["FIO_D"] + "  " + Convert.ToString(row["name_v"]) + "  " + Convert.ToString(row["NUM_IP"])) + "  " + Convert.ToString(Convert.ToDateTime(row["BASE_T"]).ToShortDateString()) + "  " + Convert.ToString(Convert.ToDateTime(row["DATE_Z"]).ToShortDateString());
                            //sch_line++;
                            //string txtResponse = Convert.ToString(row["BASE_T"]) + " " + Convert.ToString(row["DATE_Z"]);
                            par.Range.Text += txtResponse;
                            sch_line++;
                            if (txtResponse.Length > 200)
                            {
                                sch_line++; // если был перевод строки
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                if (con != null) con.Close();
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
            openFileDialog1.Filter = "DBF файлы(*.dbf)|*.dbf";
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


                            //    //if (!(DateTime.TryParse(Convert.ToString(row["DATZAPR"]), out DatZapr)))// реально это дата регистрации ИП
                            //    //{
                            //    //    DatZapr = DateTime.Today;
                            //    //}
                            //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_ZAPR", DateTime.Today));

                            //    // а вот тут надо строку анализировать

                            //    vid_d = 1; // физ. лицо
                            //    //if (Convert.ToString(row["LITZDOLG"]).StartsWith("/1/"))
                            //    //{
                            //    //    vid_d = 1;// физ. лицо
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

                            //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_RES", ""));// это номер ответа от взаимод орг-ии

                            //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_RES", DateTime.Today));// дата ответа, не забыть обновить при ответе

                            //    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(1)));// (0 - ДОЛЖНИК НЕ ИДЕНТИФИЦИРОВАН, 1 - НЕТ ИНФ. ПО ДОЛЖНИКУ, БОЛЬШЕ 1 - ЕСТЬ ИНФ-Я ПО ДОЛЖНИКУ) (ВСЕ ЧТО В FIND - ВСЕ БОЛЬШЕ 1, ИЗНАЧАЛЬНО 0)

                            //    Int32 iKey = -1;
                            //    //if (!Int32.TryParse(Convert.ToString(row["FK_DOC"]), out iKey))
                            //    //{
                            //    //    iKey = 0;
                            //    //}

                            //    // TODO: ПРОВЕРЬ, ЧТО ЭТО РЕАЛЬНО ССЫЛКА НА DOCUMENTS!!!
                            //    // надо еще в таблицу DOCUMENTS вставлять запись, пока не когда
                            //    m_cmd.Parameters.Add(new OleDbParameter(":FK_DOC", iKey));

                            //    iKey = 0;
                            //    m_cmd.Parameters.Add(new OleDbParameter(":FK_IP", iKey));
                            //    m_cmd.Parameters.Add(new OleDbParameter(":FK_ID", iKey));


                            //    //OleDbCommand num_id_cmd = new OleDbCommand("Select NUM_ID from ID where ID.PK = " + Convert.ToString(row["FK_ID"]), con, tran);
                            //    //String num_id = Convert.ToString(num_id_cmd.ExecuteScalar());

                            //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ID", Convert.ToString(row["NOMID"])));

                            //    if (Convert.ToInt32(row["SUMPL"])!=0)
                            //        m_cmd.Parameters.Add(new OleDbParameter(":TEXT", "Должник выплатил сумму задолженности: " + Convert.ToString(row["SUMPL"]) + " . Дата выплаты: " + Convert.ToString(row["DATPL"])));// текст ответа - пустой
                            //    else
                            //        m_cmd.Parameters.Add(new OleDbParameter(":TEXT", "За период с " + Convert.ToString(row["DATZAPR1"]) + " по " + Convert.ToString(row["DATZAPR1"]) + " платежей не поступало"));// текст ответа - пустой

                            //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_BEG", DatZapr1_krc));
                            //    m_cmd.Parameters.Add(new OleDbParameter(":DATE_END", DatZapr2_krc));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":ADRESS", cutEnd(Convert.ToString(row["ADRES"]).Trim(), 250)));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_PACK", cutEnd(uid.Trim(), 30)));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":NUM_ZAPR_IN_PACK", iCnt));

                            //    //m_cmd.Parameters.Add(new OleDbParameter(":ADRESAT", Convert.ToString(Legal_Name_List[0])));

                            //    m_cmd.Parameters.Add(new OleDbParameter(":STATUS", "ответ"));

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

                            //    // лезем в таблицу persons и там ищем нужные данные
                            //    // таблицу persons по ip.fk; в таблицу physical по person.tablename + person.FK
                            //    // select PH.SER_PASSPORT, PH.NOMPASSPORT, PH.D_PASS FROM PERSON PR LEFT JOIN PHYSICAL PH ON PR.FK = PH.PK WHERE PR.TABLENAME=1 AND PR.MAIN = 1 AND PR.FK_IP = @FK_IP
                            //    //

                            //    m_cmd.Parameters.Add(new OleDbParameter(":PASSPORT", ""));// потом напишу пасп. данные

                            //    Double sum = 0;
                            //    //if (!(Double.TryParse(Convert.ToString(row["SUMMA"]), out sum)))
                            //    //{
                            //    //    sum = 0;
                            //    //}
                            //    m_cmd.Parameters.Add(new OleDbParameter(":SUMM", sum));

                            //    // TODO: пишет что нет NOMIP в tofind
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
                                m_cmd.CommandText = "UPDATE ZAPROS SET RESULT = :RESULT, TEXT = :TEXT, DATE_RESP = :DATE_RESP, DATE_RES = :DATOTV, STATUS = 'ответ'";
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
                                //    txtResponse = "Должник является получателем пенсии.\n";
                                //    txtResponse += "Адрес: " + Convert.ToString(row["ADRES"]).TrimEnd() + "\n";
                                //    txtResponse += "Cумма дохода, на которую можно обратить взыскание: " + Convert.ToString(row["SUMMA"]).TrimEnd() + ". " + Convert.ToString(row["KOMMENT"]).TrimEnd() + "\n";

                                //}
                                //else
                                //{
                                //    if (priz == 0)
                                //    {
                                //        m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));
                                //        txtResponse = "нет данных о должнике по запросу от " + Convert.ToDateTime(row["DATZAP"]).ToShortDateString();
                                //    }
                                //    else
                                //    {
                                //        m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));
                                //        txtResponse = "нет данных о должнике по запросу от " + Convert.ToDateTime(row["DATZAP"]).ToShortDateString() + " " + Convert.ToString(row["SUMMA"]).TrimEnd();
                                //    }
                                //}
                                if (Convert.ToInt32(row["SUMPL"]) != 0)
                                {
                                    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(1)));
                                    m_cmd.Parameters.Add(new OleDbParameter(":TEXT", "Должник выплатил сумму задолженности: " + Convert.ToString(row["SUMPL"]) + " . Дата выплаты: " + Convert.ToString(row["DATPL"])));// текст ответа - пустой
                                }
                                else
                                {
                                    m_cmd.Parameters.Add(new OleDbParameter(":RESULT", Convert.ToInt32(0)));
                                    m_cmd.Parameters.Add(new OleDbParameter(":TEXT", "По данному исполнительному документу платежей не поступало"));// текст ответа - пустой
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
                                //    MessageBox.Show("Ура, записали ответ!", "Внимание!", MessageBoxButtons.OK);
                                //}
                            }
                            tran.Commit();
                            con.Close();
                            MessageBox.Show("Обработано ответов: " + iCnt.ToString() + ".\n Сейчас будет сформирован реестр ответов.", "Сообщение", MessageBoxButtons.OK);

                            //**********Формирование**реестра**pens************
                            //Надо вспомнить запись в Ворд + подсчитывая количество строк 
                            //для разделения по приставам. 

                            //список всех приставов
                            DataTable dtspi = ds.Tables.Add("SPI");

                            DBFcon.Open();
                            m_cmd = new OleDbCommand();
                            m_cmd.Connection = DBFcon;
                            // добавил фильтр по полю priz чтобы не показывать пустые страницы
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
                            //    //      пример для Ворда

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

                            //        par.Range.Text += "Реестр ответов по запросам СП-И в ПФР о наличии пенсии\n";
                            //        par.Range.Text += "На запрос в ПФР от " + DateTime.Today.ToShortDateString() + "\n";
                            //        //par.Range.Text += "За период с " + dat1.ToShortDateString() + " по " + dat2.ToShortDateString() + "\n";

                            //        spi = Convert.ToInt32(drspi["NOMSPI"]);

                            //        par.Range.Text += "СП-И: " + GetSpiName3(Convert.ToInt32(drspi["NOMSPI"])) + "\n";

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
                            //                    par.Range.Text += "Должник является получателем пенсии. Cумма дохода, на которую можно обратить взыскание: " + Convert.ToString(row["SUMMA"]).TrimEnd() + "\n";
                            //                    sch_line += 5;
                            //                }

                            //                //if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                            //                //string priz = Convert.ToString(row["PRIZ"]).TrimEnd();
                            //                //if (priz.ToUpper().Equals("T"))
                            //            }
                            //        }
                            //        // если ничего положительного в ответах нет, то так и пишем
                            //        if (sch_line == 6)
                            //        {
                            //            par.Range.Text += "Нет положительных ответов по запросам о наличии пенсии у должников.";
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
                                MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                            }
                            //return false;
                        }
                        catch (Exception ex)
                        {
                            //if (DBFcon != null) DBFcon.Close();
                            MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("Выберите файл для разделения его на несколько частей.");
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
                        MessageBox.Show("Внимание. Файл разрезан на " + nNumOfNewFiles.ToString() + " частей.\nТеперь обрабатывать надо файлы с суффиксом .partX\n X - порядковый номер файла");
                    else MessageBox.Show("Внимание. Файл не требуется делить на части.\n");
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

                        // удалить все учтенные и все использованные
                        DeleteUsedGibddPlat(ConG, true, true);

                        DataSet ds = new DataSet();
                        string txtSql = "SELECT * FROM " + tablename + " WHERE  FL_USE = 0"; // выбрать только те, которые еще не были рассмотрены
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

                            // сделать запрос в базу ПК ОСП для поиска такого номера ИД
                            decimal id = FindIDNum(txtNomID, nSumID, dtDatID);
                            if (id > 0)
                            {
                                iCnt++;
                                // вычистить лишние пробелы из ФИО
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

                                // заполнить txtContent
                                txtContent = "Должник " + txtFIO_D + " "; ;
                                if (!dtBornD.Equals(Convert.ToDateTime("01.01.1800")))
                                {
                                    txtContent += "(дата рождения " + dtBornD.ToShortDateString() + ") ";
                                }
                                txtContent += dtDateDoc.ToShortDateString() + " оплатил " + Money_ToStr(nSumDoc) + " № документа об оплате " + txtNumKvit + " по ИД № " + txtNomID + " от " + dtDatID.ToShortDateString() + ".";

                                //MessageBox.Show("Нашли ИД № " + txtNomID + ". IP_ID = " + id.ToString(), "Внимание!", MessageBoxButtons.OK);
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


                        MessageBox.Show("Данные успешно проверены.", "Сообщение", MessageBoxButtons.OK);

                    }
                    catch (OleDbException ole_ex)
                    {
                        foreach (OleDbError err in ole_ex.Errors)
                        {
                            MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                        }
                        //return false;
                    }
                    catch (Exception ex)
                    {
                        //if (DBFcon != null) DBFcon.Close();
                        MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", System.Windows.Forms.MessageBoxButtons.OK);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", System.Windows.Forms.MessageBoxButtons.OK);
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
            // загрузка данных об оплаченных суммах по долгам из КРЦ
            DT_krc_oplat = null;
            Int32 iCnt = 0;
            openFileDialog1.Filter = "DBF файлы(*.dbf)|*.dbf";
            DialogResult res = openFileDialog1.ShowDialog();

            if (res == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {

                    //ChangeByte(openFileDialog1.FileName, 0x65, 30);
                    string tablename = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - 4);
                    tablename = tablename.Substring(tablename.LastIndexOf("\\") + 1);

                    // вычленить из имени файла дату исх. документа (реестра)
                    string txtDateIsh = tablename.Substring(6, 2) + '.' + tablename.Substring(4, 2) + '.' + tablename.Substring(0, 4);

                    DateTime dtDateIsh; 
                    if (!DateTime.TryParse(txtDateIsh, out dtDateIsh)) // +
                    {
                        dtDateIsh = DateTime.MinValue;
                    }

                    // вычленить из имени файла исх номер документа
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
                            MessageBox.Show("Данные успешно подгружены, запускается процедура сверки.", "Сообщение", MessageBoxButtons.OK);
                        }
                        // данные загрузили - теперь нужно провести сверку


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

                                    // Пока закомментируем - пока не выясним что нужно на самом деле
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
                                    //m_cmd.Parameters.Add(new OleDbParameter(":DATID", dtDatID)); // Дата ИД
                                    //m_cmd.Parameters.Add(new OleDbParameter(":SUMM", nSum));     // Сумма ИД
                                    //m_cmd.Parameters.Add(new OleDbParameter(":SUMM_DOC", nSumDoc));   //Сумма оплаты
                                    //m_cmd.Parameters.Add(new OleDbParameter(":FIO_D", txtFioD));
                                    //m_cmd.Parameters.Add(new OleDbParameter(":DATE_DOC", dtDateDoc)); // Дата оплаты


                                    //m_cmd.Parameters.Add(new OleDbParameter(":ISH_NUMBER", txtIshNumber));   // исх номер
                                    //m_cmd.Parameters.Add(new OleDbParameter(":DATE_ISH", dtDateIsh));   // дата исх 
                                    //m_cmd.Parameters.Add(new OleDbParameter(":FL_USE", Convert.ToInt32(0)));   // флаг - учтен/неучтен
                                    //m_cmd.Parameters.Add(new OleDbParameter(":NUM_DOC", txtNumDoc));   // номер квитанции
                                    //m_cmd.Parameters.Add(new OleDbParameter(":BORN_D", dtBornD));   // дата рождения должника
                                    //m_cmd.Parameters.Add(new OleDbParameter(":DATE_VH", DateTime.Today));   // дата загрузки реестра

                                
                                    // сделать запрос в базу ПК ОСП для поиска такого номера ИД
                                    decimal id = FindIDNum(txtGibddIDNumber, nSum, dtDatID);
                                    if (id > 0)
                                    {
                                        iCnt++;
                                        // вычистить лишние пробелы из ФИО
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

                                        // заполнить txtContent
                                        txtContent = "Должник " + txtFioD + " "; ;
                                        if (!dtBornD.Equals(Convert.ToDateTime("01.01.1800")))
                                        {
                                            txtContent += "(дата рождения " + dtBornD.ToShortDateString() + ") ";
                                        }
                                        txtContent += dtDateDoc.ToShortDateString() + " оплатил " + Money_ToStr(nSumDoc) + " № документа об оплате " + txtNumDoc + " по ИД № " + txtGibddIDNumber + " от " + dtDatID.ToShortDateString() + ".";

                                        //MessageBox.Show("Нашли ИД № " + txtNomID + ". IP_ID = " + id.ToString(), "Внимание!", MessageBoxButtons.OK);
                                        
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

                            MessageBox.Show("Данные успешно проверены.", "Сообщение", MessageBoxButtons.OK);
                        }

                    }
                    catch (OleDbException ole_ex)
                    {
                        foreach (OleDbError err in ole_ex.Errors)
                        {
                            MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                        }
                        //return false;
                    }
                    catch (Exception ex)
                    {
                        //if (DBFcon != null) DBFcon.Close();
                        MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

                // получить новый ключ
                cmd = new OleDbCommand("SELECT gen_id(GEN_LOCAL_LOG_ID, 1) FROM RDB$DATABASE", gibdd_con, tran);
                nID = Convert.ToDecimal(cmd.ExecuteScalar());

                // получить OSPNUM
                nOspNum = 10000 + GetOSP_Num();

                // вставить DOCUMENT
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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
            // заглушка 
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
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

                // получить новый ключ
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
                    MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
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
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
            }
            return nID;

        }

        private void сводныеРеестрОбработкиЗапросовИОтветовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // получить дату начала и дату окончания
            frmSelectDate SelectDate = new frmSelectDate();
            DatePeriod DateStartEnd;
            DateStartEnd = SelectDate.ShowForm();

            // проверить что пользователь не нажал Cancel  DateStartEnd.DateStart/End = DateTime.MinValue
            if (!((DateStartEnd.DateStart == DateTime.MinValue) && (DateStartEnd.DateEnd == DateTime.MinValue)))
            {
                DataTable dtReport = null;
                ReportMaker report = new ReportMaker();
                report.StartReport();
                report.AddToReport("<h3>");
                report.AddToReport("Журнал учета работы с электронными запросами и ответами в " + GetOSP_Name() + "<br />");
                report.AddToReport("</h3>");
                report.AddToReport("<p>Дата формирования: " + DateTime.Now.ToString() + "</p>");
                string txtSql = "select ll.id, ll.packdate, agr.name_agreement, ll.pack_count, ls.status_name,  ll.pack_type, ll_resp.pack_type, pt.type as resp_type, ls_resp.status_name as resp_status, ll_resp.packdate as resp_date, ll_resp.pack_count as resp_count, ll.fl_find, ll.fl_nofind, ll.fl_e_tofind from local_logs ll join agreements agr on ll.conv_code = agr.agreement_code join logs_status ls on ll.pack_status = ls.id left join local_logs ll_resp on ll_resp.parent_id = ll.id left join pack_type pt on ll_resp.pack_type = pt.id left join logs_status ls_resp on ls_resp.id = ll_resp.pack_status and ll.packdate >= '" + DateStartEnd.DateStart.ToShortDateString() + "' and ll.packdate <='" + DateStartEnd.DateEnd.ToShortDateString() + "' where ll.pack_type = 1 order by ll.id";

                dtReport = GetDataTableFromFB(constrGIBDD, txtSql, "report", IsolationLevel.Unspecified);
                DateTime dPackDate, dRespDate;
                string txtPackDate, txtRespDate, txtAgr, txtReqCount, txtRespCount, txtReqStatus, txtRespStatus, txtRespType;
                Decimal nReqCount, nRespCount;

                if (dtReport != null)
                {
                    report.AddToReport("<table border=\"1\" cellpadding=\"2\" style=\"border:1px #000000 solid;\"><tbody>");

                    report.AddToReport("<tr>");
                    report.AddToReport("<td>Дата выгрузки запросов</td>");
                    report.AddToReport("<td>Вид запросов</td>");
                    report.AddToReport("<td>Количество выгруженных запросов</td>");
                    report.AddToReport("<td>Статус пакета запросов</td>");
                    report.AddToReport("<td>Дата обработки ответов</td>");
                    report.AddToReport("<td>Тип ответов</td>");
                    report.AddToReport("<td>Количество обработанных ответов</td>");
                    report.AddToReport("<td>Статус пакета ответов</td>");
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

