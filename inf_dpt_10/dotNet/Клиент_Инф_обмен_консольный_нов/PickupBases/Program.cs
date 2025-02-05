using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Collections;
using System.Data.OleDb;
using System.Data;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Net;
using System.Threading;
using System.Net.Mail;
using System.Net.Mime;



namespace InfoChange
{
    class PickupBases
    {

        static bool b20reliz, bAutoLoadGibdd, bAutoLoadKRC, bAutoLoadRicZH, bAutoLoadKESK, bAutoLoadTGK1, bAutoLoadSberKvit, bAutoLoadSberReport, bAutoLoadVU, bAutoSberReqOut, bAutoLoadGIMS, bAutoSberRespIn, bAutoSberDocsOut, bRunSverka, bRunGibdd10_out, bRunRtn10_selfrequest, bRunIcmvdFms10_selfrequest, bRunGimsLodka10_selfrequest, bRunGibdd10_selfrequest, bRunGims10_selfrequest, bRunIC_MVD_out, bRunIC_MVD_in, bRunCredOrgReqOut, bRunCredOrgAnsIn, bRunPotdOut, bRunPotdOut2017, bRunPotdIn, bRunPotdIn2017, bRunPfrRepIn;
        static string txtUploadDirGibdd, txtUploadDirKESK, txtUploadDirTGK1, txtUploadDirRicZH, txtUploadDirKRC, txtUploadDirSberKvit, txtUploadDirSberReport, txtUploadDirVU, txtUploadDirGims, txtSberReqOutPath, txtSberRespInPath, txtSberDocsPath, txtGibddOutPath, txtGibdd10ConString, txtGibdd10DataBase, txtIC_MVD_Path, txtIC_MVD_Path_in, txtCredOrgReq_Path_out, txtPotdOutPath, txtPotdInPath, txtPfrRepIn, txtOldPwd, txtNewPwd;
        //string txtCistomLegalIds = "86200999999007, 86200999999009";
        string txtCistomLegalIds = "86200999999009"; // убрали 86200999999007 - БаренцБанк
        static string txtZaprosOut = "0"; // если '1' - то это понедельник

        private static ArrayList ReadPaths(string FromFilename)
        {
            ArrayList Filepaths = new ArrayList();
            using (StreamReader sr = new StreamReader(FromFilename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line[0] != '#')
                    {
                        Filepaths.Add(line);
                    }
                }
            }
            return Filepaths;
        }

        private static ArrayList ReadPaths(string FromFilename, Encoding enc)
        {
            ArrayList Filepaths = new ArrayList();
            using (StreamReader sr = new StreamReader(FromFilename, enc))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line[0] != '#')
                    {
                        Filepaths.Add(line);
                    }
                }
            }
            return Filepaths;
        }

        private static DataTable GetDataTableFromFB(string txtSql, string tblName,string connectionstring)
        {
            OleDbConnection con = new OleDbConnection(connectionstring);
            DataSet ds = new DataSet();
            DataTable tbl = ds.Tables.Add(tblName);
            try
            {
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
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
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
            }
            return tbl;
        }
        
        //TODO: Разработать функцию для импорта в mysql
        //Добавить строку в файл Outfile.txt
        private static void InsertDataToDB(string tblName, string connectionstring,DataTable dt, string txtSql)
        {
            OleDbConnection con = new OleDbConnection(connectionstring);
            try
            {
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.RepeatableRead);
                OleDbCommand cmd = new OleDbCommand(txtSql, con, tran);
                foreach (DataRow dr in dt.Rows)
                {
                    cmd.ExecuteNonQuery();
                    tran.Commit();
                    con.Close();
                }
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
            }
        }


        private static bool WriteTofile(string txtText, string outfile)
        {
            using (StreamWriter sw = new StreamWriter(outfile, true))
            {
                sw.WriteLine(txtText);
                sw.Close();
            }
            return true;

        }

        static Dictionary<int, string> GetEmails()
        {
            Dictionary<int, string> OspEmails = new Dictionary<int, string>();
            OspEmails.Add(0, "mvv_report@r10.fssp.gov.ru"); // нулевой адрес для отправки отчета в ОИиОИБ
            for (int i = 1; i <= 24; i++)
            {

                OspEmails.Add(i, "osp" + i.ToString().PadLeft(2, '0') + "@r10.fssp.gov.ru");
            }

            return OspEmails;
        }


        
        static void Main(string[] argc)
        {

            //if (argc.Length < 2)
            //{
            //    Console.WriteLine("usage:\nPickupbases infile outfile\ninfile contain connectionstrings for Firebird bases\noutfile contain Information for MySql base");
            //    return;
            //}
            //string FromFilename = argc[0];
            //string ToFilename = argc[1];
            
            string txtLogFileName = "sverka_gibdd.log";
            int iTotalFiles = 0;
            string txtMailServ = "mail10";

            FormUpdater up = new FormUpdater(argc, txtLogFileName);
           
            if (up.skipped())
            {
                //  Application.Run(new Form1());

                File_funcs ff = new File_funcs(); // чтобы можно было ff функции вызывать

                string FromFilename = "InFile.txt";
                decimal mvd_id;
                string txtErr = "";

                try
                {
                    Dictionary<int, string> OspEmails = new Dictionary<int, string>();
                    OspEmails = GetEmails();

                    ArrayList InPaths = ReadPaths(FromFilename, Encoding.GetEncoding(1251));

                    string txtConStrPKOSP, constrGIBDD, txtConStrED;
                    txtConStrPKOSP = InPaths[0].ToString();
                    constrGIBDD = InPaths[1].ToString();

                    // если в строке подключения есть <domen>.karelia.ssp то это надо убирать
                    // варианты - сразу поменять на rdb
                    // попробовать отрезать по маске и сделать то что идет первым, до имени домена
                    // пробуем первый вариант Provider=LCPI.IBProvider.3;Data Source=rdb-petr1/3051;Persist Security Info=True;Password=ksv1193;User ID=SYSDBA;Location=rdb-petr1/3051:ncore-fssp;ctype=win1251
                    txtConStrPKOSP = ff.RemoveDomainFromConString(txtConStrPKOSP);
                    constrGIBDD = ff.RemoveDomainFromConString(constrGIBDD);

                    mvd_id = Convert.ToDecimal(InPaths[2].ToString());

                    txtLogFileName = InPaths[3].ToString();

                    txtOldPwd = "ksv1193";

                    if (InPaths.Count > 53)
                    {
                        txtOldPwd = InPaths[53].ToString();

                    }

                    txtNewPwd = "Rkfcnth2023";
                    if (InPaths.Count > 54)
                    {
                        txtNewPwd = InPaths[54].ToString();
                    }

                    // проведем замену паролей
                    txtConStrPKOSP = ff.ReplacePwdConString(txtConStrPKOSP, txtOldPwd, txtNewPwd);
                    constrGIBDD = ff.ReplacePwdConString(constrGIBDD, txtOldPwd, txtNewPwd);

                    PKOSP_mvv mvv = new PKOSP_mvv();
                    OleDbConnection conPK_OSP = new OleDbConnection(txtConStrPKOSP);
                    decimal nOspNum = mvv.GetOSP_Num(conPK_OSP, out txtErr);
                    int iDiv = Convert.ToInt32(nOspNum);

                    // получить код ОСП
                    // если код = 24 - то база = \\\\fs-petr13\\Inf_obmen_24\\
                    // если код = 13 - то база = \\\\fs\\inf_obmen_petr13\\
                    // остальные коды - база = \\\\fs\\Inf_obmen\\
                    string txtPathBase = "\\\\fs\\Inf_obmen\\";

                    
                    switch (iDiv)
                    {
                        case 15: txtPathBase = "\\\\rdb-pud\\Inf_obmen\\"; break;
                        case 14: txtPathBase = "\\\\fs-pryag\\Inf_obmen\\"; break;
                        case 13: txtPathBase = "H:\\"; break;
                        case 7: txtPathBase = "\\\\rdb-lahd\\Inf_obmen\\"; break;
                        case 2: txtPathBase = "\\\\rdb-bel\\Inf_obmen\\"; break;
                        // case 21: txtPathBase = "\\\\rdb-petr3\\Inf_obmen\\"; break;
                        case 21: txtPathBase = "U:\\"; break;
                        // case 1: txtPathBase = "\\\\rdb-petr1\\Inf_obmen\\"; break;
                        case 1: txtPathBase = "X:\\"; break;
                        // case 20: txtPathBase = "\\\\rdb-petr2\\Inf_obmen\\"; break;
                        case 20: txtPathBase = "T:\\"; break;
                        case 11: txtPathBase = "\\\\rdb-muez\\Inf_obmen\\"; break;
                        case 6: txtPathBase = "\\\\rdb-kost\\Inf_obmen\\"; break;
                        // case 24: txtPathBase = "\\\\rdb-mosp24\\Inf_obmen\\"; break;
                        case 24: txtPathBase = "W:\\"; break;
                        // case 24: txtPathBase = "\\\\fs\\Inf_obmen\\"; break;
                        default: txtPathBase = "\\\\fs\\Inf_obmen\\"; break;
                    }


                    // добавили новый параметр - 20 релиз
                    b20reliz = true;
                    if (InPaths.Count > 4)
                    {
                        //b20reliz = Convert.ToBoolean(InPaths[4].ToString());
                        txtZaprosOut = InPaths[4].ToString().Trim();
                    }

                    // добавили новый параметр - каталог для поиска реестров оплач. штрафов
                    txtUploadDirGibdd = txtPathBase + "ответ\\Gibdd_reestr_in";
                    if (InPaths.Count > 5)
                    {
                        txtUploadDirGibdd = InPaths[5].ToString();
                    }

                    // добавили новый параметр - каталог для поиска реестров оплач. штрафов
                    bAutoLoadGibdd = true;
                    if (InPaths.Count > 6)
                    {
                        bAutoLoadGibdd = Convert.ToBoolean(InPaths[6].ToString());
                    }
                    // 20180518 - больше не грузим ничего из ГИБДД - поэтому блокируем
                    bAutoLoadGibdd = false;

                    txtUploadDirKRC = txtPathBase + "ответ\\КРЦ";
                    if (InPaths.Count > 7)
                    {
                        txtUploadDirKRC = InPaths[7].ToString();
                    }

                    bAutoLoadKRC = false;
                    if (InPaths.Count > 8)
                    {
                        bAutoLoadKRC = Convert.ToBoolean(InPaths[8].ToString());
                    }
                    bAutoLoadKRC = false; // отключено т.к. нет обмена с КРЦ

                    txtUploadDirSberKvit = txtPathBase + "EDO\\to_osp\\sber";
                    if (InPaths.Count > 9)
                    {
                        txtUploadDirSberKvit = InPaths[9].ToString();
                    }

                    bAutoLoadSberKvit = true;
                    if (InPaths.Count > 10)
                    {
                        bAutoLoadSberKvit = Convert.ToBoolean(InPaths[10].ToString());
                    }
                    // 20200208  - все отключаем кроме ic_mvd и квитанция на постановления рег. ПФР
                    bAutoLoadSberKvit = false;

                    txtSberReqOutPath = txtPathBase + "sberbank_xml\\Запросы";
                    if (InPaths.Count > 11)
                    {
                        txtSberReqOutPath = InPaths[11].ToString();
                    }

                    bAutoSberReqOut = false;
                    if (InPaths.Count > 12)
                    {
                        bAutoSberReqOut = Convert.ToBoolean(InPaths[12].ToString());
                    }
                    bAutoSberReqOut = false; // не выгружаем запросы в Сбер по Рег.МВВ

                    txtSberRespInPath = txtPathBase + "sberbank_xml\\Ответы";
                    if (InPaths.Count > 13)
                    {
                        txtSberRespInPath = InPaths[13].ToString();
                    }

                    bAutoSberRespIn = false;
                    if (InPaths.Count > 14)
                    {
                        bAutoSberRespIn = Convert.ToBoolean(InPaths[14].ToString());
                    }
                    bAutoSberRespIn = false;  // не загружаем ответы в Сбер по Рег.МВВ


                    txtSberDocsPath = txtPathBase + "EDO\\from_osp\\sber";
                    if (InPaths.Count > 15)
                    {
                        txtSberDocsPath = InPaths[15].ToString();
                    }

                    bAutoSberDocsOut = true;
                    if (InPaths.Count > 16)
                    {
                        bAutoSberDocsOut = Convert.ToBoolean(InPaths[16].ToString());
                    }

                    // 20200208  - все отключаем кроме ic_mvd и квитанция на постановления рег. ПФР
                    bAutoSberDocsOut = false;

                    txtUploadDirKESK = txtPathBase + "ответ\\КЭСК";
                    if (InPaths.Count > 17)
                    {
                        txtUploadDirKESK = InPaths[17].ToString();
                    }
                    bAutoLoadKESK = false;
                    if (InPaths.Count > 18)
                    {
                        bAutoLoadKESK = Convert.ToBoolean(InPaths[18].ToString());
                    }
                    bAutoLoadKESK = false; // отключено т.к. нет больше обмена с КЭСК

                    // 20160401 - сверку отключил - временно
                    // 20160504 - сверку включил обратно
                    bRunSverka = true;
                    if (InPaths.Count > 19)
                    {
                        bRunSverka = Convert.ToBoolean(InPaths[19].ToString());
                    }
                    // 20180515 - т.к. больше ни от МВД ни от ЖКХ у нас нет информации - то сверку больше не делаем
                    bRunSverka = false;

                    bRunGibdd10_out = false;
                    if (InPaths.Count > 20)
                    {
                        bRunGibdd10_out = Convert.ToBoolean(InPaths[20].ToString());
                    }
                    bRunGibdd10_out = false; // этого обмена нет вобще

                    txtGibddOutPath = txtPathBase + "запрос\\GIBDD";
                    if (InPaths.Count > 21)
                    {
                        txtGibddOutPath = InPaths[21].ToString();
                    }
                                        
                    bRunIC_MVD_out = true;
                    if (InPaths.Count > 22)
                    {
                        bRunIC_MVD_out = Convert.ToBoolean(InPaths[22].ToString());
                    }

                    txtIC_MVD_Path = txtPathBase + "запрос\\IC_MVD";
                    if (InPaths.Count > 23)
                    {
                        txtIC_MVD_Path = InPaths[23].ToString();
                    }

                    bAutoLoadSberReport = true;
                    if (InPaths.Count > 24)
                    {
                        bAutoLoadSberReport = Convert.ToBoolean(InPaths[24].ToString());
                    }
                    // 20200208  - все отключаем кроме ic_mvd и квитанция на постановления рег. ПФР
                    bAutoLoadSberReport = false;

                    txtUploadDirSberReport = txtPathBase + "EDO\\to_osp\\sber";
                    if (InPaths.Count > 25)
                    {
                        txtUploadDirSberReport = InPaths[25].ToString();
                    }

                    // 20160506 - bAutoLoadVU = true; - теперь везде грузим в локалбную базу
                    bAutoLoadVU = true;  // загрузка вод. удостоверений
                    if (InPaths.Count > 26)
                    {
                        bAutoLoadVU = Convert.ToBoolean(InPaths[26].ToString());
                    }
                    //20190218 - отключаем т.к по ВУ теперь Фед. МВВ и не нужно ничего загружать
                    bAutoLoadVU = false;  // загрузка вод. удостоверений

                    txtUploadDirVU = txtPathBase + "mvd_vu";
                    if (InPaths.Count > 27)
                    {
                        txtUploadDirVU = InPaths[27].ToString();
                    }

                    bRunIC_MVD_in = true;
                    if (InPaths.Count > 28)
                    {
                        bRunIC_MVD_in = Convert.ToBoolean(InPaths[28].ToString());
                    }

                    txtIC_MVD_Path_in = txtPathBase + "ответ\\IC_MVD";
                    if (InPaths.Count > 29)
                    {
                        txtIC_MVD_Path_in = InPaths[29].ToString();
                    }

                    bRunGibdd10_selfrequest = true;
                    if (InPaths.Count > 30)
                    {
                        bRunGibdd10_selfrequest = Convert.ToBoolean(InPaths[30].ToString());
                    }
                                       
                    // bRunGibdd10_selfrequest = false; // 20210422 - отмена обработки запросов о ВУ в ГИБДД

                    // txtGibdd10ConString  = "Provider=LCPI.IBProvider.3;Data Source=rdb-test/3051;Persist Security Info=True;Password=masterkey;User ID=SYSDBA;Location=rdb-test/3051:/opt/pkosp/db/ufssprk-tools.rdb;ctype=win1251";
                    // 20160506 - теперь запросы будут делаться в локальные базы
                    txtGibdd10ConString = constrGIBDD;
                    if (InPaths.Count > 31)
                    {
                        txtGibdd10ConString = InPaths[31].ToString();
                    }
                    
                    txtGibdd10DataBase = "MVD_VOD_PR";
                    if (InPaths.Count > 32)
                    {
                        txtGibdd10DataBase = InPaths[32].ToString();
                    }

                    bAutoLoadGIMS = true;  // загрузка вод. удостоверений
                    if (InPaths.Count > 33)
                    {
                        bAutoLoadGIMS = Convert.ToBoolean(InPaths[33].ToString());
                    }

                    // 20200208  - все отключаем кроме ic_mvd и квитанция на постановления рег. ПФР
                    bAutoLoadGIMS = false;

                    txtUploadDirGims = txtPathBase + "mvd_vu";
                    if (InPaths.Count > 34)
                    {
                        txtUploadDirGims = InPaths[34].ToString();
                    }

                    bRunGims10_selfrequest = true;
                    if (InPaths.Count > 35)
                    {
                        bRunGims10_selfrequest = Convert.ToBoolean(InPaths[35].ToString());
                    }
                    bRunGims10_selfrequest = false;  // 20210422 - отмена загрузки ответов на запросы о вод. удостоверений ГИМС

                    txtUploadDirTGK1 = txtPathBase + "ответ\\TGK-1";
                    if (InPaths.Count > 36)
                    {
                        txtUploadDirTGK1 = InPaths[36].ToString();
                    }
                    // bAutoLoadTGK1
                    bAutoLoadTGK1 = true;
                    if (iDiv == 0)
                    {
                        bAutoLoadTGK1 = true; // загрузка 
                        txtUploadDirTGK1 = "\\\\fs\\inf_obmen_rozjsk\\ответ\\TGK-1";
                    }
                    bAutoLoadTGK1 = false; // отключено т.к. нет больше обмена с ТГК-1

                    // предлагаю значение по-умолчанию прописывать в зависимости от района
                    // так будет проще - меньше конфигов писать
                    if (InPaths.Count > 37)
                    {
                        bAutoLoadTGK1 = Convert.ToBoolean(InPaths[37].ToString());
                    }
                    // 20200208  - все отключаем кроме ic_mvd и квитанция на постановления рег. ПФР
                    bAutoLoadTGK1 = false;

                    // новые параметры для автоматич выгрузки запросов в Кред.Орг
                    txtCredOrgReq_Path_out = txtPathBase + "запрос\\cred_org\\";
                    if (InPaths.Count > 38)
                    {
                        txtCredOrgReq_Path_out = InPaths[38].ToString();
                    }

                    bRunCredOrgReqOut = true;
                    if (iDiv == 0) bRunCredOrgReqOut = false;
                    if (InPaths.Count > 39)
                    {
                        bRunCredOrgReqOut = Convert.ToBoolean(InPaths[39].ToString());
                    }
                    bRunCredOrgReqOut = false; // в феврале 2019 вывели Связь-Банк из обмена т.к. перешли на ФЕД.Мвв, больше нет Рег.МВВ банков

                    // 86200999999007 - БаренцБанк
                    // 86200999999009 - СвязьБанк
                    //string txtCistomLegalIds = "86200999999007, 86200999999009";
                    string txtCistomLegalIds = "86200999999009";
                    if (InPaths.Count > 40)
                    {
                        txtCistomLegalIds = InPaths[40].ToString().Trim();
                        //if (txtCistomLegalIds.Trim().Length.Equals(0)) txtCistomLegalIds = "86200999999007, 86200999999009";
                        if (txtCistomLegalIds.Trim().Length.Equals(0)) txtCistomLegalIds = "86200999999009";
                    }

                    bRunCredOrgAnsIn = true;
                    if (iDiv == 0) bRunCredOrgAnsIn = false;
                    if (InPaths.Count > 41)
                    {
                        bRunCredOrgAnsIn = Convert.ToBoolean(InPaths[41].ToString());
                    }
                    bRunCredOrgAnsIn = false; // в феврале 2019 вывели Связь-Банк из обмена т.к. перешли на ФЕД.Мвв, больше нет Рег.МВВ банков;

                    txtPotdOutPath = txtPathBase + "запрос\\potd\\";
                    if (InPaths.Count > 42)
                    {
                        txtPotdOutPath = InPaths[42].ToString();
                    }
                    bRunPotdOut2017 = false; // 20171019 - больше не требуется выгружать
                    bRunPotdOut = true;
                    if (iDiv == 0) bRunPotdOut = false;
                    if (iDiv == 0) bRunPotdOut2017 = false;
                    
                    if (InPaths.Count > 43)
                    {
                        bRunPotdOut = Convert.ToBoolean(InPaths[43].ToString());
                    }

                    // 20200208  - все отключаем кроме ic_mvd и квитанция на постановления рег. ПФР
                    bRunPotdOut = false;
                                        
                    txtPotdInPath = txtPathBase + "ответ\\Пенсионный_фонд_расширенный\\auto2\\";
                    if (InPaths.Count > 44)
                    {
                        txtPotdInPath = InPaths[44].ToString();
                    }

                    bRunPotdIn2017 = false; // 20171103 больше не загружаем

                    bRunPotdIn = true;
                    if (iDiv == 0) bRunPotdIn = false;
                    if (InPaths.Count > 45)
                    {
                        bRunPotdIn = Convert.ToBoolean(InPaths[45].ToString());
                    }

                    // 20200208  - все отключаем кроме ic_mvd и квитанция на постановления рег. ПФР
                    bRunPotdIn = false;

                    bRunRtn10_selfrequest = true;
                    if (iDiv == 0) bRunRtn10_selfrequest = false;
                    if (InPaths.Count > 46)
                    {
                        bRunRtn10_selfrequest = Convert.ToBoolean(InPaths[46].ToString());
                    }
                    bRunRtn10_selfrequest = false; // 20210422 - отмена обработки запросов в РТН

                    bRunGimsLodka10_selfrequest = true;
                    if (iDiv == 0) bRunGimsLodka10_selfrequest = false;
                    if (InPaths.Count > 47)
                    {
                        bRunGimsLodka10_selfrequest = Convert.ToBoolean(InPaths[47].ToString());
                    }
                    bRunGimsLodka10_selfrequest = false; // 20210422 - отмена обработки запросов в ГИМС о лодках


                    txtUploadDirRicZH = txtPathBase + "ответ\\РИЦ_ЖХ";
                    if (InPaths.Count > 48)
                    {
                        txtUploadDirRicZH = InPaths[48].ToString().Trim();
                    }

                    bAutoLoadRicZH = false;
                    switch (iDiv)
                    {
                        case 1: bAutoLoadRicZH = true; break;
                        case 13: bAutoLoadRicZH = true; break;
                        case 20: bAutoLoadRicZH = true; break;
                        case 21: bAutoLoadRicZH = true; break;
                        case 24: bAutoLoadRicZH = true; break;
                        default: bAutoLoadRicZH = false; break;
                    }
                    bAutoLoadRicZH = false; // отключено т.к. нет больше обмена с РИЦ ЖХ
                    
                    if (InPaths.Count > 49)
                    {
                        bAutoLoadRicZH = Convert.ToBoolean(InPaths[49].ToString().Trim());
                    }

                    // 20200208  - все отключаем кроме ic_mvd и квитанция на постановления рег. ПФР
                    bAutoLoadRicZH = false;


                    // txtGibdd10ConString  = "Provider=LCPI.IBProvider.3;Data Source=rdb-test/3051;Persist Security Info=True;Password=masterkey;User ID=SYSDBA;Location=rdb-test/3051:/opt/pkosp/db/ufssprk-tools.rdb;ctype=win1251";
                    // 20160506 - теперь запросы будут делаться в локальные базы
                    txtConStrED = "Provider=LCPI.IBProvider.3;Data Source=ed10/3050;Persist Security Info=True;Password=ksvptz;User ID=SYSDBA;Location=10.10.4.243/3050:ncore-fssp;ctype=win1251";
                    if (InPaths.Count > 50)
                    {
                        txtConStrED = InPaths[50].ToString();
                        txtConStrED = ff.RemoveDomainFromConString(txtConStrED);
                    }

                    txtPfrRepIn = txtPathBase + "ответ\\Пенсионный_фонд\\report\\";
                    if (InPaths.Count > 51)
                    {
                        txtPfrRepIn = InPaths[51].ToString();
                    }

                    bRunPfrRepIn = true; // 20190910 - включил

                    if (iDiv == 0) bRunPfrRepIn = false;
                    if (InPaths.Count > 52)
                    {
                        bRunPfrRepIn = Convert.ToBoolean(InPaths[52].ToString());
                    }

                    // обработка запросов в ИЦ МВД о паспорте (ФМС)
                    bRunIcmvdFms10_selfrequest = true;
                    if (iDiv == 0) bRunRtn10_selfrequest = false;
                    if (InPaths.Count > 53)
                    {
                        bRunRtn10_selfrequest = Convert.ToBoolean(InPaths[53].ToString());
                    }
                    // bRunRtn10_selfrequest = false; 

                    string txtAgreementCode = "0";
                    int nPack_type = 0;
                    
                    
                    
                    // ПРОВЕРКА ЧТО ПОЛУЧЕН КОД осп БЕЗ ОШИБОК
                    if (txtErr.Length > 0 && iDiv == 0)
                    {
                        Exception ex = new Exception("Error in GetOSP_Num." + txtErr);
                        throw ex;
                    }
                    else
                    {
                        // Logger для всего подряд - проходной
                        // up.my_version
                        Logger_ufssprk_tools lCommonLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 0, "SverkaGibdd", 0, nOspNum, "Общий лог запуска SverkaGibdd v." + up.GetVersion());
                        if (lCommonLogger.ValidCon)
                        {
                            lCommonLogger.UpdateLLogFileName(up.GetVersion()); // версию ПО записываю в поле FILENAME

                            // провести автозагрузку
                            #region AutoLoadGibdd
                            if (bAutoLoadGibdd)
                            {
                                //INSERT INTO AGREEMENTS (ID, AGREEMENT_CODE, NAME_AGREEMENT, DESCRIPTION) VALUES (250, '250', 'МВД Карелии (реестры)', NULL);
                                txtAgreementCode = "250";

                                nPack_type = 6;
                                // INSERT INTO PACK_TYPE (ID, TYPE) VALUES (6, 'Загрузка реестра оплаченных штрафов МВД');

                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка реестров оплаченных штрафов МВД.");
                                lLogger.ErrMessage += txtErr;

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);

                                if (lLogger.ErrMessage.Length > 0)
                                {
                                    // SendEmail(lLogger.ErrMessage, "Ошибки при загрузке реестров оплаченных штрафов МВД в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");
                                    ff.SendEmail(lLogger.ErrMessage, "Внимание! Ошибка загрузки реестра оплаченных штрафов МВД в ОСП " + lLogger.OspNum.ToString().PadLeft(2, '0'), txtAdminEmail, txtAdminEmail, txtMailServ, "");
                                }

                                ShowLoggerError(lLogger);

                                Console.WriteLine("Итого загружено реестров МВД: " + iTotalFiles.ToString());

                                //OleDbConnection conPK_OSP = new OleDbConnection(txtConStrPKOSP);
                                //int iDiv = Convert.ToInt32(GetOSP_Num(conPK_OSP));
                                //string txtOspEmail = "";
                                //string txtAdminEmail = "";

                                //if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                //if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                //else txtOspEmail = txtAdminEmail;



                                //AutoLoadGibddReestrs(constrGIBDD, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail);
                            }
#endregion
                            // провести автозагрузку по КРЦ
                            #region AutoLoadKRC
                            if (bAutoLoadKRC)
                            {
                                //INSERT INTO AGREEMENTS (ID, AGREEMENT_CODE, NAME_AGREEMENT, DESCRIPTION) VALUES (260, '260', 'КРЦ (реестры)', NULL);
                                txtAgreementCode = "260";

                                nPack_type = 6;
                                // INSERT INTO PACK_TYPE (ID, TYPE) VALUES (6, 'Загрузка реестра оплаченных штрафов МВД');

                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка реестров оплат из КРЦ.");
                                lLogger.ErrMessage += txtErr;

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                iTotalFiles = mvv.AutoLoadKRCReestrs(constrGIBDD, txtUploadDirKRC, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);

                                ShowLoggerError(lLogger);

                                Console.WriteLine("Итого загружено реестров КРЦ: " + iTotalFiles.ToString());

                                //OleDbConnection conPK_OSP = new OleDbConnection(txtConStrPKOSP);
                                //int iDiv = Convert.ToInt32(GetOSP_Num(conPK_OSP));
                                //string txtOspEmail = "";
                                //string txtAdminEmail = "";

                                //if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                //if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                //else txtOspEmail = txtAdminEmail;



                                //AutoLoadGibddReestrs(constrGIBDD, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail);
                            }

#endregion

                            #region AutoLoadRicZH
                            if (bAutoLoadRicZH)
                            {
                                //INSERT INTO AGREEMENTS (ID, AGREEMENT_CODE, NAME_AGREEMENT, DESCRIPTION) VALUES (260, '260', 'КРЦ (реестры)', NULL);
                                txtAgreementCode = "310";

                                nPack_type = 6;
                                // INSERT INTO PACK_TYPE (ID, TYPE) VALUES (6, 'Загрузка реестра оплаченных штрафов МВД');

                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка реестров оплат из РИЦ ЖХ.");
                                lLogger.ErrMessage += txtErr;

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                // iTotalFiles = mvv.AutoLoadKRCReestrs(constrGIBDD, txtUploadDirKRC, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                iTotalFiles = mvv.AutoLoadRicZHReestrs(constrGIBDD, txtUploadDirRicZH, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                lLogger.UpdateLLogCount(iTotalFiles);
                                lLogger.WriteLLog("Итого загружено реестров РИЦ ЖХ: " + iTotalFiles.ToString());

                                ShowLoggerError(lLogger);

                                Console.WriteLine("Итого загружено реестров РИЦ ЖХ: " + iTotalFiles.ToString());

                                //OleDbConnection conPK_OSP = new OleDbConnection(txtConStrPKOSP);
                                //int iDiv = Convert.ToInt32(GetOSP_Num(conPK_OSP));
                                //string txtOspEmail = "";
                                //string txtAdminEmail = "";

                                //if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                //if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                //else txtOspEmail = txtAdminEmail;



                                //AutoLoadGibddReestrs(constrGIBDD, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail);
                            }
#endregion

                            #region AutoSberDocsOut
                            if (bAutoSberDocsOut) // это выгрузка постановлений в Сбербанк (.1ss файлы)
                            {
                                // 20160721 - в связи с внедрение меняю расписание выгрузки
                                // т.к. задача планировщика запускается ежедневно, но выгружать оператору удобно только по рабочим дням
                                // поэтому делаем выгрузку только по рабочим дням, а если будет суббота рабочая,
                                // то просто подождем то след. рабочего дня
                                // день недели от 1 до 7
                                int day = ((int)DateTime.Now.DayOfWeek == 0) ? 7 : (int)DateTime.Now.DayOfWeek;
                                if ((day > 0) && (day < 6)) // выгрузка в субб (6) и воскр(7) не работает
                                {

                                    Int64 cnt = 0;

                                    int iSplitLimitSber = 2500;
                                    int nRowsCount = 0;

                                    DataTable dtRestrictionsSber = null;
                                    DataTable dtRestrictionsSberNoOldIP = null;
                                    DataTable dtRestrictionsSberOldIp = null;
                                    DataTable dtFedMvvSberEndGaccount = null;
                                    //string txtAgreementSber = "СБЕР_ЭДО_10";
                                    txtAgreementCode = "СБЕР_ЭДО_10";
                                    nPack_type = 9;

                                    Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Пакет постановлений в сбербанк (*.1SS формат).");

                                    // получить строки отсортированные c пустым old_regnumber
                                    
                                    // 20191211 - отключаю на веремя тестов
                                    // 20200114 - вернул обратно
                                    dtRestrictionsSber = mvv.ReadSberRestrictions(conPK_OSP, lLogger);
                                    

                                    // получить строки  С НЕПУСТЫМ old_regnumber отсортированные по-убыванию

                                    // 20191211 - отключаю на веремя тестов
                                    // 20200114 - вернул обратно
                                    dtRestrictionsSberOldIp = mvv.ReadSberOldIpRestrictions(conPK_OSP, lLogger);
                                    

                                    // получить отмены фед.МВВ которые не ушли
                                    // 20191227 - отключаем т.к. возникли проблемы - почему-то повторно выгружается
                                    // наверное update не работает
                                    // 20200114 - вернул обратно
                                    dtFedMvvSberEndGaccount = mvv.ReadFedMvvSberEndGaccount(conPK_OSP, lLogger);
                                    
                                    // меняем подход - теперь выбираем все сразу, но валим вариант с xss просто рядом
                                    //dtRestrictionsSberNoOldIP = mvv.ReadSberRestrictionsNoOldNumber(conPK_OSP, lLogger);
                                    // dtRestrictionsSberOldIp = mvv.ReadSberRestrictionsOldNumber(conPK_OSP, lLogger);

                                    ShowLoggerError(lLogger);

                                    if (dtRestrictionsSber != null)
                                    {
                                        nRowsCount += dtRestrictionsSber.Rows.Count;
                                    }

                                    // Сформировать папку и имя файла для выгрузки
                                    // TODO: научиьтся по логам узнавать номер пакета по счету в сегодня
                                    // 1 - завести свой тип пакета (постановления СПИ)
                                    // 2 - если пакет не был успешно отправлен, то его статус не должен считаться
                                    int nFileNum = 1;
                                    nFileNum = mvv.GetDayPackCount(constrGIBDD, nPack_type, txtAgreementCode, DateTime.Today, lLogger) + 1;
                                    ShowLoggerError(lLogger);

                                    int nRestrSberCount = 0;
                                    if(dtRestrictionsSber != null && dtRestrictionsSber.Rows.Count > 0) nRestrSberCount = dtRestrictionsSber.Rows.Count;

                                    int col_files = nRestrSberCount / iSplitLimitSber;
                                    bool isFreeFileNum = true; // флаг - осталось место для выгрузки файла со старыми номерами ИП
                                    bool isFreeFileNum2 = true; // флаг - осталось место для выгрузки файла со старыми номерами ИП

                                    if (nFileNum + col_files > 35)
                                    {
                                        isFreeFileNum = false; // нет свободного имени файла на сегодня
                                        // выгружать будем не более 35 пакетов в день (+1)
                                        col_files = 35 - nFileNum;
                                        // MessageBox.Show("Подготовлено к выгрузке " + dtRestrictionsSber.Rows.Count.ToString() + " строк со счетами.\n По условиям соглашения со Сбербанком за один день в файл нельзя выгрузить более чем " + (iSplitLimitSber * (col_files + 1)).ToString() + " строк со счетами в постановлениях СПИ. Оставшиеся строки будут выгружены в следующие дни.", "Внимание!", MessageBoxButtons.OK);
                                        lLogger.MemoryLLog(" \nПодготовлено к выгрузке " + nRowsCount.ToString() + " строк со счетами.");
                                        if (nRowsCount > (iSplitLimitSber * (col_files + 1)))
                                        {
                                            lLogger.MemoryLLog(" \nПо условиям соглашения со Сбербанком за один день в файл нельзя выгрузить более чем " + (iSplitLimitSber * (col_files + 1)).ToString() + " строк со счетами в постановлениях СПИ. Оставшиеся строки будут выгружены в следующие дни.");
                                        }
                                    }

                                    if (nFileNum + col_files > 34)
                                    {
                                        isFreeFileNum2 = false;
                                    }

                                    // зафиксировать папку на этом верхнем уровне и далее ее не менять
                                    // папка для выгрузки?
                                    //sber_path =  sber_doc_out
                                    string txtSberPathWithDate = ff.CreatePathWithDateS(txtSberDocsPath);

                                    Dictionary<string, int> dictDocsCount = new Dictionary<string, int>();
                                    Dictionary<string, int> dictActIDsCount = new Dictionary<string, int>();
                                    string txtCommonUserLog = "";

                                    bool bUnloadLastRegNumber = false;
                                    bool bUnloadFedMvvSberEndGaccount = false;
                                    // делаем тестовую выгрузку 
                                    // bool bUnloadLastRegNumber = true;


                                    OspOptions paramOsp = mvv.GetOspOptions(conPK_OSP, lLogger);

                                    //for (int i = nFileNum; i <= nFileNum + col_files; i++)
                                    for (int i = 0; i <= col_files; i++)
                                    {

                                        string txtSberOutFileName = ff.makenewSberFileName2(nOspNum) + '.' + ff.fileCode(i + nFileNum) + "SS"; // N;
                                        cnt += mvv.WriteSberDocsToTxt(constrGIBDD, txtConStrPKOSP, txtSberPathWithDate, txtSberOutFileName, dtRestrictionsSber, txtAgreementCode, i, iSplitLimitSber, ref dictDocsCount, ref dictActIDsCount, ref txtCommonUserLog, nFileNum, bUnloadLastRegNumber, paramOsp, lLogger);
                                    }



                                    // если есть строки с непрошедшими отменами фед.МВВ
                                    if (dtFedMvvSberEndGaccount != null && dtFedMvvSberEndGaccount.Rows.Count > 0 && isFreeFileNum2)
                                    {

                                        bUnloadFedMvvSberEndGaccount = true;
                                        bUnloadLastRegNumber = false;

                                        // добавляем 1, т.к. это в любом случае новый файл
                                        string txtSberOutFileName = ff.makenewSberFileName2(nOspNum) + ".YSS"; // N;

                                        // ограничение выставляем равное кол-ву строк в выборке со старыми номерами ИП
                                        // порядковый номер файла делаем 0, чтобы верно считалось количество строк внутри
                                        
                                        //cnt += mvv.WriteSberDocsToTxt(constrGIBDD, txtConStrPKOSP, txtSberPathWithDate, txtSberOutFileName, dtFedMvvSberEndGaccount, txtAgreementCode, 0, dtRestrictionsSberOldIp.Rows.Count, ref dictDocsCount, ref dictActIDsCount, ref txtCommonUserLog, nFileNum, bUnloadLastRegNumber, paramOsp, lLogger);
                                        cnt += mvv.WriteSberDocsToTxtFedMvv(constrGIBDD, txtConStrPKOSP, txtSberPathWithDate, txtSberOutFileName, dtFedMvvSberEndGaccount, txtAgreementCode, 0, dtFedMvvSberEndGaccount.Rows.Count, ref dictDocsCount, ref dictActIDsCount, ref txtCommonUserLog, 25, false, paramOsp, lLogger);
                                        
                                        
                                    }



                                    // если есть строки со старыми номерами ИП и есть свободное имя файла для них
                                    if (dtRestrictionsSberOldIp != null && dtRestrictionsSberOldIp.Rows.Count > 0 && isFreeFileNum)
                                    {

                                        // а теперь выгружаем дополнительно еще отмены со старыми ИП в файлы с номерами, идущими после предыдщих
                                        // на файл со старыми номерами никаких ограничений не делаю - все буду выгружать в одном файле
                                        // поэтому ничего проверять не буду, кроме того - есть еще резерв для одного файла или нет

                                        bUnloadLastRegNumber = true;

                                        // добавляем 1, т.к. это в любом случае новый файл
                                        string txtSberOutFileName = ff.makenewSberFileName2(nOspNum) + '.' + ff.fileCode(col_files + 1 + nFileNum) + "SS"; // N;

                                        // ограничение выставляем равное кол-ву строк в выборке со старыми номерами ИП
                                        // порядковый номер файла делаем 0, чтобы верно считалось количество строк внутри
                                        cnt += mvv.WriteSberDocsToTxt(constrGIBDD, txtConStrPKOSP, txtSberPathWithDate, txtSberOutFileName, dtRestrictionsSberOldIp, txtAgreementCode, 0, dtRestrictionsSberOldIp.Rows.Count, ref dictDocsCount, ref dictActIDsCount, ref txtCommonUserLog, nFileNum, bUnloadLastRegNumber, paramOsp, lLogger);

                                        //20180220 - делаем файл-сопровод
                                        // копируем с нужным именем файла
                                        char cDelim = '\\';
                                        // !!! Временно - убрать после тестов
                                        // txtPathBase = "\\\\fs\\inf_obmen_rozjsk";

                                        string txtSoprPath = string.Format(@"{0}" + cDelim + "{1}", txtPathBase, "sys");
                                        string txtSoprFileName = "document.xml";

                                        //if(File.Exists(string.Format(@"{0}" + cDelim + "{1}", txtSoprPath, txtSoprFileName)))
                                        //File.Copy(string.Format(@"{0}" + cDelim + "{1}", txtSoprPath, txtSoprFileName), string.Format(@"{0}" + cDelim + "{1}", txtSberPathWithDate, txtSberOutFileName + ".xml"));

                                        // заменить в Xml файле 1001202.2 заменить на подстроку имени файла без последних 2-х SS
                                        // Добавить в string.Format(@"{0}" + cDelim + "{1}", txtSberPathWithDate, txtSberOutFileName + ".docx")
                                        string txtToFind = "1001202.2";
                                        string txtToReplace = "";
                                        if (txtSberOutFileName.Length > 2) txtToReplace = txtSberOutFileName.Substring(0, txtSberOutFileName.Length - 2);
                                        ff.FileStringReplace(string.Format(@"{0}" + cDelim + "{1}", txtSoprPath, txtSoprFileName), string.Format(@"{0}" + cDelim + "{1}", txtSberPathWithDate, txtSberOutFileName + ".xml"), txtToFind, txtToReplace, Encoding.UTF8);
                                    }

                                    string txtEmailSubj = "Сформирован пакет электронных постановлений в Сбербанк в ОСП " + lLogger.OspNum.ToString().PadLeft(2, '0');
                                    string txtOspEmailMess = "";

                                    // вывести лог
                                    // в логе указать сначала что выгружено удачно, потом что неудачно
                                    lLogger.MemoryLLog("\nЖурнал выгрузки постановлений СПИ от " + DateTime.Today.ToShortDateString());
                                    txtOspEmailMess += "\nЖурнал выгрузки постановлений СПИ от " + DateTime.Today.ToShortDateString();


                                    lLogger.MemoryLLog("\nИтого во всех пакетах выгружено:");
                                    txtOspEmailMess += "\nИтого во всех пакетах выгружено:";

                                    int iDocsCount = 0;
                                    foreach (string key in dictDocsCount.Keys)
                                    {
                                        lLogger.MemoryLLog("\n" + mvv.GetDocName(key) + ": ");
                                        txtOspEmailMess += "\n" + mvv.GetDocName(key) + ": ";
                                        lLogger.MemoryLLog("\n" + dictDocsCount[key].ToString() + ";");
                                        txtOspEmailMess += "\n" + dictDocsCount[key].ToString() + ";";
                                        iDocsCount += dictDocsCount[key];
                                    }
                                    lLogger.MemoryLLog("\nИтого счетов в постановлениях:" + cnt.ToString());
                                    txtOspEmailMess += "\nИтого счетов в постановлениях:" + cnt.ToString();
                                    lLogger.MemoryLLog("\nИтого постановлений:" + iDocsCount.ToString());
                                    txtOspEmailMess += "\nИтого постановлений:" + iDocsCount.ToString();

                                    // 20160708 - добавляю поиск файла zss на rdb сервере ОСП и копирование его в папку куда уже выгружен 1ss-файл
                                    // 0 - работать будет, но zss каждый день формируются, а это как-то не очень..
                                    // поэтому тут вариант или планировщик крутить - только с пн по пт. или выгружать автоматом ежедневно
                                    // что лучше?

                                    // 1. из connString получить адрес сервера
                                    string txtFsServerName = ff.GetServerNameFromConStr(txtConStrPKOSP);

                                    // 2. проверить есть ли файл zss на rdb сервере

                                    // как сформировать строку до zss файла?
                                    // txtSberPathWithDate - путь до файла
                                    // 
                                    string txtZssFileName = ff.makenewSberFileName2(nOspNum) + '.' + "zss"; // N;
                                    if (txtFsServerName.Length > 0)
                                    {
                                        // добавить путь к fs - \\ в начале и путь до выгрузки в конце
                                        txtFsServerName = "\\\\" + txtFsServerName;
                                        txtFsServerName += "\\pksp\\sber\\docs";
                                        if (File.Exists(string.Format(@"{0}\{1}", txtFsServerName, txtZssFileName)))
                                        {
                                            // 3. если есть - то копировать в папку куда уже положили 1ss файл
                                            File.Copy(string.Format(@"{0}\{1}", txtFsServerName, txtZssFileName), string.Format(@"{0}\{1}", txtSberPathWithDate, txtZssFileName));
                                            //  тут было бы неплохо написать что не нужно ничего выгружать из ПК ОСП
                                            // MessageBox.Show("По пути " + string.Format(@"{0}\{1}", txtSberPathWithDate, txtZssFileName) + " автоматически размещен файл .ZSS с электронными постановлениями в формате ПИЭВ." , "Внимание", MessageBoxButtons.OK);
                                            // по идее неплохо бы отправить в почту что выгрузили 1ss и к нему скопировали zss
                                            // пока не будем - долго тестировать
                                            txtOspEmailMess += "\nПо пути " + string.Format(@"{0}\{1}", txtSberPathWithDate, txtZssFileName) + " автоматически размещен файл .ZSS с электронными постановлениями в формате ПИЭВ.";
                                            //SendEmail(txtCommonUserLog, "Ошибки при выгрузке постановлений в Сбербанк в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                        }

                                    }
                                    // Отправеить e-mail в ОСП
                                    string txtOspEmail = "";
                                    string txtAdminEmail = "";

                                    if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];
                                    //txtAdminEmail += ";nadezhda.smirnova1@r10.fssprus.ru";

                                    if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                    else txtOspEmail = txtAdminEmail;

                                    SendEmail(txtOspEmailMess, txtEmailSubj, txtOspEmail, "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                    // Отправеить e-mail админу
                                    SendEmail(txtOspEmailMess, txtEmailSubj, txtAdminEmail, "mvv_report@r10.fssprus.ru", txtMailServ, "");
                                    SendEmail(txtOspEmailMess, txtEmailSubj, "nadezhda.smirnova1@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");



                                    if (lLogger.ErrMessage.Length > 0)
                                    {
                                        // если есть лог с ошибками - отправить соответствующий отчет
                                        //SendEmail("\nСообщения об ошибках:\n " + lLogger.ErrMessage, "Ошибки при выгрузке постановлений в Сбербанк в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru;nadezhda.smirnova1@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");
                                        SendEmail("\nСообщения об ошибках:\n " + lLogger.ErrMessage, "Ошибки при выгрузке постановлений в Сбербанк в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");
                                        SendEmail("\nСообщения об ошибках:\n " + lLogger.ErrMessage, "Ошибки при выгрузке постановлений в Сбербанк в ОСП " + lLogger.OspNum.ToString(), "nadezhda.smirnova1@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                        lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                        lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                        ShowLoggerError(lLogger);
                                    }

                                    Console.WriteLine("Выгружено постановлений СПИ в Сбербанк: " + iDocsCount.ToString());
                                }

                            }
#endregion


                            #region AutoSberReqOut
                            if (bAutoSberReqOut)
                            {
                                Int64 cnt = 0;

                                DataTable DT_doc_fiz_sber = null;
                                string txtAgreementSber = "10";

                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 1, txtAgreementSber, 0, nOspNum, "Пакет запросов в сбербанк (новый XML формат).");
                                DT_doc_fiz_sber = mvv.ReadSberZaprosNewFormat(conPK_OSP, lLogger);
                                // DT_doc_fiz_sber = mvv.ReadSberZaprosNewFormatTest(conPK_OSP, lLogger);
                                ShowLoggerError(lLogger);

                                // записать XML Сбербанка

                                // TODO: посмотреть по логам какой по счету файл в Сбербанк за сегодня мы выгружаем
                                int nFileNum = 0;
                                //// packType 1 - обычный запрос
                                nFileNum = mvv.GetDayPackCount(constrGIBDD, 1, txtAgreementSber, DateTime.Today, lLogger); //  +1;  - тут нумерация с 0
                                ShowLoggerError(lLogger);

                                string txtSberXmlFileName = ff.makenewSberFileName() + '.' + ff.makenewSberFileExt(nFileNum, nOspNum);
                                // путь для файла-запроса в Сбер - путь из настроек + дата за сегодня
                                string txtSberFolder = ff.CreatePathWithDateS(txtSberReqOutPath);
                                cnt = mvv.WriteToXML(constrGIBDD, txtConStrPKOSP, txtSberFolder, txtSberXmlFileName, DT_doc_fiz_sber, nOspNum, lLogger);

                                ShowLoggerError(lLogger);

                            }

                            if (bAutoSberRespIn)
                            {
                                txtAgreementCode = "10";
                                nPack_type = 2; // ответ на запрос


                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка ответов XML из Сбербанка.");

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                //iTotalFiles = mvv.AutoLoadSberKvit(constrGIBDD, txtConStrPKOSP, txtUploadDirSberKvit, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                iTotalFiles = mvv.AutoLoadSberRespIn(constrGIBDD, txtConStrPKOSP, txtSberRespInPath, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);

                                ShowLoggerError(lLogger);

                                Console.WriteLine("Итого загружено файлов с квитанциями Сбербанка: " + iTotalFiles.ToString());
                            }
#endregion

                            #region AutoLoadSberKvit
                            // провести автозагрузку квитанций Сбербанка
                            if (bAutoLoadSberKvit)
                            {
                                txtAgreementCode = "СБЕР_ЭДО_10";
                                nPack_type = 10; // Загрузка уведомлений о принятии в обработку из Сбербанка
                                int nlogPack_type = 6; // общий лог автозагрузки идет как загрузка реестров МВД


                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nlogPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка квитанций из Сбербанка.");
                                lLogger.ErrMessage += txtErr;

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                iTotalFiles = mvv.AutoLoadSberKvit(constrGIBDD, txtConStrPKOSP, txtUploadDirSberKvit, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                lLogger.UpdateLLogCount(iTotalFiles);
                                // ShowLoggerError(lLogger);
                                if (lLogger.ErrMessage.Length > 0)
                                {
                                    // если есть лог с ошибками - отправить соответствующий отчет
                                    SendEmail(lLogger.ErrMessage, "Ошибки при загрузке квитанций из Сбербанка в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                    lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                    lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                    ShowLoggerError(lLogger);
                                }

                                Console.WriteLine("Итого загружено файлов с квитанциями Сбербанка: " + iTotalFiles.ToString());
                            }
#endregion

                            #region SberReportLoad
                            // провести автозагрузку квитанций Сбербанка
                            if (bAutoLoadSberReport)
                            {
                                txtAgreementCode = "СБЕР_ЭДО_10";
                                // Исправить nPack_type в lLogg - чтобы пакет был не кодом 13, а с кодом обычным (как общий лог авотзагрузки и т.п.)
                                // т.к. с кодом 13 будут только реальные пакеты
                                // а тут будет код 6 - тк же как с реестрами МВД и квитанциями о приеме в обработку
                                nPack_type = 6;


                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Пакет отчетов о результатах обработки постановлений в Сбербанке.");
                                lLogger.ErrMessage += txtErr;
                                //ShowLoggerError(lLogger); // вывод сообщения об ошибке, если она вдруг случилась внутри вызванной функции

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                iTotalFiles = mvv.AutoLoadSberReport(constrGIBDD, txtConStrPKOSP, txtUploadDirSberReport, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                lLogger.UpdateLLogCount(iTotalFiles);
                                //lLogger.UpdateLLogStatus

                                // ShowLoggerError(lLogger);
                                if (lLogger.ErrMessage.Length > 0)
                                {
                                    // если есть лог с ошибками - отправить соответствующий отчет
                                    SendEmail(lLogger.ErrMessage, "Ошибки при загрузке отчетов из Сбербанка в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                    lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                    lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                    ShowLoggerError(lLogger);
                                }

                                Console.WriteLine("Итого загружено файлов с отчетами о результатах обработки из Сбербанка: " + iTotalFiles.ToString());
                            }

                            #endregion

                            #region GIMS_Load
                            // провести автозагрузку реестров прав управления на маломерные суда, выданные ГИМС
                            if (bAutoLoadGIMS)
                            {
                                txtAgreementCode = "250"; // мвд
                                nPack_type = 14; // INSERT INTO PACK_TYPE (ID, TYPE)  VALUES (14, 'Загрузка реестров сведени о вод. удостоверении ');


                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрука сведений о прав. управления на маломерные суда, выданные ГИМС.");
                                lLogger.ErrMessage += txtErr;
                                //ShowLoggerError(lLogger); // вывод сообщения об ошибке, если она вдруг случилась внутри вызванной функции

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                OleDbConnection conGIBDD = new OleDbConnection(constrGIBDD);
                                iTotalFiles = mvv.AutoLoadGIMS(constrGIBDD, txtConStrPKOSP, txtUploadDirGims, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                // AutoLoadVU(constrGIBDD, txtConStrPKOSP, txtUploadDirVU, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                //iTotalFiles = mvv.AutoLoadSberReport(constrGIBDD, txtConStrPKOSP, txtUploadDirSberReport, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);

                                lLogger.UpdateLLogCount(iTotalFiles);
                                //lLogger.UpdateLLogStatus

                                // ShowLoggerError(lLogger);
                                if (lLogger.ErrMessage.Length > 0)
                                {
                                    // если есть лог с ошибками - отправить соответствующий отчет
                                    SendEmail(lLogger.ErrMessage, "Ошибки при загрузке сведений об удостоверениях от ГИМС в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                    lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                    lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                    ShowLoggerError(lLogger);
                                }

                                Console.WriteLine("Итого загружено файлов со сведениями об удостоверениях от ГИМС: " + iTotalFiles.ToString());
                            }

                            #endregion

                            #region VU_Load
                            // провести автозагрузку квитанций Сбербанка
                            if (bAutoLoadVU)
                            {
                                txtAgreementCode = "250"; // мвд
                                nPack_type = 14; // INSERT INTO PACK_TYPE (ID, TYPE)  VALUES (14, 'Загрузка реестров сведени о вод. удостоверении ');


                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрука сведений о вод. удостоверениях из МВД.");
                                lLogger.ErrMessage += txtErr;
                                //ShowLoggerError(lLogger); // вывод сообщения об ошибке, если она вдруг случилась внутри вызванной функции

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                OleDbConnection conGIBDD = new OleDbConnection(constrGIBDD);
                                iTotalFiles = mvv.AutoLoadVU(constrGIBDD, txtConStrPKOSP, txtUploadDirVU, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                //iTotalFiles = mvv.AutoLoadSberReport(constrGIBDD, txtConStrPKOSP, txtUploadDirSberReport, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);

                                lLogger.UpdateLLogCount(iTotalFiles);
                                //lLogger.UpdateLLogStatus

                                // ShowLoggerError(lLogger);
                                if (lLogger.ErrMessage.Length > 0)
                                {
                                    // если есть лог с ошибками - отправить соответствующий отчет
                                    SendEmail(lLogger.ErrMessage, "Ошибки при загрузке сведений о ВУ от ГИБДД в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                    lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                    lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                    ShowLoggerError(lLogger);
                                }

                                Console.WriteLine("Итого загружено файлов со сведениями о ВУ от ГИБДД: " + iTotalFiles.ToString());
                            }

                            #endregion

                            #region IcMvdOtvetLoad
                            // провести автозагрузку ответов из ИЦ МВД
                            // 20160331 - сделать разбивку ответа с учетом txtListID - разделитель ;
                            if (bRunIC_MVD_in)
                            {
                                txtAgreementCode = "ИЦ_МВД_10";
                                nPack_type = 2; // 2 - простой ответ


                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка ответов из ИЦ МВД (по требованиям).");
                                lLogger.ErrMessage += txtErr;
                                //ShowLoggerError(lLogger); // вывод сообщения об ошибке, если она вдруг случилась внутри вызванной функции

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                //iTotalFiles = mvv.AutoLoadSberReport(constrGIBDD, txtConStrPKOSP, txtUploadDirSberReport, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                iTotalFiles = mvv.AutoLoadIcMvdOtvet(constrGIBDD, txtConStrPKOSP, txtIC_MVD_Path_in, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, txtAgreementCode, nPack_type, lLogger);

                                lLogger.UpdateLLogCount(iTotalFiles);
                                //lLogger.UpdateLLogStatus

                                // ShowLoggerError(lLogger);
                                if (lLogger.ErrMessage.Length > 0)
                                {
                                    // если есть лог с ошибками - отправить соответствующий отчет
                                    SendEmail(lLogger.ErrMessage, "Ошибки при загрузке ответов от ИЦ МВД в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                    lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                    lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                    ShowLoggerError(lLogger);
                                }

                                Console.WriteLine("Итого загружено ответов из ИЦ МВД: " + iTotalFiles.ToString());
                            }

                            #endregion
                            
                            #region "KESK"
                            // провести автозагрузку по КРЦ
                            if (bAutoLoadKESK)
                            {
                                //INSERT INTO AGREEMENTS (ID, AGREEMENT_CODE, NAME_AGREEMENT, DESCRIPTION) VALUES (260, '260', 'КРЦ (реестры)', NULL);
                                txtAgreementCode = "ВК_КЭСК_10";

                                nPack_type = 7;
                                // Загрузка реестра оплат по ИД из РИЦ ЖХ

                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка реестров оплат из КЭСК.");
                                lLogger.ErrMessage += txtErr;

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                iTotalFiles = mvv.AutoLoadKESKReestrs(constrGIBDD, txtUploadDirKESK, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                //AutoLoadKRCReestrs(constrGIBDD, txtUploadDirKRC, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);

                                ShowLoggerError(lLogger);

                                Console.WriteLine("Итого загружено реестров КЭСК: " + iTotalFiles.ToString());

                                //OleDbConnection conPK_OSP = new OleDbConnection(txtConStrPKOSP);
                                //int iDiv = Convert.ToInt32(GetOSP_Num(conPK_OSP));
                                //string txtOspEmail = "";
                                //string txtAdminEmail = "";

                                //if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                //if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                //else txtOspEmail = txtAdminEmail;



                                //AutoLoadGibddReestrs(constrGIBDD, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail);
                            }
                            #endregion

                            #region "TGK1"
                            // провести автозагрузку по КРЦ
                            if (bAutoLoadTGK1)
                            {
                                //INSERT INTO AGREEMENTS (ID, AGREEMENT_CODE, NAME_AGREEMENT, DESCRIPTION) VALUES (260, '260', 'КРЦ (реестры)', NULL);
                                txtAgreementCode = "ВК_ТГК1_10";

                                nPack_type = 7;
                                // Загрузка реестра оплат по ИД из РИЦ ЖХ

                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка реестров оплат из ТГК1.");
                                lLogger.ErrMessage += txtErr;

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadKESKReestrs(constrGIBDD, txtUploadDirKESK, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                iTotalFiles = mvv.AutoLoadTGK1Reestrs(constrGIBDD, txtUploadDirTGK1, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);

                                ShowLoggerError(lLogger);

                                Console.WriteLine("Итого загружено реестров ТГК1: " + iTotalFiles.ToString());
                            }
                            #endregion

                            # region "GIBDD_10"
                            if (bRunGibdd10_out)
                            {
                                Int64 cnt = 0;

                                DataTable DT_doc_gibdd10 = null;
                                txtAgreementCode = "ГИБДД_10";
                                string txtAgreementSber = "ГИБДД_10";

                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 1, txtAgreementSber, 0, nOspNum, "Пакет запросов в ГИБДД_10.");
                                DT_doc_gibdd10 = mvv.ReadGibdd10Zapros(conPK_OSP, lLogger);
                                // DT_doc_fiz_sber = mvv.ReadSberZaprosNewFormatTest(conPK_OSP, lLogger);
                                ShowLoggerError(lLogger);

                                // TODO: посмотреть по логам какой по счету файл за сегодня мы выгружаем
                                int nFileNum = 0;
                                //// packType 1 - обычный запрос
                                nFileNum = mvv.GetDayPackCount(constrGIBDD, 1, txtAgreementSber, DateTime.Today, lLogger); //  +1;  - тут нумерация с 0
                                ShowLoggerError(lLogger);

                                //string txtSberXmlFileName = ff.makenewSberFileName() + '.' + ff.makenewSberFileExt(nFileNum, nOspNum);
                                // zp_yyyymmdd_ХХ.txt, где:
                                string txtFileName = "zp";
                                txtFileName += DateTime.Today.Year.ToString("D4"); // yyyy
                                txtFileName += DateTime.Today.Month.ToString("D2").PadLeft(2, '0'); // mm
                                txtFileName += DateTime.Today.Day.ToString("D2").PadLeft(2, '0'); // dd
                                txtFileName += "_" + nOspNum.ToString().PadLeft(2, '0');
                                //txtFileName += nFileNum.ToString("D5").PadLeft(5, '0');// nnnNN – порядковый номер электронного сообщения за указанный день (00001-99999).
                                txtFileName += ".txt";

                                // путь для файла-запроса в Сбер - путь из настроек + дата за сегодня
                                string txtSberFolder = ff.CreatePathWithDateS(txtGibddOutPath);
                                //cnt = mvv.WriteToXML(constrGIBDD, txtConStrPKOSP, txtSberFolder, txtSberXmlFileName, DT_doc_fiz_sber, nOspNum, lLogger);
                                cnt = 0;
                                foreach (DataRow row in DT_doc_gibdd10.Rows)
                                {
                                    string txtText = Convert.ToString(row["ZAPROS"]).Trim();
                                    txtText += ";" + Convert.ToString(row["debtor_surname"]).Trim();
                                    txtText += ";" + Convert.ToString(row["debtor_firstname"]).Trim();
                                    txtText += ";" + Convert.ToString(row["debtor_patronymic"]).Trim();
                                    txtText += ";" + Convert.ToDateTime(row["DATROZHD"]).ToShortDateString();
                                    if (ff.WriteTofile(txtText, string.Format(@"{0}\{1}", txtSberFolder, txtFileName), Encoding.GetEncoding(1251)))
                                    {
                                        lLogger.WriteLLog("Обработан запрос # " + cnt.ToString() + " request_id = " + Convert.ToString(row["ZAPROS"]).Trim().ToString() + "\n");
                                        cnt++;
                                    }
                                    else
                                    {
                                        lLogger.WriteLLog("Не удалось записать в файл запрос # " + cnt.ToString() + " request_id = " + Convert.ToString(row["ZAPROS"]).Trim().ToString() + "\n");
                                        row["GOD"] = -1;
                                    }
                                }

                                ShowLoggerError(lLogger);
                                if (DT_doc_gibdd10 != null)
                                {
                                    foreach (DataRow row in DT_doc_gibdd10.Rows)// select только ради сортировки
                                    {
                                        //UpdatePackRequest(row);
                                        //UpdateKredOrgRequest(row);

                                        mvv.UpdateExtRequestRow(conPK_OSP, row, lLogger);
                                        //UpdateExtRequestThrowLegalList(row);

                                        //UpdateExtRequestRow(row);
                                        //prbWritingDBF.PerformStep();
                                        //prbWritingDBF.Refresh();
                                        //System.Windows.Forms.Application.DoEvents();
                                    }
                                }
                                if (cnt > 0)
                                {
                                    lLogger.UpdateLLogCount(Convert.ToInt32(cnt));
                                    lLogger.UpdateLLogStatus(2);
                                    lLogger.UpdateLLogFileName(txtFileName);
                                }



                            }
                            # endregion

                            # region "IC_MVD_10"
                            if (bRunIC_MVD_out)
                            {
                                // сейчас выгружается каждый раз при запуске
                                // предлагаю выгружать только в понедельник

                                // день недели от 1 до 7
                                int day = ((int)DateTime.Now.DayOfWeek == 0) ? 7 : (int)DateTime.Now.DayOfWeek;
                                if (txtZaprosOut.Equals("1")) day = 1;

                                // добавить проверку что если вчера был понедельник, и не было выгрузки - то тоже выгружать
                                // проверять через запрос к ufssprk-tools по коду соглашения и типу пакета
                                int nTomorrowUnloaded = 0;
                                txtAgreementCode = "ИЦ_МВД_10";
                                if (day.Equals(2))
                                {
                                    nTomorrowUnloaded = mvv.GetDayPackCount(constrGIBDD, 1, txtAgreementCode, DateTime.Today.AddDays(-1), lCommonLogger);
                                    if (nTomorrowUnloaded.Equals(0)) day = 1;
                                }


                                if (day == 1)
                                // || (DateTime.Today.ToShortDateString().Equals("26.05.2016")))
                                {
                                    Int64 cnt = 0;
                                    DataTable DT_doc_ic_mvd_10 = null;
                                    txtAgreementCode = "ИЦ_МВД_10";
                                    string txtAgreementSber = "ИЦ_МВД_10";

                                    Logger_ufssprk_tools lLoggerError = null;

                                    Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 1, txtAgreementSber, 0, nOspNum, "Пакет запросов в ИЦ МВД.");
                                    DT_doc_ic_mvd_10 = mvv.ReadIcMvd10Zapros(conPK_OSP, lLogger);
                                    ShowLoggerError(lLogger);

                                    // TODO: посмотреть по логам какой по счету файл за сегодня мы выгружаем
                                    int nFileNum = 0;
                                    //// packType 1 - обычный запрос
                                    nFileNum = mvv.GetDayPackCount(constrGIBDD, 1, txtAgreementSber, DateTime.Today, lLogger); //  +1;  - тут нумерация с 0
                                    ShowLoggerError(lLogger);

                                    //string txtSberXmlFileName = ff.makenewSberFileName() + '.' + ff.makenewSberFileExt(nFileNum, nOspNum);
                                    // SSP_yyyymmdd_ХХ_aaaaa_bbb_ccc_nnnNN.{XML/ZIP}, где:
                                    string txtFileName = "tr";
                                    txtFileName += DateTime.Today.Year.ToString("D4"); // yyyy
                                    txtFileName += DateTime.Today.Month.ToString("D2").PadLeft(2, '0'); // mm
                                    txtFileName += DateTime.Today.Day.ToString("D2").PadLeft(2, '0'); // dd
                                    txtFileName += "_" + nOspNum.ToString().PadLeft(2, '0');
                                    txtFileName += "_" + nFileNum.ToString("D3").PadLeft(3, '0');// nNN – порядковый номер электронного сообщения за указанный день (001-999).
                                    txtFileName += ".rc1"; // 20160330 - просят именно rc1

                                    // путь для файла-запроса в Сбер - путь из настроек + дата за сегодня
                                    string txtSberFolder = ff.CreatePathWithDateS(txtIC_MVD_Path);
                                    //cnt = mvv.WriteToXML(constrGIBDD, txtConStrPKOSP, txtSberFolder, txtSberXmlFileName, DT_doc_fiz_sber, nOspNum, lLogger);
                                    cnt = 0;
                                    decimal iECnt = 0;
                                    /*
        ЖЖЖOSKЖЖЖ
        --------------------------------------------
        ФАМИЛИЯ: ИВАНОВ 
        ИМЯ: ИВАН
        ОТЧЕСТВО: ИВАНОВИЧ
        ДАТА РОЖДЕНИЯ: 19.06.1979
        МЕСТО РОЖДЕНИЯ: Коми АССР Г.Микунь
        ЦЕЛЬ ПРОВЕРКИ: СВЕДЕНИЯ О СУДИМОСТИ
        ИНИЦИАТОР ПРОВЕРКИ: УФССП ПО РЕСП КАРЕЛИЯ
        ФАМИЛИЯ ИНИЦИАТОРА ПРОВЕРКИ: СИДОРОВ
        ДАТА: 02.02.2016
        РЕГИОН: 162
        --------------------------------------------
        ЖЖЖКККЖЖЖ
                                     */

                                    string txtOldFio = "";
                                    string txtCurrentFio = "";

                                    string txtOldDateBornD = "";
                                    string txtDateBornD = "";
                                    DateTime OldBornDate = DateTime.MinValue;
                                    DateTime CurrentBornDate = DateTime.MinValue;

                                    string txtOldBirthPlace = "";
                                    string txtBirthPlace = "";
                                    int iIDCount = 0;
                                    int nMultiReqLength = 0;
                                    string txtListID = "";
                                    Int32 iRowCount = 0; // счетчик строк
                                    // добавляем параметр - макс длина строки
                                    // начнем с 4 ID
                                    int nMaxMultiReqLength = 4;

                                    

                                    //foreach (DataRow row in DT_doc_ic_mvd_10.Rows) 
                                    for (int j = 0; j < DT_doc_ic_mvd_10.Rows.Count; j++)
                                    {
                                        DataRow row = DT_doc_ic_mvd_10.Rows[j];

                                        // параметры для сбора списка ID
                                        txtCurrentFio = Convert.ToString(row["FIOVK"]).Trim().ToUpper();
                                        txtBirthPlace = Convert.ToString(row["DEBTOR_BIRTHPLACE"]).Trim().ToUpper();

                                        DateTime dtDate = DateTime.MaxValue;
                                        string txtDate = "";

                                        if (!row["DATROZHD"].Equals(System.DBNull.Value))
                                        {
                                            txtDateBornD = Convert.ToString(row["DATROZHD"]).Trim();
                                        }
                                        else
                                        {
                                            txtDateBornD = "";
                                        }

                                        // если это такая же строчка как и выше
                                        // 20171218 - нужно убрать этот механизм - попробую просто закомментировать эту проверку вобще
                                        //bool bMultiRequest = false;

                                        // 20201116 - вернул свертку запросов по разным ИП на одного должника
                                        // 20210122 - вернул временно на место bool bMultiRequest = false;
                                        bool bMultiRequest = true;

                                        if (bMultiRequest)
                                            if ((nMultiReqLength < nMaxMultiReqLength )
                                         && (txtCurrentFio == txtOldFio) && (txtDateBornD == txtOldDateBornD)
                                         && ((txtBirthPlace == txtOldBirthPlace) || (txtBirthPlace.Length == 0) || (txtOldBirthPlace.Length == 0)))
                                        {
                                            if (txtListID.Length > 0) txtListID += ";"; // это только если не первый раз
                                            txtListID += Convert.ToString(row["ZAPROS"]).Trim();
                                            iIDCount++; // счетчик ID
                                            nMultiReqLength++;  // счетчик ID в строке
                                        }
                                        else
                                        {
                                            // а если строчка не такая-же - то будем вставлять
                                            // при этом писать надо не текущую, а предыдущую
                                            // ... iRowCount-1, и не забыть дату поменять и убрать поля ненужные
                                            // если это первая строчка - то не надо писать, просто начать новую
                                            if (j > 0)
                                            {
                                                // TODO: постараться максимально найти место рождения
                                                // вопрос - что если предыдущая строчка была невалидной по дате рождения - то там была ошибка вставлена
                                                // но нужно делать вставку ошибки по txtIdList
                                                if (mvv.WriteIcMvdReqRow(DT_doc_ic_mvd_10.Rows[j - 1], txtListID, txtSberFolder, txtFileName, conPK_OSP, constrGIBDD, txtAgreementSber, ref lLoggerError, lLogger))
                                                {
                                                    lLogger.WriteLLog("Обработан запрос # " + cnt.ToString() + " request_id = " + txtListID + "\n");
                                                    cnt++;
                                                }
                                                else
                                                {
                                                    lLogger.WriteLLog("Не удалось записать в файл запрос # " + cnt.ToString() + " request_id = " + txtListID + "\n");
                                                    iECnt++;
                                                }

                                                iRowCount++; // считаем строчку обработанной
                                            }

                                            // сначала проверим а не отвалились ли мы по условию     (nMultiReqLength < nMaxMultiReqLength )
                                            if (nMultiReqLength >= nMaxMultiReqLength)
                                            {
                                                // будет ли это работать если все совпадает
                                                txtListID = "";
                                                txtOldFio = txtCurrentFio;
                                                txtOldDateBornD = txtDateBornD;
                                                txtOldBirthPlace = txtBirthPlace;

                                                iIDCount = 0;
                                                nMultiReqLength = 0;

                                                txtListID += Convert.ToString(row["ZAPROS"]).Trim();
                                                iIDCount++; // счетчик ИД
                                                nMultiReqLength++;

                                            }
                                            else
                                            {
                                            // Инициализировть новый цикл сбора статистики
                                                txtListID = "";
                                                txtOldFio = txtCurrentFio;
                                                txtOldDateBornD = txtDateBornD;
                                                txtOldBirthPlace = txtBirthPlace;

                                                iIDCount = 0;
                                                nMultiReqLength = 0;

                                                txtListID += Convert.ToString(row["ZAPROS"]).Trim();
                                                iIDCount++; // счетчик ИД
                                                nMultiReqLength++;

                                            } // end else (nMultiReqLength < nMaxMultiReqLength )

                                        }

                                        // если это была последняя строчка - то ее надо записать в любом! случае
                                        if (j == DT_doc_ic_mvd_10.Rows.Count - 1)
                                        {
                                            if (mvv.WriteIcMvdReqRow(DT_doc_ic_mvd_10.Rows[j], txtListID, txtSberFolder, txtFileName, conPK_OSP, constrGIBDD, txtAgreementSber, ref lLoggerError, lLogger))
                                            {
                                                lLogger.WriteLLog("Обработан запрос # " + cnt.ToString() + " request_id = " + txtListID + "\n");
                                                cnt++;
                                            }
                                            else
                                            {
                                                lLogger.WriteLLog("Не удалось записать в файл запрос # " + cnt.ToString() + " request_id = " + txtListID + "\n");
                                                iECnt++;
                                            }
                                        }
                                    } // end_for

                                    ShowLoggerError(lLogger);
                                    if (DT_doc_ic_mvd_10 != null)
                                    {
                                        foreach (DataRow row in DT_doc_ic_mvd_10.Rows)// select только ради сортировки
                                        {
                                            //UpdatePackRequest(row);
                                            //UpdateKredOrgRequest(row);

                                            mvv.UpdateExtRequestRow(conPK_OSP, row, lLogger);
                                            //UpdateExtRequestThrowLegalList(row);

                                            //UpdateExtRequestRow(row);
                                            //prbWritingDBF.PerformStep();
                                            //prbWritingDBF.Refresh();
                                            //System.Windows.Forms.Application.DoEvents();
                                        }
                                    }
                                    if (cnt > 0)
                                    {
                                        lLogger.UpdateLLogCount(Convert.ToInt32(cnt));
                                        lLogger.UpdateLLogStatus(2);
                                        lLogger.UpdateLLogFileName(txtFileName);
                                    }
                                }
                            }
                            # endregion

                            # region "GIBDD_10_SELFREQUEST"
                            // автоматический ответ на запрос через базу данных
                            // параметры - bRunGibdd10_selfrequest
                            //           - txtGibdd10ConString
                            //           - txtGibdd10DataBase

                            if (bRunGibdd10_selfrequest)
                            {
                                Int64 cnt = 0;

                                DataTable DT_doc_gibdd10 = null;

                                string txtAgreementSber = "ГИБДД_10";
                                txtAgreementCode = "ГИБДД_10";

                                Logger_ufssprk_tools lLoggerError = null;
                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 1, txtAgreementSber, 0, nOspNum, "Пакет запросов в ГИБДД_10.");
                                DT_doc_gibdd10 = mvv.ReadGibdd10Zapros(conPK_OSP, lLogger);
                                // DT_doc_fiz_sber = mvv.ReadSberZaprosNewFormatTest(conPK_OSP, lLogger);
                                ShowLoggerError(lLogger);

                                //  сразу вопрос - как писать логи
                                //  менять ли тип пакета (сделать специальный packType для таких Самообрабатывающихся запросов)
                                //  или писать сначала лог на выгрузку, а потом лог на вставку - наверное проще так сначала,
                                //  а потом переделать уже на новый тип лога

                                // записать в таблицу что получили строльк-то запросов и все, переходим к обработке - то есть загрузке ответов
                                if (DT_doc_gibdd10 != null)
                                {
                                    lLogger.WriteLLog("Выгружено запросов в ГИБДД о ВУ: " + DT_doc_gibdd10.Rows.Count.ToString() + "\n");
                                    lLogger.UpdateLLogCount(DT_doc_gibdd10.Rows.Count);
                                    lLogger.UpdateLLogStatus(2);
                                }
                                // после получения таблицы будем делать 

                                // открыть лог ответа
                                nPack_type = 2; // 2 - простой ответ
                                decimal nParentID = lLogger.logID;
                                lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, nParentID, nOspNum, "Пакет загрузки ответов на запрос о наличии ВУ в ГИБДД МВД");

                                if (DT_doc_gibdd10 != null)
                                {
                                    // написать когда и что делаем
                                    lLogger.WriteLLog(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "Начало обработки запросов по базе водительских удостоверений из ГИБДД МВД. Строк запросов всего: " + DT_doc_gibdd10.Rows.Count.ToString() + "\r\n");

                                    OleDbConnection VUcon = new OleDbConnection(txtGibdd10ConString);

                                    // в цикле по всем строкам DT_doc_gibdd10 
                                    string txtId = "";
                                    decimal nId = 0;
                                    decimal nRespId = 0;
                                    decimal nStatus = 19; // ответ получен
                                    string txtAnswerType = "01";
                                    string txtEntityName = "ГИБДД МВД РК";


                                    foreach (DataRow row in DT_doc_gibdd10.Rows)
                                    {
                                        // получить ID запроса
                                        txtId = Convert.ToString(row["ZAPROS"]);
                                        Decimal.TryParse(txtId, out nId);
                                        if (nId > 0)
                                        {
                                            string txtFioD = Convert.ToString(row["FIOVK"]).Trim().ToUpper();
                                            // убрать двойные пробелы
                                            txtFioD = RemoveDoubleSpaces(txtFioD, 200);
                                            DateTime dtBornD = DateTime.MinValue;
                                            string txtBornD = "";
                                            string txtDocNum = "";
                                            DateTime dtStart = DateTime.MinValue;
                                            DateTime dtEnd = DateTime.MinValue;
                                            DateTime dtDateVh = DateTime.MinValue;


                                            if (!row["DATROZHD"].Equals(DBNull.Value))
                                            {
                                                txtBornD = Convert.ToString(row["DATROZHD"]);
                                                if (DateTime.TryParse(txtBornD, out dtBornD))
                                                {
                                                    // TODO: получить ответ как DataTable и тут уже разбирать параметры и собирать строчку - иначе не вставить структурированные сведения
                                                    // string txtResp = mvv.FindVU(VUcon, txtGibdd10DataBase, txtFioD, dtBornD, lLogger);
                                                    DataTable tbl = mvv.FindVUdt(VUcon, txtGibdd10DataBase, txtFioD, dtBornD, lLogger);
                                                    string txtResp = "";
                                                    if (tbl != null && tbl.Rows.Count > 0)
                                                    {
                                                        txtResp += "Получены сведения о наличии у должника " + txtFioD + " (" + dtBornD.ToShortDateString() + " г.р.) водительских удостоверений.\r\n";
                                                        if (tbl.Rows.Count > 1) txtResp += "Всего удостоверений: " + tbl.Rows.Count.ToString() + "\r\n";
                                                        foreach (DataRow rrow in tbl.Rows)
                                                        {
                                                            txtDocNum = Convert.ToString(rrow["VU_NUMBER"]).Trim();
                                                            txtResp += "Номер ВУ: " + txtDocNum + "\r\n";
                                                            dtStart = DateTime.MinValue;
                                                            string txtStart = "";
                                                            txtResp += "Дата выдачи ВУ: ";
                                                            if (!rrow["OUT_D"].Equals(DBNull.Value))
                                                            {
                                                                txtStart = Convert.ToString(rrow["OUT_D"]);
                                                                if (DateTime.TryParse(txtStart, out dtStart))
                                                                    txtResp += dtStart.ToShortDateString() + "\r\n";
                                                                else txtResp += "сведения отсутствуют\r\n";
                                                            }
                                                            else txtResp += "сведения отсутствуют\r\n";

                                                            dtEnd = DateTime.MinValue;
                                                            string txtEnd = "";
                                                            txtResp += "Дата окончания срока действия ВУ: ";
                                                            if (!rrow["END_D"].Equals(DBNull.Value))
                                                            {
                                                                txtEnd = Convert.ToString(rrow["END_D"]);
                                                                if (DateTime.TryParse(txtEnd, out dtEnd))
                                                                    txtResp += dtEnd.ToShortDateString() + "\r\n";
                                                                else txtResp += "сведения отсутствуют\r\n";
                                                            }
                                                            else txtResp += "сведения отсутствуют\r\n";

                                                            dtDateVh = DateTime.MinValue;
                                                            string txtDateVh = "";
                                                            if (!rrow["DATE_VH"].Equals(DBNull.Value))
                                                            {
                                                                txtDateVh = Convert.ToString(rrow["DATE_VH"]);
                                                                DateTime.TryParse(txtDateVh, out dtDateVh);
                                                            }

                                                            if (tbl.Rows.Count > 1) txtResp += "\r\n"; // перевод строки чтобы разделить записи если их больше 1
                                                        }
                                                    }
                                                    else txtResp = "Нет данных";

                                                    txtAnswerType = "01";
                                                    if (txtResp == "Нет данных") txtAnswerType = "02";
                                                    else if (txtResp.Length == 0) txtAnswerType = "03"; //  требует уточнения
                                                    // вставить ответ в ИТ
                                                    nRespId = mvv.InsertResponseIntTable(conPK_OSP, nId, txtResp, DateTime.Today, nStatus, lLogger.logID, txtAgreementCode, txtAgreementCode, txtAgreementCode, txtEntityName, txtAnswerType, lLogger);
                                                    if (nRespId > 0)
                                                    {

                                                        // если ответ был с данными - вставить в ИТ EXT_IDENTIFICATION_DATA структурированные сведения
                                                        if (txtAnswerType == "01")
                                                        {
                                                            // это код 
                                                            // Значения заполняются в соответствии со справочником DIRECTORY_TYPES (Коды документов, удостоверяющих личность)
                                                            string txtTypeDocCode = "91"; // Иные документы, предусмотренные законодательством Российской Федерации
                                                            // ENTITY_NAME - для доп. сведений это ФИО должника
                                                            // 20160513 - закомментировал т.к. это пока не работает
                                                            // decimal nExtIdentID = mvv.InsertExtIdentificationData(conPK_OSP, nRespId, txtFioD, dtDateVh, txtDocNum, dtStart, txtFioD, txtTypeDocCode, lLogger);
                                                        }
                                                        // теперь нужно обозначить запрос в ИТ как выгруженный
                                                        if (mvv.UpdateExtRequestRow(conPK_OSP, row, lLogger))
                                                        {
                                                            lLogger.WriteLLog("Запрос № " + txtId.ToString() + " успешно обработан. Получен ответ № " + nRespId.ToString() + "\r\n");
                                                            cnt++;
                                                        }
                                                        else
                                                            lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Несмотря на это в базу загружен ответ № " + nRespId.ToString() + "\r\n");
                                                    }
                                                    else
                                                    {
                                                        lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n");
                                                        lLogger.ErrMessage += "Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n";
                                                    }
                                                }
                                                else
                                                {
                                                    // TODO: вставить ответ что нет даты рождения
                                                    // вставить ответ что нет даты рождения
                                                    decimal nErrId = 0;
                                                    nErrId = mvv.InsertIcMvdErrorResp(conPK_OSP, constrGIBDD, txtAgreementCode, nId, txtEntityName, ref lLoggerError, lLogger);
                                                    if (nErrId > 0)
                                                    {
                                                        lLogger.WriteLLog("Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n");
                                                        lLogger.ErrMessage += "Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n";
                                                    }
                                                    else
                                                    {
                                                        lLogger.WriteLLog("Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n");
                                                        lLogger.ErrMessage += "Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n";
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                // вставить ответ что нет даты рождения
                                                decimal nErrId = 0;
                                                nErrId = mvv.InsertIcMvdErrorResp(conPK_OSP, constrGIBDD, txtAgreementCode, nId, txtEntityName, ref lLoggerError, lLogger);
                                                if (nErrId > 0)
                                                {
                                                    lLogger.WriteLLog("Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n");
                                                    lLogger.ErrMessage += "Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n";
                                                }
                                                else
                                                {
                                                    lLogger.WriteLLog("Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n");
                                                    lLogger.ErrMessage += "Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n";
                                                }
                                            }
                                        }
                                    } // end_for
                                }
                                if (cnt > 0)
                                {
                                    lLogger.UpdateLLogCount(Convert.ToInt32(cnt));
                                    lLogger.UpdateLLogStatus(2);
                                }
                            }
                            # endregion

                            # region "GIMS_10_SELFREQUEST"
                            // автоматический ответ на запрос через базу данных

                            if (bRunGims10_selfrequest)
                            {
                                Int64 cnt = 0;

                                DataTable DT_doc_gims10 = null;

                                string txtAgreementSber = "ГИМС_10";
                                txtAgreementCode = "ГИМС_10";

                                Logger_ufssprk_tools lLoggerError = null;
                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 1, txtAgreementSber, 0, nOspNum, "Пакет запросов в ГИМС_10.");
                                DT_doc_gims10 = mvv.ReadGims10Zapros(conPK_OSP, lLogger);
                                // DT_doc_fiz_sber = mvv.ReadSberZaprosNewFormatTest(conPK_OSP, lLogger);
                                ShowLoggerError(lLogger);

                                //  сразу вопрос - как писать логи
                                //  менять ли тип пакета (сделать специальный packType для таких Самообрабатывающихся запросов)
                                //  или писать сначала лог на выгрузку, а потом лог на вставку - наверное проще так сначала,
                                //  а потом переделать уже на новый тип лога

                                // записать в таблицу что получили строльк-то запросов и все, переходим к обработке - то есть загрузке ответов
                                if (DT_doc_gims10 != null)
                                {
                                    lLogger.WriteLLog("Выгружено запросов в ГИМС МЧС о ВУ: " + DT_doc_gims10.Rows.Count.ToString() + "\n");
                                    lLogger.UpdateLLogCount(DT_doc_gims10.Rows.Count);
                                    lLogger.UpdateLLogStatus(2);
                                }
                                // после получения таблицы будем делать 

                                // открыть лог ответа
                                nPack_type = 2; // 2 - простой ответ
                                decimal nParentID = lLogger.logID;
                                lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, nParentID, nOspNum, "Пакет загрузки ответов на запрос о наличии ВУ в ГИМС МЧС");

                                if (DT_doc_gims10 != null)
                                {
                                    // написать когда и что делаем
                                    lLogger.WriteLLog(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "Начало обработки запросов по базе водительских удостоверений из ГИМС МЧС. Строк запросов всего: " + DT_doc_gims10.Rows.Count.ToString() + "\r\n");

                                    OleDbConnection VUcon = new OleDbConnection(txtGibdd10ConString);

                                    // в цикле по всем строкам DT_doc_gibdd10 
                                    string txtId = "";
                                    decimal nId = 0;
                                    decimal nRespId = 0;
                                    decimal nStatus = 19; // ответ получен
                                    string txtAnswerType = "01";
                                    string txtEntityName = "ГИМС МЧС России по Республике Карелия";


                                    foreach (DataRow row in DT_doc_gims10.Rows)
                                    {
                                        // получить ID запроса
                                        txtId = Convert.ToString(row["ZAPROS"]);
                                        Decimal.TryParse(txtId, out nId);
                                        if (nId > 0)
                                        {
                                            string txtFioD = Convert.ToString(row["FIOVK"]).Trim().ToUpper();
                                            // убрать двойные пробелы
                                            txtFioD = RemoveDoubleSpaces(txtFioD, 200);


                                            DateTime dtBornD = DateTime.MinValue;
                                            string txtBornD = "";
                                            string txtDocNum = "";
                                            DateTime dtStart = DateTime.MinValue;
                                            DateTime dtDateVh = DateTime.MinValue;


                                            if (!row["DATROZHD"].Equals(DBNull.Value))
                                            {
                                                txtBornD = Convert.ToString(row["DATROZHD"]);
                                                if (DateTime.TryParse(txtBornD, out dtBornD))
                                                {
                                                    // TODO: получить ответ как DataTable и тут уже разбирать параметры и собирать строчку - иначе не вставить структурированные сведения
                                                    // string txtResp = mvv.FindVU(VUcon, txtGibdd10DataBase, txtFioD, dtBornD, lLogger);
                                                    DataTable tbl = mvv.FindGimsVUdt(VUcon, txtGibdd10DataBase, txtFioD, dtBornD, lLogger);
                                                    string txtResp = "";
                                                    if (tbl != null && tbl.Rows.Count > 0)
                                                    {
                                                        txtResp += "Получены сведения о наличии у должника " + txtFioD + " (" + dtBornD.ToShortDateString() + " г.р.) удостоверения на право управления маломерными судами.\r\n";
                                                        if (tbl.Rows.Count > 1) txtResp += "Всего удостоверений: " + tbl.Rows.Count.ToString() + "\r\n";
                                                        foreach (DataRow rrow in tbl.Rows)
                                                        {
                                                            txtDocNum = Convert.ToString(rrow["VU_NUMBER"]).Trim();
                                                            txtResp += "Номер удостоверения: " + txtDocNum + "\r\n";
                                                            dtStart = DateTime.MinValue;
                                                            string txtStart = "";
                                                            txtResp += "Дата выдачи удостоверения: ";
                                                            if (!rrow["OUT_D"].Equals(DBNull.Value))
                                                            {
                                                                txtStart = Convert.ToString(rrow["OUT_D"]);
                                                                if (DateTime.TryParse(txtStart, out dtStart))
                                                                    txtResp += dtStart.ToShortDateString() + "\r\n";
                                                                else txtResp += "сведения отсутствуют\r\n";
                                                            }
                                                            else txtResp += "сведения отсутствуют\r\n";


                                                            dtDateVh = DateTime.MinValue;
                                                            string txtDateVh = "";
                                                            if (!rrow["DATE_VH"].Equals(DBNull.Value))
                                                            {
                                                                txtDateVh = Convert.ToString(rrow["DATE_VH"]);
                                                                DateTime.TryParse(txtDateVh, out dtDateVh);
                                                            }

                                                            if (tbl.Rows.Count > 1) txtResp += "\r\n"; // перевод строки чтобы разделить записи если их больше 1
                                                        }
                                                    }
                                                    else txtResp = "Нет данных";

                                                    txtAnswerType = "01";
                                                    if (txtResp == "Нет данных") txtAnswerType = "02";
                                                    else if (txtResp.Length == 0) txtAnswerType = "03"; //  требует уточнения
                                                    // вставить ответ в ИТ
                                                    nRespId = mvv.InsertResponseIntTable(conPK_OSP, nId, txtResp, DateTime.Today, nStatus, lLogger.logID, txtAgreementCode, txtAgreementCode, txtAgreementCode, txtEntityName, txtAnswerType, lLogger);
                                                    if (nRespId > 0)
                                                    {

                                                        // если ответ был с данными - вставить в ИТ EXT_IDENTIFICATION_DATA структурированные сведения
                                                        if (txtAnswerType == "01")
                                                        {
                                                            // это код 
                                                            // Значения заполняются в соответствии со справочником DIRECTORY_TYPES (Коды документов, удостоверяющих личность)
                                                            string txtTypeDocCode = "91"; // Иные документы, предусмотренные законодательством Российской Федерации
                                                            // ENTITY_NAME - для доп. сведений это ФИО должника
                                                            // 20160513 - закомментировал т.к. это не работает
                                                            // decimal nExtIdentID = mvv.InsertExtIdentificationData(conPK_OSP, nRespId, txtFioD, dtDateVh, txtDocNum, dtStart, txtFioD, txtTypeDocCode, lLogger);
                                                        }
                                                        // теперь нужно обозначить запрос в ИТ как выгруженный
                                                        if (mvv.UpdateExtRequestRow(conPK_OSP, row, lLogger))
                                                        {
                                                            lLogger.WriteLLog("Запрос № " + txtId.ToString() + " успешно обработан. Получен ответ № " + nRespId.ToString() + "\r\n");
                                                            cnt++;
                                                        }
                                                        else
                                                            lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Несмотря на это в базу загружен ответ № " + nRespId.ToString() + "\r\n");
                                                    }
                                                    else
                                                    {
                                                        lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n");
                                                        lLogger.ErrMessage += "Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n";
                                                    }
                                                }
                                                else
                                                {
                                                    // TODO: вставить ответ что нет даты рождения
                                                    // вставить ответ что нет даты рождения
                                                    decimal nErrId = 0;
                                                    nErrId = mvv.InsertIcMvdErrorResp(conPK_OSP, constrGIBDD, txtAgreementCode, nId, txtEntityName, ref lLoggerError, lLogger);
                                                    if (nErrId > 0)
                                                    {
                                                        lLogger.WriteLLog("Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n");
                                                        lLogger.ErrMessage += "Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n";
                                                    }
                                                    else
                                                    {
                                                        lLogger.WriteLLog("Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n");
                                                        lLogger.ErrMessage += "Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n";
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                // вставить ответ что нет даты рождения
                                                decimal nErrId = 0;
                                                nErrId = mvv.InsertIcMvdErrorResp(conPK_OSP, constrGIBDD, txtAgreementCode, nId, txtEntityName, ref lLoggerError, lLogger);
                                                if (nErrId > 0)
                                                {
                                                    lLogger.WriteLLog("Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n");
                                                    lLogger.ErrMessage += "Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n";
                                                }
                                                else
                                                {
                                                    lLogger.WriteLLog("Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n");
                                                    lLogger.ErrMessage += "Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n";
                                                }
                                            }
                                        }
                                    } // end_for
                                }
                                if (cnt > 0)
                                {
                                    lLogger.UpdateLLogCount(Convert.ToInt32(cnt));
                                    lLogger.UpdateLLogStatus(2);
                                }
                            }
                            # endregion

                            # region "RTN_10_SELFREQUEST"
                            // автоматический ответ на запрос через базу данных

                            if (bRunRtn10_selfrequest)
                            {
                                Int64 cnt = 0;

                                DataTable DT_doc_rtn10 = null;

                                string txtAgreementSber = "РТН_10";
                                txtAgreementCode = "РТН_10";

                                Logger_ufssprk_tools lLoggerError = null;
                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 1, txtAgreementSber, 0, nOspNum, "Пакет запросов в РТН_10.");
                                DT_doc_rtn10 = mvv.ReadRtn10Zapros(conPK_OSP, lLogger);
                                // DT_doc_fiz_sber = mvv.ReadSberZaprosNewFormatTest(conPK_OSP, lLogger);
                                ShowLoggerError(lLogger);

                                //  сразу вопрос - как писать логи
                                //  менять ли тип пакета (сделать специальный packType для таких Самообрабатывающихся запросов)
                                //  или писать сначала лог на выгрузку, а потом лог на вставку - наверное проще так сначала,
                                //  а потом переделать уже на новый тип лога

                                // записать в таблицу что получили строльк-то запросов и все, переходим к обработке - то есть загрузке ответов
                                if (DT_doc_rtn10 != null)
                                {
                                    lLogger.WriteLLog("Выгружено запросов в Ростехнадзор: " + DT_doc_rtn10.Rows.Count.ToString() + "\n");
                                    lLogger.UpdateLLogCount(DT_doc_rtn10.Rows.Count);
                                    lLogger.UpdateLLogStatus(2);
                                }
                                // после получения таблицы будем делать 

                                // открыть лог ответа
                                nPack_type = 2; // 2 - простой ответ
                                decimal nParentID = lLogger.logID;
                                lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, nParentID, nOspNum, "Пакет загрузки ответов на запрос в Ростехнадзор");

                                if (DT_doc_rtn10 != null)
                                {
                                    // написать когда и что делаем
                                    lLogger.WriteLLog(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "Начало обработки запросов по базе Ростехнадзора. Строк запросов всего: " + DT_doc_rtn10.Rows.Count.ToString() + "\r\n");

                                    OleDbConnection VUcon = new OleDbConnection(txtGibdd10ConString);

                                    // в цикле по всем строкам DT_doc_gibdd10 
                                    string txtId = "";
                                    decimal nId = 0;
                                    decimal nRespId = 0;
                                    decimal nStatus = 19; // ответ получен
                                    string txtAnswerType = "01";
                                    string txtEntityName = "Северо-Западное управление Ростехнадзора (Республика Карелия)";


                                    foreach (DataRow row in DT_doc_rtn10.Rows)
                                    {
                                        // получить ID запроса
                                        // учитывать что если класс контрагента не физ. лицо - то смотреть на ИНН
                                        // классы физ. лиц 2, 71,95,96,97,666
                                        // инд. предприниматель  95
                                        // итого (2,71,95,96,97,666)
                                        txtId = Convert.ToString(row["ZAPROS"]);
                                        Decimal.TryParse(txtId, out nId);
                                        if (nId > 0)
                                        {
                                            string txtFioD = Convert.ToString(row["FIOVK"]).Trim().ToUpper();
                                            // убрать двойные пробелы
                                            txtFioD = RemoveDoubleSpaces(txtFioD, 200);


                                            DateTime dtBornD = DateTime.MinValue;
                                            string txtBornD = "";
                                            string txtDocNum = "";
                                            DateTime dtStart = DateTime.MinValue;
                                            DateTime dtDateVh = DateTime.MinValue;

                                            List<int> fiz = new List<int>(new[] { 2, 71, 95, 96, 97, 666 });
                                            int iDbtrCls = Convert.ToInt32(row["ID_DBTRCLS"]);
                                            // дату рождения проверяем только для физ лиц
                                            // для юр лиц нужно будет проверять ИНН и искать по нему

                                            if (fiz.Contains(iDbtrCls) && !row["DATROZHD"].Equals(DBNull.Value))
                                            {
                                                txtBornD = Convert.ToString(row["DATROZHD"]);
                                                if (DateTime.TryParse(txtBornD, out dtBornD))
                                                {
                                                    // TODO: получить ответ как DataTable и тут уже разбирать параметры и собирать строчку - иначе не вставить структурированные сведения
                                                    // string txtResp = mvv.FindVU(VUcon, txtGibdd10DataBase, txtFioD, dtBornD, lLogger);
                                                    DataTable tbl = mvv.FindRtnFlDt(VUcon, txtFioD, dtBornD, lLogger);
                                                    string txtResp = "";
                                                    if (tbl != null && tbl.Rows.Count > 0)
                                                    {
                                                        txtResp += "Получены сведения о наличии зарегистрированного за должником " + txtFioD + " (" + dtBornD.ToShortDateString() + " г.р.) движимого имущества.\r\n";
                                                        if (tbl.Rows.Count > 1)
                                                        {
                                                            txtResp += "Всего объектов имущества: " + tbl.Rows.Count.ToString() + "\r\n";
                                                        }
                                                        if (tbl.Rows.Count > 0)
                                                        {
                                                            string txtAct_date = "";
                                                            if (tbl.Columns.Contains("ACT_DATE")) txtAct_date = Convert.ToString(tbl.Rows[0]["ACT_DATE"]);
                                                            if (txtAct_date.Length > 0)
                                                            {
                                                                txtResp += "Дата актуальности сведений: " + txtAct_date + "\r\n";
                                                            }
                                                        }

                                                        foreach (DataRow rrow in tbl.Rows)
                                                        {
                                                            string txtREGNUM = Convert.ToString(rrow["REGNUM"]).Trim();
                                                            txtResp += "Госзнак: " + txtREGNUM + "\r\n";

                                                            string txtTITLE = Convert.ToString(rrow["TITLE"]).Trim();
                                                            txtResp += "Марка: " + txtTITLE + "\r\n";

                                                            string txtCLASS = Convert.ToString(rrow["CLASS"]).Trim();
                                                            txtResp += "Класс: " + txtCLASS + "\r\n";

                                                            string txtM_YEAR = Convert.ToString(rrow["M_YEAR"]).Trim();
                                                            txtResp += "Год выпуска: " + txtM_YEAR + "\r\n";

                                                            string txtREG_DATE = Convert.ToString(rrow["REG_DATE"]).Trim();
                                                            txtResp += "Дата регистрации: " + txtREG_DATE + "\r\n";

                                                            //string txtCHECK_DATE = Convert.ToString(rrow["CHECK_DATE"]).Trim();
                                                            //txtResp += "Дата ТО: " + txtCHECK_DATE + "\r\n";

                                                            string txtOWNER_ADDR = Convert.ToString(rrow["OWNER_ADDR"]).Trim();
                                                            txtResp += "Адрес владельца: " + txtOWNER_ADDR + "\r\n";

                                                            string txtINN = Convert.ToString(rrow["INN"]).Trim();
                                                            if (txtINN.Length >= 10)
                                                                txtResp += "ИНН владельца: " + txtINN + "\r\n";

                                                            if (tbl.Rows.Count > 1) txtResp += "\r\n"; // перевод строки чтобы разделить записи если их больше 1
                                                        }
                                                    }
                                                    else txtResp = "Нет данных";

                                                    txtAnswerType = "01";
                                                    if (txtResp == "Нет данных") txtAnswerType = "02";
                                                    else if (txtResp.Length == 0) txtAnswerType = "03"; //  требует уточнения
                                                    // вставить ответ в ИТ
                                                    // nRespId = mvv.InsertResponseIntTable(conPK_OSP, nId, txtResp, DateTime.Today, nStatus, lLogger.logID, txtAgreementCode, txtAgreementCode, txtAgreementCode, txtEntityName, txtAnswerType, lLogger);
                                                    decimal nExtKey = mvv.InsertResponseIntTableNewExtKey(conPK_OSP, nId, txtResp, DateTime.Today, nStatus, lLogger.logID, txtAgreementCode, txtAgreementCode, txtAgreementCode, txtEntityName, txtAnswerType, lLogger);
                                                    if (nExtKey > 0)
                                                    {

                                                        // если ответ был с данными - вставить в ИТ EXT_IDENTIFICATION_DATA структурированные сведения
                                                        if (txtAnswerType == "01")
                                                        {
                                                            // пройти по всем строчкам и вставить в таблицу только те, которые актуальны (не сняты с учета)
                                                            if (tbl != null && tbl.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow rrow in tbl.Rows)
                                                                {
                                                                    decimal nExtIdentID = mvv.InsertExtDatumMachineRtnIntTable(conPK_OSP, nExtKey, txtFioD, rrow, lLogger);
                                                                }
                                                            }
                                                        }
                                                        // теперь нужно обозначить запрос в ИТ как выгруженный
                                                        if (mvv.UpdateExtRequestRow(conPK_OSP, row, lLogger))
                                                        {
                                                            lLogger.WriteLLog("Запрос № " + txtId.ToString() + " успешно обработан. Получен ответ № " + nExtKey.ToString() + "\r\n");
                                                            cnt++;
                                                        }
                                                        else
                                                            lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Несмотря на это в базу загружен ответ № " + nExtKey.ToString() + "\r\n");
                                                    }
                                                    else
                                                    {
                                                        lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n");
                                                        lLogger.ErrMessage += "Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n";
                                                    }
                                                }
                                                else
                                                {
                                                    // TODO: вставить ответ что нет даты рождения
                                                    // вставить ответ что нет даты рождения
                                                    decimal nErrId = 0;
                                                    nErrId = mvv.InsertIcMvdErrorResp(conPK_OSP, constrGIBDD, txtAgreementCode, nId, txtEntityName, ref lLoggerError, lLogger);
                                                    if (nErrId > 0)
                                                    {
                                                        lLogger.WriteLLog("Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n");
                                                        lLogger.ErrMessage += "Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n";
                                                    }
                                                    else
                                                    {
                                                        lLogger.WriteLLog("Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n");
                                                        lLogger.ErrMessage += "Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n";
                                                    }
                                                }
                                            }
                                            else if (fiz.Contains(iDbtrCls) && row["DATROZHD"].Equals(DBNull.Value))
                                            {
                                                // вставить ответ что нет даты рождения
                                                decimal nErrId = 0;
                                                nErrId = mvv.InsertIcMvdErrorResp(conPK_OSP, constrGIBDD, txtAgreementCode, nId, txtEntityName, ref lLoggerError, lLogger);
                                                if (nErrId > 0)
                                                {
                                                    lLogger.WriteLLog("Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n");
                                                    lLogger.ErrMessage += "Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n";
                                                }
                                                else
                                                {
                                                    lLogger.WriteLLog("Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n");
                                                    lLogger.ErrMessage += "Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n";
                                                }
                                            }
                                            else
                                            {
                                                string txtInn = Convert.ToString(row["id_dbtr_inn"]).Trim();
                                                // 95 - это ИП - он и ФИЗ и ЮР - если нет даты рождения - будем искать по ИНН
                                                if ((!fiz.Contains(iDbtrCls) || iDbtrCls.Equals(95)) && txtInn.Length >= 10)
                                                {
                                                    // ищем как юр.лицо по ИНН
                                                    DataTable tbl = mvv.FindRtnUlDt(VUcon, txtFioD, txtInn, lLogger);
                                                    string txtResp = "";
                                                    if (tbl != null && tbl.Rows.Count > 0)
                                                    {
                                                        txtResp += "Получены сведения о наличии зарегистрированного за должником " + txtFioD + " (поиск велся по ИНН " + txtInn + ") движимого имущества.\r\n";
                                                        if (tbl.Rows.Count > 1) txtResp += "Всего объектов имущества: " + tbl.Rows.Count.ToString() + "\r\n";
                                                        if (tbl.Rows.Count > 0)
                                                        {
                                                            string txtAct_date = "";
                                                            if (tbl.Columns.Contains("ACT_DATE")) txtAct_date = Convert.ToString(tbl.Rows[0]["ACT_DATE"]);
                                                            if (txtAct_date.Length > 0)
                                                            {
                                                                txtResp += "Дата актуальности сведений: " + txtAct_date + "\r\n";
                                                            }
                                                        }
                                                        foreach (DataRow rrow in tbl.Rows)
                                                        {
                                                            string txtREGNUM = Convert.ToString(rrow["REGNUM"]).Trim();
                                                            txtResp += "Госзнак: " + txtREGNUM + "\r\n";

                                                            string txtTITLE = Convert.ToString(rrow["TITLE"]).Trim();
                                                            txtResp += "Марка: " + txtTITLE + "\r\n";

                                                            string txtCLASS = Convert.ToString(rrow["CLASS"]).Trim();
                                                            txtResp += "Класс: " + txtCLASS + "\r\n";

                                                            string txtM_YEAR = Convert.ToString(rrow["M_YEAR"]).Trim();
                                                            txtResp += "Год выпуска: " + txtM_YEAR + "\r\n";

                                                            string txtREG_DATE = Convert.ToString(rrow["REG_DATE"]).Trim();
                                                            txtResp += "Дата регистрации: " + txtREG_DATE + "\r\n";

                                                            string txtCHECK_DATE = Convert.ToString(rrow["CHECK_DATE"]).Trim();
                                                            txtResp += "Дата ТО: " + txtCHECK_DATE + "\r\n";

                                                            string txtOWNER_ADDR = Convert.ToString(rrow["OWNER_ADDR"]).Trim();
                                                            txtResp += "Адрес владельца: " + txtOWNER_ADDR + "\r\n";

                                                            /* закомментировал т.к. ИНН есть в шапке
                                                            string txtINN = Convert.ToString(rrow["INN"]).Trim();
                                                            if (txtINN.Length >= 10)
                                                                txtResp += "ИНН владельца: " + txtINN + "\r\n";
                                                            */

                                                            if (tbl.Rows.Count > 1) txtResp += "\r\n"; // перевод строки чтобы разделить записи если их больше 1
                                                        }
                                                    }
                                                    else txtResp = "Нет данных";

                                                    txtAnswerType = "01";
                                                    if (txtResp == "Нет данных") txtAnswerType = "02";
                                                    else if (txtResp.Length == 0) txtAnswerType = "03"; //  требует уточнения
                                                    // вставить ответ в ИТ
                                                    // decimal nExtKey = mvv.InsertResponseIntTable         (conPK_OSP, nId, txtResp, DateTime.Today, nStatus, lLogger.logID, txtAgreementCode, txtAgreementCode, txtAgreementCode, txtEntityName, txtAnswerType, lLogger);
                                                    // 20180620 - неверно вставлялся ответ для юлиц и ИП с ИНН - ошибка с внешним ключом была - есть скрипт для исправления
                                                    decimal nExtKey = mvv.InsertResponseIntTableNewExtKey(conPK_OSP, nId, txtResp, DateTime.Today, nStatus, lLogger.logID, txtAgreementCode, txtAgreementCode, txtAgreementCode, txtEntityName, txtAnswerType, lLogger);
                                                    if (nExtKey > 0)
                                                    {
                                                        if (txtAnswerType == "01")
                                                        {
                                                            // пройти по всем строчкам и вставить в таблицу только те, которые актуальны (не сняты с учета)
                                                            if (tbl != null && tbl.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow rrow in tbl.Rows)
                                                                {
                                                                    decimal nExtIdentID = mvv.InsertExtDatumMachineRtnIntTable(conPK_OSP, nExtKey, txtFioD, rrow, lLogger);
                                                                }
                                                            }

                                                        }
                                                        // теперь нужно обозначить запрос в ИТ как выгруженный
                                                        if (mvv.UpdateExtRequestRow(conPK_OSP, row, lLogger))
                                                        {
                                                            lLogger.WriteLLog("Запрос № " + txtId.ToString() + " успешно обработан. Получен ответ № " + nExtKey.ToString() + "\r\n");
                                                            cnt++;
                                                        }
                                                        else
                                                            lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Несмотря на это в базу загружен ответ № " + nExtKey.ToString() + "\r\n");
                                                    }
                                                    else
                                                    {
                                                        lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n");
                                                        lLogger.ErrMessage += "Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n";
                                                    }
                                                }
                                            }
                                        }
                                    } // end_for
                                }
                                if (cnt > 0)
                                {
                                    lLogger.UpdateLLogCount(Convert.ToInt32(cnt));
                                    lLogger.UpdateLLogStatus(2);
                                }
                            }
                            # endregion

                            # region "ГИМС_10_SELFREQUEST"
                            // автоматический ответ на запрос через базу данных

                            if (bRunGimsLodka10_selfrequest)
                            {
                                Int64 cnt = 0;

                                DataTable DT_doc_rtn10 = null;

                                string txtAgreementSber = "ГИМС_ЛОДКА_10";
                                txtAgreementCode = "ГИМС_ЛОДКА_10";

                                string txtAgentCode = "ГИМС_10";
                                string txtAgentDeptCode = "ГИМС_10";

                                Logger_ufssprk_tools lLoggerError = null;
                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 1, txtAgreementSber, 0, nOspNum, "Пакет запросов в ГИМС_ЛОДКА_10.");
                                DT_doc_rtn10 = mvv.ReadGimsLodka10Zapros(conPK_OSP, lLogger);
                                // DT_doc_fiz_sber = mvv.ReadSberZaprosNewFormatTest(conPK_OSP, lLogger);
                                ShowLoggerError(lLogger);

                                //  сразу вопрос - как писать логи
                                //  менять ли тип пакета (сделать специальный packType для таких Самообрабатывающихся запросов)
                                //  или писать сначала лог на выгрузку, а потом лог на вставку - наверное проще так сначала,
                                //  а потом переделать уже на новый тип лога

                                // записать в таблицу что получили строльк-то запросов и все, переходим к обработке - то есть загрузке ответов
                                if (DT_doc_rtn10 != null)
                                {
                                    lLogger.WriteLLog("Выгружено запросов в ГИМС МЧС о зергистрированных судах: " + DT_doc_rtn10.Rows.Count.ToString() + "\n");
                                    lLogger.UpdateLLogCount(DT_doc_rtn10.Rows.Count);
                                    lLogger.UpdateLLogStatus(2);
                                }
                                // после получения таблицы будем делать 

                                // открыть лог ответа
                                nPack_type = 2; // 2 - простой ответ
                                decimal nParentID = lLogger.logID;
                                lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, nParentID, nOspNum, "Пакет загрузки ответов на запрос в ГИМС МЧС о зарегистрированных судах");

                                if (DT_doc_rtn10 != null)
                                {
                                    // написать когда и что делаем
                                    lLogger.WriteLLog(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "Начало обработки запросов по базе ГИМС МЧС. Строк запросов всего: " + DT_doc_rtn10.Rows.Count.ToString() + "\r\n");

                                    OleDbConnection VUcon = new OleDbConnection(txtGibdd10ConString);

                                    // в цикле по всем строкам DT_doc_gibdd10 
                                    string txtId = "";
                                    decimal nId = 0;
                                    decimal nRespId = 0;
                                    decimal nStatus = 19; // ответ получен
                                    string txtAnswerType = "01";
                                    string txtEntityName = "ГИМС МЧС России по Республике Карелия";


                                    foreach (DataRow row in DT_doc_rtn10.Rows)
                                    {
                                        // получить ID запроса
                                        // учитывать что если класс контрагента не физ. лицо - то смотреть на ИНН
                                        // !!! пока временно не смотрим ИНН и работаем только с ФЛ
                                        // классы физ. лиц 2, 71,95,96,97,666
                                        // инд. предприниматель  95
                                        // итого (2,71,95,96,97,666)
                                        txtId = Convert.ToString(row["ZAPROS"]);
                                        Decimal.TryParse(txtId, out nId);
                                        if (nId > 0)
                                        {
                                            string txtFioD = Convert.ToString(row["FIOVK"]).Trim().ToUpper();
                                            // убрать двойные пробелы
                                            txtFioD = RemoveDoubleSpaces(txtFioD, 200);


                                            DateTime dtBornD = DateTime.MinValue;
                                            string txtBornD = "";
                                            string txtDocNum = "";
                                            DateTime dtStart = DateTime.MinValue;
                                            DateTime dtDateVh = DateTime.MinValue;

                                            List<int> fiz = new List<int>(new[] { 2, 71, 95, 96, 97, 666 });
                                            int iDbtrCls = Convert.ToInt32(row["ID_DBTRCLS"]);
                                            // дату рождения проверяем только для физ лиц
                                            // для юр лиц нужно будет проверять ИНН и искать по нему

                                            if (fiz.Contains(iDbtrCls) && !row["DATROZHD"].Equals(DBNull.Value))
                                            {
                                                txtBornD = Convert.ToString(row["DATROZHD"]);
                                                if (DateTime.TryParse(txtBornD, out dtBornD))
                                                {
                                                    // TODO: получить ответ как DataTable и тут уже разбирать параметры и собирать строчку - иначе не вставить структурированные сведения
                                                    // string txtResp = mvv.FindVU(VUcon, txtGibdd10DataBase, txtFioD, dtBornD, lLogger);
                                                    DataTable tbl = mvv.FindGimsLodkaFlDt(VUcon, txtFioD, dtBornD, lLogger);
                                                    string txtResp = "";
                                                    if (tbl != null && tbl.Rows.Count > 0)
                                                    {
                                                        txtResp += "Получены сведения о наличии зарегистрированного за должником " + txtFioD + " (" + dtBornD.ToShortDateString() + " г.р.) имущества.\r\n";
                                                        if (tbl.Rows.Count > 1) txtResp += "Всего объектов имущества: " + tbl.Rows.Count.ToString() + "\r\n";
                                                        if (tbl.Rows.Count > 0)
                                                        {
                                                            string txtAct_date = "";
                                                            if (tbl.Columns.Contains("ACT_DATE")) txtAct_date = Convert.ToString(tbl.Rows[0]["ACT_DATE"]);
                                                            if (txtAct_date.Length > 0)
                                                            {
                                                                txtResp += "Дата актуальности сведений: " + txtAct_date + "\r\n";
                                                            }
                                                        }
                                                        foreach (DataRow rrow in tbl.Rows)
                                                        {
                                                            string txtSb_regn = Convert.ToString(rrow["reg_number"]).Trim();
                                                            txtResp += "Регистрационный номер судна: " + txtSb_regn + "\r\n";

                                                            string txtSud_model = Convert.ToString(rrow["model_name"]).Trim();
                                                            txtResp += "Модель: " + txtSud_model + "\r\n";

                                                            string txtSud_godvyp = Convert.ToString(rrow["made_year"]).Trim();
                                                            txtResp += "Год производства: " + txtSud_godvyp + "\r\n";

                                                            string txtSb_datavyd = Convert.ToString(rrow["register_date"]).Trim();
                                                            txtResp += "Дата регистрации: " + txtSb_datavyd + "\r\n";

                                                            string txtReg_datasnsuchet = Convert.ToString(rrow["unregister_date"]).Trim(); // это дата снятия с учета
                                                            if (txtReg_datasnsuchet.Length > 0)
                                                                txtResp += "Дата снятия с учета: " + txtReg_datasnsuchet + "\r\n";

                                                            if (tbl.Rows.Count > 1) txtResp += "\r\n"; // перевод строки чтобы разделить записи если их больше 1
                                                        }
                                                    }
                                                    else txtResp = "Нет данных";

                                                    txtAnswerType = "01";
                                                    if (txtResp == "Нет данных") txtAnswerType = "02";
                                                    else if (txtResp.Length == 0) txtAnswerType = "03"; //  требует уточнения
                                                    // вставить ответ в ИТ
                                                    // nRespId = mvv.InsertResponseIntTable(conPK_OSP, nId, txtResp, DateTime.Today, nStatus, lLogger.logID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName, txtAnswerType, lLogger);
                                                    decimal nExtKey = mvv.InsertResponseIntTableNewExtKey(conPK_OSP, nId, txtResp, DateTime.Today, nStatus, lLogger.logID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName, txtAnswerType, lLogger);
                                                    if (nExtKey > 0)
                                                    {

                                                        // если ответ был с данными - вставить в ИТ EXT_IDENTIFICATION_DATA структурированные сведения
                                                        if (txtAnswerType == "01")
                                                        {
                                                            // это код 
                                                            // ENTITY_NAME - для доп. сведений это ФИО должника
                                                            // 20160513 - закомментировал т.к. это не работает
                                                            // decimal nExtIdentID = mvv.InsertExtIdentificationData(conPK_OSP, nRespId, txtFioD, dtDateVh, txtDocNum, dtStart, txtFioD, txtTypeDocCode, lLogger);
                                                            if (tbl != null && tbl.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow rrow in tbl.Rows)
                                                                {
                                                                    string txtReg_datasnsuchet = Convert.ToString(rrow["unregister_date"]).Trim();
                                                                    decimal nExtIdentID = 0;
                                                                    // вставлять в базу только те лодки, которые не сняты с учета
                                                                    // остальные пусть останутся только в тексте
                                                                    if (txtReg_datasnsuchet.Length.Equals(0))
                                                                        nExtIdentID = mvv.InsertExtTransportGimsIntTable(conPK_OSP, nExtKey, txtFioD, rrow, lLogger);
                                                                }
                                                            }
                                                        }
                                                        // теперь нужно обозначить запрос в ИТ как выгруженный
                                                        if (mvv.UpdateExtRequestRow(conPK_OSP, row, lLogger))
                                                        {
                                                            lLogger.WriteLLog("Запрос № " + txtId.ToString() + " успешно обработан. Получен ответ № " + nExtKey.ToString() + "\r\n");
                                                            cnt++;
                                                        }
                                                        else
                                                            lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Несмотря на это в базу загружен ответ № " + nExtKey.ToString() + "\r\n");
                                                    }
                                                    else
                                                    {
                                                        lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n");
                                                        lLogger.ErrMessage += "Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n";
                                                    }
                                                }
                                                else
                                                {
                                                    // TODO: вставить ответ что нет даты рождения
                                                    // вставить ответ что нет даты рождения
                                                    decimal nErrId = 0;
                                                    nErrId = mvv.InsertErrorResp(conPK_OSP, constrGIBDD, txtAgentCode, txtAgentDeptCode, txtAgreementCode, nId, txtEntityName, ref lLoggerError, lLogger);
                                                    if (nErrId > 0)
                                                    {
                                                        lLogger.WriteLLog("Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n");
                                                        lLogger.ErrMessage += "Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n";
                                                    }
                                                    else
                                                    {
                                                        lLogger.WriteLLog("Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n");
                                                        lLogger.ErrMessage += "Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n";
                                                    }
                                                }
                                            }
                                            else if (fiz.Contains(iDbtrCls) && row["DATROZHD"].Equals(DBNull.Value))
                                            {
                                                // вставить ответ что нет даты рождения
                                                decimal nErrId = 0;
                                                nErrId = mvv.InsertErrorResp(conPK_OSP, constrGIBDD, txtAgentCode, txtAgentDeptCode, txtAgreementCode, nId, txtEntityName, ref lLoggerError, lLogger);
                                                if (nErrId > 0)
                                                {
                                                    lLogger.WriteLLog("Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n");
                                                    lLogger.ErrMessage += "Ошибка! На запрос № " + txtId.ToString() + " вставлен ответ - не заполнена дата рождения. id ответа = " + nErrId.ToString() + "\r\n";
                                                }
                                                else
                                                {
                                                    lLogger.WriteLLog("Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n");
                                                    lLogger.ErrMessage += "Ошибка! В запросе № " + txtId.ToString() + " не заполнена дата рождения. Ответ с таким сообщение не удалось вставить в базу данных.\r\n";
                                                }
                                            }

                                            /* работаем с юр. лицами по ИНН */
                                            else
                                            {
                                                string txtInn = Convert.ToString(row["id_dbtr_inn"]).Trim();
                                                // 95 - это ИП - он и ФИЗ и ЮР - если нет даты рождения - будем искать по ИНН
                                                if ((!fiz.Contains(iDbtrCls) || iDbtrCls.Equals(95)) && txtInn.Length >= 10)
                                                {
                                                    // ищем как юр.лицо по ИНН
                                                    DataTable tbl = mvv.FindGimsLodkaUlDt(VUcon, txtFioD, txtInn, lLogger);
                                                    string txtResp = "";
                                                    if (tbl != null && tbl.Rows.Count > 0)
                                                    {
                                                        txtResp += "Получены сведения о наличии зарегистрированного за должником " + txtFioD + " (поиск велся по ИНН " + txtInn + ") имущества.\r\n";
                                                        if (tbl.Rows.Count > 1) txtResp += "Всего объектов имущества: " + tbl.Rows.Count.ToString() + "\r\n";
                                                        if (tbl.Rows.Count > 0)
                                                        {
                                                            string txtAct_date = "";
                                                            if (tbl.Columns.Contains("ACT_DATE")) txtAct_date = Convert.ToString(tbl.Rows[0]["ACT_DATE"]);
                                                            if (txtAct_date.Length > 0)
                                                            {
                                                                txtResp += "Дата актуальности сведений: " + txtAct_date + "\r\n";
                                                            }
                                                        }
                                                        foreach (DataRow rrow in tbl.Rows)
                                                        {
                                                            string txtSb_regn = Convert.ToString(rrow["reg_number"]).Trim();
                                                            txtResp += "Регистрационный номер судна: " + txtSb_regn + "\r\n";

                                                            string txtSud_model = Convert.ToString(rrow["model_name"]).Trim();
                                                            txtResp += "Модель: " + txtSud_model + "\r\n";

                                                            string txtSud_godvyp = Convert.ToString(rrow["made_year"]).Trim();
                                                            txtResp += "Год производства: " + txtSud_godvyp + "\r\n";

                                                            string txtSb_datavyd = Convert.ToString(rrow["register_date"]).Trim();
                                                            txtResp += "Дата регистрации: " + txtSb_datavyd + "\r\n";

                                                            string txtReg_datasnsuchet = Convert.ToString(rrow["unregister_date"]).Trim(); // это дата снятия с учета
                                                            if (txtReg_datasnsuchet.Length > 0)
                                                                txtResp += "Дата снятия с учета: " + txtReg_datasnsuchet + "\r\n";

                                                            if (tbl.Rows.Count > 1) txtResp += "\r\n"; // перевод строки чтобы разделить записи если их больше 1

                                                        }
                                                    }

                                                    else txtResp = "Нет данных";

                                                    txtAnswerType = "01";
                                                    if (txtResp == "Нет данных") txtAnswerType = "02";
                                                    else if (txtResp.Length == 0) txtAnswerType = "03"; //  требует уточнения
                                                    // вставить ответ в ИТ
                                                    decimal nExtKey = mvv.InsertResponseIntTableNewExtKey(conPK_OSP, nId, txtResp, DateTime.Today, nStatus, lLogger.logID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName, txtAnswerType, lLogger);
                                                    if (nExtKey > 0)
                                                    {

                                                        // если ответ был с данными - вставить в ИТ EXT_TRANSPORT_DATA структурированные сведения
                                                        if (txtAnswerType == "01")
                                                        {
                                                            if (tbl != null && tbl.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow rrow in tbl.Rows)
                                                                {
                                                                    string txtReg_datasnsuchet = Convert.ToString(rrow["unregister_date"]).Trim();
                                                                    decimal nExtIdentID = 0;
                                                                    // вставлять в базу только те лодки, которые не сняты с учета
                                                                    // остальные пусть останутся только в тексте
                                                                    if (txtReg_datasnsuchet.Length.Equals(0))
                                                                        nExtIdentID = mvv.InsertExtTransportGimsIntTable(conPK_OSP, nExtKey, txtFioD, rrow, lLogger);
                                                                }
                                                            }
                                                        }
                                                        // теперь нужно обозначить запрос в ИТ как выгруженный
                                                        if (mvv.UpdateExtRequestRow(conPK_OSP, row, lLogger))
                                                        {
                                                            lLogger.WriteLLog("Запрос № " + txtId.ToString() + " успешно обработан. Получен ответ № " + nExtKey.ToString() + "\r\n");
                                                            cnt++;
                                                        }
                                                        else
                                                            lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Несмотря на это в базу загружен ответ № " + nExtKey.ToString() + "\r\n");
                                                    }
                                                    else
                                                    {
                                                        lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n");
                                                        lLogger.ErrMessage += "Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n";
                                                    }
                                                }
                                            }
                                        }
                                    } // end_for
                                }
                                if (cnt > 0)
                                {
                                    lLogger.UpdateLLogCount(Convert.ToInt32(cnt));
                                    lLogger.UpdateLLogStatus(2);
                                }
                            }
                            # endregion

                            # region "RunCredOrgReqOut"
                            if (bRunCredOrgReqOut)
                            {

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;
                                // сейчас выгружается каждый раз при запуске
                                // предлагаю выгружать только в понедельник

                                // день недели от 1 до 7
                                int day = ((int)DateTime.Now.DayOfWeek == 0) ? 7 : (int)DateTime.Now.DayOfWeek;
                                if (txtZaprosOut.Equals("1")) day = 1;
                                // добавить проверку что если вчера был понедельник, и не было выгрузки - то тоже выгружать
                                // проверять через запрос к ufssprk-tools по коду соглашения и типу пакета
                                int nTomorrowUnloaded = 0;
                                string txtAgreementSber = "ALL_CRED_ORG";
                                if (day.Equals(2))
                                {
                                    // срабатывает потому что статус не 2, а 1 - нужно сделать ниже UpdateOackStatus(2) если pack_count > 0
                                    nTomorrowUnloaded = mvv.GetDayPackCount(constrGIBDD, 1, txtAgreementSber, DateTime.Today.AddDays(-1), lCommonLogger);
                                    if (nTomorrowUnloaded.Equals(0)) day = 1;
                                }



                                if ((day == 1) || (DateTime.Today.ToShortDateString().Equals("21.12.2017")) || (DateTime.Today.ToShortDateString().Equals("22.12.2017")))
                                // || (DateTime.Today.ToShortDateString().Equals("26.05.2016")))
                                {
                                    Int64 cnt = 0;
                                    DataTable DT_doc_jur;
                                    DataTable DT_doc_fiz;
                                    decimal nOrg_id = 0;
                                    bool bDateFolderAdd = true; // укладывать файлы в папки с именем по текущей дате

                                    //txtAgreementCode = "170";
                                    txtAgreementCode = "190";
                                    txtAgreementSber = "ALL_CRED_ORG";

                                    // особенность в том, что при выгрузке формируется столько логов - сколько id в списке банков
                                    // сейчас их должно быть всего 2
                                    // поэтому лог вобще не должен быть привязан к соглашению - он о другом - о выгрузке всех запросов
                                    // придумаем для этого свой код - ALL_CRED_ORG
                                    Logger_ufssprk_tools lLoggerError = null;

                                    Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 1, txtAgreementSber, 0, nOspNum, "Пакет запросов в БАНКИ.");
                                    DT_doc_jur = mvv.GetDataTableFromFB(conPK_OSP, "select 190 agreement_id, ext_request_id,  pack_id,  1 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID where mvv_agreement_code = '190' and processed = 0 and (req.ID_DBTRCLS = 1 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 1)))", "TOFIND", lLogger);
                                    ShowLoggerError(lLogger);

                                    DT_doc_fiz = mvv.GetDataTableFromFB(conPK_OSP, "select 190 agreement_id, ext_request_id,  pack_id,  2 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID where mvv_agreement_code = '190' and processed = 0 and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))", "TOFIND", lLogger);
                                    ShowLoggerError(lLogger);

                                    // получить номер ОСП
                                    // decimal nOspNum = mvv.GetOSP_Num(con, out txtErrString); - уже есть
                                    int iDocCnt = 0;
                                    if (DT_doc_jur != null) iDocCnt += DT_doc_jur.Rows.Count;
                                    if (DT_doc_fiz != null) iDocCnt += DT_doc_fiz.Rows.Count;

                                    Int64 iCnt;
                                    string fullpath = "";

                                    if (bDateFolderAdd)
                                    {
                                        // CreatePathWithDate(cred_org_path);
                                        fullpath = ff.CreatePathWithDateS(txtCredOrgReq_Path_out);
                                    }
                                    else
                                    {
                                        if (!Directory.Exists(txtCredOrgReq_Path_out))
                                            Directory.CreateDirectory(txtCredOrgReq_Path_out);
                                        fullpath = txtCredOrgReq_Path_out;

                                    }
                                    DateTime DatZapr1;
                                    DateTime DatZapr2;
                                    //string txtCistomLegalIds = "86200999999007, 86200999999009, 86200999999010, 86200999999011, 86200999999013";

                                    String[] Legal_List;
                                    String[] Legal_Name_List;
                                    String[] Legal_Сonv_List;

                                    Legal_List = txtCistomLegalIds.Split(',');

                                    Legal_Name_List = (String[])Legal_List.Clone();
                                    Legal_Сonv_List = (String[])Legal_List.Clone();
                                    //int code = 0;
                                    Decimal code = 0;
                                    //for (int i = 1; i < Legal_Name_List.Length; i++)
                                    for (int i = 0; i < Legal_Name_List.Length; i++)
                                    {
                                        //code = Convert.ToInt32((Legal_List[i]).Trim());
                                        code = Convert.ToDecimal((Legal_List[i]).Trim());
                                        Legal_Name_List[i] = mvv.GetLegal_Name(code, conPK_OSP, lLogger);
                                        Legal_Сonv_List[i] = mvv.GetLegal_Conv(code, conPK_OSP, lLogger);

                                    }

                                    DatZapr1 = (DateTime.Today).AddDays(-7);
                                    DatZapr2 = DateTime.Today;


                                    string txtFileName = "tofind" + nOspNum.ToString().PadLeft(2, '0') + ".dbf";
                                    cnt = mvv.WriteToDBF(true, fullpath, "tofind1.dbf", DatZapr1, DatZapr2, txtFileName, constrGIBDD, txtConStrPKOSP, DT_doc_fiz, DT_doc_jur, Legal_List, lLogger);
                                    // cnt = WriteToDBF(true, fullpath, "tofind1.dbf", DatZapr1, DatZapr2, "tofind" + nOspNum.ToString().PadLeft(2, '0') + ".dbf");

                                    // есть ли в логе информация о результатах?
                                    // как минимум нужно время окончания работы зафиксировать
                                    lLogger.UpdateLLogCount(Convert.ToInt32(cnt));
                                    // сделать UpdateLLogStatus(2) если pack_count > 0
                                    if (cnt > 0) lLogger.UpdateLLogStatus(2);

                                    lLogger.WriteLLog(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + " Всего в файл выгружено запросов: " + cnt.ToString() + "\n");
                                    ShowLoggerError(lLogger);

                                    // Пишем логи и отправляем e-mail
                                    Console.WriteLine("Итого выгружено запросов в кред. организации: " + cnt.ToString());

                                    // ShowLoggerError(lLogger);
                                    if (lLogger.ErrMessage.Length > 0)
                                    {
                                        // если есть лог с ошибками - отправить соответствующий отчет
                                        SendEmail(lLogger.ErrMessage, "Ошибки при выгрузке запросов в кред. организации в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                        lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                        lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                        ShowLoggerError(lLogger);
                                    }

                                    lLogger.WriteLLog("Выгрузка окончена. Всего выгружено запросов: " + cnt.ToString());

                                    // отправить e-mail
                                    string txtMessage = DateTime.Now.ToString() + " Выгружено " + cnt.ToString() + " в пакет запросов: " + txtFileName;
                                    // записать в лог автозагрузки сколько строк загружено (txtMessage)
                                    lLogger.WriteLLog("\n" + txtMessage);
                                    ff.WriteTofile(txtMessage, txtLogFileName);
                                    // в ОСП не отправляем
                                    // ff.SendEmail(txtMessage, "Загрузка реестра оплаченных штрафов МВД", txtOspEmail, txtAdminEmail, txtMailServ, "");
                                    // email админу
                                    ff.SendEmail(txtMessage, "Выгрузка запросов в кред. организации в ОСП " + iDiv.ToString().PadLeft(2, '0'), txtAdminEmail, txtAdminEmail, txtMailServ, "");
                                    // email в ОСП
                                    ff.SendEmail(txtMessage, "Выгрузка запросов в кред. организации в ОСП " + iDiv.ToString().PadLeft(2, '0'), txtOspEmail, txtAdminEmail, txtMailServ, "");

                                    /*
                                    if (cnt > 0)
                                    {
                                        lLogger.UpdateLLogCount(Convert.ToInt32(cnt));
                                        lLogger.UpdateLLogStatus(2);
                                        lLogger.UpdateLLogFileName(txtFileName);
                                    }
                                    */
                                }
                            }
                            # endregion

                            # region "bRunCredOrgAnsIn"
                            if (bRunCredOrgAnsIn)
                            {
                                String[] Legal_List;
                                String[] Legal_Name_List;
                                Legal_List = txtCistomLegalIds.Split(',');
                                Legal_Name_List = (String[])Legal_List.Clone();

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;


                                decimal nOrg_id;
                                decimal nAgr_id;

                                nPack_type = 2; // 2 - простой ответ
                                txtAgreementCode = "ALL_CRED_ORG";

                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка ответов из кредитных организаций.");

                                for (int i = 0; i < Legal_Name_List.Length; i++)
                                {
                                    //code = Convert.ToInt32((Legal_List[i]).Trim());
                                    decimal code = Convert.ToDecimal((Legal_List[i]).Trim());
                                    Legal_Name_List[i] = mvv.GetLegal_Name(code, conPK_OSP, lLogger);

                                }

                                foreach (string txtOrg_id in Legal_List)
                                {
                                    // по коду организации узнать номер согашения
                                    nOrg_id = Convert.ToDecimal(txtOrg_id);
                                    // получить Agreement_ID по номеру контрагента

                                    nAgr_id = mvv.GetAgr_by_Org(conPK_OSP, nOrg_id, lLogger);
                                    txtAgreementCode = mvv.GetAgreement_Code(conPK_OSP, nAgr_id, lLogger);
                                    // вопрос - нужен ли lLogger2? может писать все в один лог?
                                    // лучше в 1 лог, я думаю
                                    // Logger_ufssprk_tools lLogger2 = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка ответов из банка.");
                                    string txtBankOtvetPath = "";
                                    // строчку пути для банка получить из Б.Д.
                                    txtBankOtvetPath = mvv.GetCredOrgPath(constrGIBDD, txtAgreementCode, lLogger);
                                    iTotalFiles += mvv.AutoLoadCredOrgOtvet(constrGIBDD, txtConStrPKOSP, txtBankOtvetPath, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, txtAgreementCode, nPack_type, nOrg_id, lLogger);
                                }



                                // берем скелет из IC_MVD_IN
                                // nPack_type = 2; // 2 - простой ответ

                                // Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка ответов из ИЦ МВД (по требованиям).");
                                // Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка ответов из кредитных организаций.");
                                // lLogger.ErrMessage += txtErr;
                                //ShowLoggerError(lLogger); // вывод сообщения об ошибке, если она вдруг случилась внутри вызванной функции


                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                //iTotalFiles = mvv.AutoLoadSberReport(constrGIBDD, txtConStrPKOSP, txtUploadDirSberReport, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                // iTotalFiles = mvv.AutoLoadIcMvdOtvet(constrGIBDD, txtConStrPKOSP, txtIC_MVD_Path_in, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, txtAgreementCode, nPack_type, lLogger);

                                lLogger.UpdateLLogCount(iTotalFiles);
                                lLogger.UpdateLLogStatus(2);

                                // ShowLoggerError(lLogger);
                                if (lLogger.ErrMessage.Length > 0)
                                {
                                    // если есть лог с ошибками - отправить соответствующий отчет
                                    SendEmail(lLogger.ErrMessage, "Ошибки при загрузке ответов из кредитных организаций в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                    lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                    lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                    ShowLoggerError(lLogger);
                                }

                                Console.WriteLine("Итого загружено ответов из банков: " + iTotalFiles.ToString());
                            }
                            # endregion

                            # region "RunPotdOut"
                            if (bRunPotdOut)
                            {
                                // сейчас выгружается каждый раз при запуске
                                // предлагаю выгружать только в понедельник

                                // день недели от 1 до 7
                                int day = ((int)DateTime.Now.DayOfWeek == 0) ? 7 : (int)DateTime.Now.DayOfWeek;
                                if (txtZaprosOut.Equals("1")) day = 1;
                                // добавить проверку что если вчера был понедельник, и не было выгрузки - то тоже выгружать
                                // проверять через запрос к ufssprk-tools по коду соглашения и типу пакета
                                int nTomorrowUnloaded = 0;
                                txtAgreementCode = "110";
                                if (day.Equals(2))
                                {
                                    nTomorrowUnloaded = mvv.GetDayPackCount(constrGIBDD, 1, txtAgreementCode, DateTime.Today.AddDays(-1), lCommonLogger);
                                    if (nTomorrowUnloaded.Equals(0)) day = 1;
                                }

                                if (day == 1)
                                // || (DateTime.Today.ToShortDateString().Equals("26.05.2016")))
                                {
                                    string txtOspEmail = "";
                                    string txtAdminEmail = "";

                                    if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                    if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                    else txtOspEmail = txtAdminEmail;

                                    Int64 cnt = 0;
                                    DataTable DT_potd_doc;
                                    decimal nOrg_id = 0;
                                    bool bDateFolderAdd = true; // укладывать файлы в папки с именем по текущей дате

                                    txtAgreementCode = "110";
                                    string txtAgreementSber = "110";

                                    // особенность в том, что при выгрузке формируется столько логов - сколько id в списке банков
                                    // сейчас их должно быть всего 2
                                    // поэтому лог вобще не должен быть привязан к соглашению - он о другом - о выгрузке всех запросов
                                    // придумаем для этого свой код - ALL_CRED_ORG
                                    Logger_ufssprk_tools lLoggerError = null;
                                    Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 1, "110", 0, nOspNum, "Выгрузка запросов в ПФ о перс. учете (месте работы) в формате XML");
                                    // DT_doc_jur = mvv.GetDataTableFromFB(conPK_OSP, "select 170 agreement_id, ext_request_id,  pack_id,  1 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID where mvv_agreement_code = '170' and processed = 0 and (req.ID_DBTRCLS = 1 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 1)))", "TOFIND", lLogger);
                                    
                                    //DT_potd_doc = mvv.GetDataTableFromFB(conPK_OSP, "select 110 agreement_id, ext_request_id,  pack_id,  2 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK, ip.ID_DBTR_BORNADR, COALESCE(ext_request.debtor_birthplace, '') as debtor_birthplace, ext_request.spi_id, ext_request.fio_spi, ext_request.h_spi, ext_request.ip_id, ext_request.ip_risedate, ext_request.id_type, ext_request.id_number, ext_request.id_date, ext_request.id_subject_type, ext_request.ip_sum as id_sum, ext_request.ip_rest_debtsum, COALESCE(d.ser_doc,'') as SerDoc, COALESCE(d.num_doc,'') as NumDoc,   coalesce(d.date_doc, '01.01.1900') DateDoc, COALESCE(d.code_dep, '') as CodeDep,    COALESCE(d.type_doc_code, 0) as TypeDoc, COALESCE(d.fio_doc, '') as FioDoc from ext_request join o_ip req on ext_request.req_id = req.id join DOC_IP_DOC ip on ip.id = req.ip_id join SPI on ext_request.spi_id = spi.SUSER_ID    join o_ip_req_ip on o_ip_req_ip.id = ext_request.req_id  left join Mvv_Datum_Identificator d on d.id = o_ip_req_ip.DATUM_DOCUMENT_ID where mvv_agreement_code = '110' and processed = 0 and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2))) order by FIOVK, GOD, DATROZHD, IPNO_NUM, ZAPROS", "TOFIND", lLogger);
                                    // добавил в выборку вместо IP.ID_DBTR_BORNADR  EXT_REQUEST.debtor_birthplace
                                    DT_potd_doc = mvv.GetDataTableFromFB(conPK_OSP, "select 110 agreement_id, ext_request_id,  pack_id,  2 LITZDOLG, req_id ZAPROS, req.IPNO_NUM, req.DIV, debtor_name FIOVK, ip_num NOMIP, spi.spi_zonenum NOMSPI, req_date DATZAPR, debtor_address ADDR,  debtor_birthdate DATROZHD,   req.ID_DBTRCLS, req.DBTR_BORN_YEAR GOD,   ip_sum SUMMA, debtor_inn INNORG, req_number DOC_NUMBER, id_subject_type VIDVZISK, EXT_REQUEST.debtor_birthplace as ID_DBTR_BORNADR, COALESCE(ext_request.debtor_birthplace, '') as debtor_birthplace, ext_request.spi_id, ext_request.fio_spi, ext_request.h_spi, ext_request.ip_id, ext_request.ip_risedate, ext_request.id_type, ext_request.id_number, ext_request.id_date, ext_request.id_subject_type, ext_request.ip_sum as id_sum, ext_request.ip_rest_debtsum, COALESCE(d.ser_doc,'') as SerDoc, COALESCE(d.num_doc,'') as NumDoc,   coalesce(d.date_doc, '01.01.1900') DateDoc, COALESCE(d.code_dep, '') as CodeDep,    COALESCE(d.type_doc_code, 0) as TypeDoc, COALESCE(d.fio_doc, '') as FioDoc from ext_request join o_ip req on ext_request.req_id = req.id join SPI on ext_request.spi_id = spi.SUSER_ID    join o_ip_req_ip on o_ip_req_ip.id = ext_request.req_id  left join Mvv_Datum_Identificator d on d.id = o_ip_req_ip.DATUM_DOCUMENT_ID where mvv_agreement_code = '110' and processed = 0 and (req.ID_DBTRCLS = 2 or (req.ID_DBTRCLS in (select ncc_id from V_COUNTERPARTY_CLS_PARENT where ncc_parent_id = 2)))  order by FIOVK, GOD, DATROZHD, IPNO_NUM, ZAPROS", "TOFIND", lLogger);
                                    
                                    ShowLoggerError(lLogger);

                                    // получить номер ОСП
                                    // decimal nOspNum = mvv.GetOSP_Num(con, out txtErrString); - уже есть
                                    int iDocCnt = 0;
                                    if (DT_potd_doc != null) iDocCnt += DT_potd_doc.Rows.Count;

                                    Int64 iCnt;
                                    string fullpath = "";

                                    if (bDateFolderAdd)
                                    {
                                        // CreatePathWithDate(cred_org_path);
                                        fullpath = ff.CreatePathWithDateS(txtPotdOutPath);
                                    }
                                    else
                                    {
                                        if (!Directory.Exists(txtPotdOutPath))
                                            Directory.CreateDirectory(txtPotdOutPath);
                                        fullpath = txtPotdOutPath;

                                    }


                                    //cnt = mvv.WriteToDBF(true, fullpath, "tofind1.dbf", DatZapr1, DatZapr2, "tofind" + nOspNum.ToString().PadLeft(2, '0') + ".dbf", constrGIBDD, txtConStrPKOSP, DT_doc_fiz, DT_doc_jur, Legal_List, lLogger);
                                    int nPackNum = 0;
                                    nPackNum = mvv.GetDayPackCount(constrGIBDD, 1, "110", DateTime.Today, lLogger);
                                    //    GetPackCount(1, "110", DateTime.Today.Year) + 1;

                                    string txtFileName = ff.makenewPersPensXMLFileName(nOspNum, 9, 0, nPackNum + 1);
                                    cnt = mvv.WritePotdToXML(fullpath, txtFileName, DT_potd_doc, lLogger.OspNum, "10", "00", "009", lLogger);

                                    if (DT_potd_doc != null && DT_potd_doc.Rows.Count > 0)
                                    {
                                        foreach (DataRow row in DT_potd_doc.Rows)
                                        {
                                            // тут меняется флаг выгрузки
                                            mvv.UpdateExtRequestRow(conPK_OSP, row, lLogger);
                                        }
                                    }

                                    if (cnt > 0)
                                    {
                                        lLogger.UpdateLLogCount(Convert.ToInt32(cnt));
                                        lLogger.UpdateLLogStatus(2);
                                        lLogger.UpdateLLogFileName(txtFileName);
                                    }


                                    // MessageBox.Show("Выгрузка окончена. Всего выгружено запросов: " + cnt.ToString(), "Внимание!", MessageBoxButtons.OK);
                                    Console.WriteLine("Итого выгружено запросов в ПФ расширенных: " + cnt.ToString());

                                    // ShowLoggerError(lLogger);
                                    if (lLogger.ErrMessage.Length > 0)
                                    {
                                        // если есть лог с ошибками - отправить соответствующий отчет
                                        SendEmail(lLogger.ErrMessage, "Ошибки при выгрузке запросов в ПФ расширенный (XML) в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                        lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                        lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                        ShowLoggerError(lLogger);
                                    }

                                    lLogger.WriteLLog("Выгрузка окончена. Всего выгружено запросов: " + cnt.ToString());

                                    // отправить e-mail
                                    string txtMessage = DateTime.Now.ToString() + " Выгружено " + cnt.ToString() + " в пакет запросов: " + txtFileName;
                                    // записать в лог автозагрузки сколько строк загружено (txtMessage)
                                    lLogger.WriteLLog("\n" + txtMessage);
                                    ff.WriteTofile(txtMessage, txtLogFileName);
                                    // в ОСП не отправляем
                                    // ff.SendEmail(txtMessage, "Загрузка реестра оплаченных штрафов МВД", txtOspEmail, txtAdminEmail, txtMailServ, "");
                                    ff.SendEmail(txtMessage, "Выгрузка запросов в ПФ расширенных (XML) в ОСП " + iDiv.ToString().PadLeft(2, '0'), txtAdminEmail, txtAdminEmail, txtMailServ, "");
                                    ff.SendEmail(txtMessage, "Выгрузка запросов в ПФ расширенных (XML) в ОСП " + iDiv.ToString().PadLeft(2, '0'), txtOspEmail, txtAdminEmail, txtMailServ, "");


                                    /*
                                    if (cnt > 0)
                                    {
                                        lLogger.UpdateLLogCount(Convert.ToInt32(cnt));
                                        lLogger.UpdateLLogStatus(2);
                                        lLogger.UpdateLLogFileName(txtFileName);
                                    }
                                    */
                                }
                            }
                            # endregion

                            #region bRunPotdIn
                            // провести автозагрузку ответов из ИЦ МВД
                            // 20160331 - сделать разбивку ответа с учетом txtListID - разделитель ;
                            if (bRunPotdIn)
                            {
                                txtAgreementCode = "110";
                                nPack_type = 2; // 2 - простой ответ


                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка ответов из ПФ расширенных (XML формат).");
                                lLogger.ErrMessage += txtErr;
                                //ShowLoggerError(lLogger); // вывод сообщения об ошибке, если она вдруг случилась внутри вызванной функции

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                //iTotalFiles = mvv.AutoLoadSberReport(constrGIBDD, txtConStrPKOSP, txtUploadDirSberReport, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                iTotalFiles = mvv.AutoLoadPotdOtvet(constrGIBDD, txtConStrPKOSP, txtPotdInPath, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, txtAgreementCode, nPack_type, lLogger);

                                lLogger.UpdateLLogCount(iTotalFiles);
                                //lLogger.UpdateLLogStatus

                                // ShowLoggerError(lLogger);
                                if (lLogger.ErrMessage.Length > 0)
                                {
                                    // если есть лог с ошибками - отправить соответствующий отчет
                                    SendEmail(lLogger.ErrMessage, "Ошибка загрузки ответов ПФ расширенные (XML) в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                    lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                    lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                    ShowLoggerError(lLogger);
                                }

                                Console.WriteLine("Итого загружено ответов ПФ расширенные (XML): " + iTotalFiles.ToString());
                            }

                            #endregion

                            # region "RunPotdOut2017"
                            if (bRunPotdOut2017)
                            {
                                if ((DateTime.Today.ToShortDateString().Equals("04.10.2017"))
                                    || (iDiv == 5 && DateTime.Today.ToShortDateString().Equals("05.10.2017"))
                                    || (iDiv == 12 && DateTime.Today.ToShortDateString().Equals("05.10.2017")))
                                {
                                    Int64 cnt = 0;

                                    DataTable DT_doc_gibdd10 = null;
                                    txtAgreementCode = "ПФР_10";
                                    string txtAgreementSber = "ПФР_10";

                                    Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 1, txtAgreementSber, 0, nOspNum, "Выгрузка запросов в ПФР_10 о перс. учете (месте работы).");
                                    DT_doc_gibdd10 = mvv.ReadPotdOut2017(conPK_OSP, lLogger);
                                    ShowLoggerError(lLogger);

                                    // TODO: посмотреть по логам какой по счету файл за сегодня мы выгружаем
                                    int nFileNum = 0;
                                    //// packType 1 - обычный запрос
                                    nFileNum = mvv.GetDayPackCount(constrGIBDD, 1, txtAgreementSber, DateTime.Today, lLogger); //  +1;  - тут нумерация с 0
                                    ShowLoggerError(lLogger);

                                    //string txtSberXmlFileName = ff.makenewSberFileName() + '.' + ff.makenewSberFileExt(nFileNum, nOspNum);
                                    // zp_yyyymmdd_ХХ.txt, где:
                                    string txtFileName = "2017_m_rab";
                                    txtFileName += DateTime.Today.Year.ToString("D4"); // yyyy
                                    txtFileName += DateTime.Today.Month.ToString("D2").PadLeft(2, '0'); // mm
                                    txtFileName += DateTime.Today.Day.ToString("D2").PadLeft(2, '0'); // dd
                                    txtFileName += "_" + nOspNum.ToString().PadLeft(2, '0');
                                    //txtFileName += nFileNum.ToString("D5").PadLeft(5, '0');// nnnNN – порядковый номер электронного сообщения за указанный день (00001-99999).
                                    txtFileName += ".txt";

                                    // путь для файла-запроса в Сбер - путь из настроек + дата за сегодня
                                    string txtSberFolder = ff.CreatePathWithDateS(txtPotdOutPath); // файлы будем выгружать в папку где запросы расширенные в ПФ
                                    //cnt = mvv.WriteToXML(constrGIBDD, txtConStrPKOSP, txtSberFolder, txtSberXmlFileName, DT_doc_fiz_sber, nOspNum, lLogger);
                                    cnt = 0;

                                    foreach (DataRow row in DT_doc_gibdd10.Rows)
                                    {
                                        string txtText = Convert.ToString(row["ZAPROS"]).Trim().Replace(';', ',');
                                        txtText += ";" + Convert.ToString(row["SNILS"]).Trim().Replace(';', ',');
                                        txtText += ";" + Convert.ToString(row["debtor_surname"]).Trim().Replace(';', ',');
                                        txtText += ";" + Convert.ToString(row["debtor_firstname"]).Trim().Replace(';', ',');
                                        txtText += ";" + Convert.ToString(row["debtor_patronymic"]).Trim().Replace(';', ',');
                                        txtText += ";" + Convert.ToDateTime(row["DATROZHD"]).ToShortDateString().Replace(';', ',');
                                        txtText += ";" + Convert.ToString(row["INN"]).Trim().Replace(';', ',');


                                        if (ff.WriteTofile(txtText, string.Format(@"{0}\{1}", txtSberFolder, txtFileName), Encoding.GetEncoding(1251)))
                                        {
                                            lLogger.MemoryLLog("Обработан запрос # " + cnt.ToString() + " request_id = " + Convert.ToString(row["ZAPROS"]).Trim().ToString() + "\n");
                                            cnt++;
                                        }
                                        else
                                        {
                                            lLogger.MemoryLLog("Не удалось записать в файл запрос # " + cnt.ToString() + " request_id = " + Convert.ToString(row["ZAPROS"]).Trim().ToString() + "\n");
                                            row["GOD"] = -1;
                                        }
                                    }

                                    lLogger.WriteLLog("Выгрузка окончена, обработано строк: " + cnt.ToString() + "\n");

                                    ShowLoggerError(lLogger);
                                    if (DT_doc_gibdd10 != null)
                                    {
                                        foreach (DataRow row in DT_doc_gibdd10.Rows)// select только ради сортировки
                                        {
                                            //UpdatePackRequest(row);
                                            //UpdateKredOrgRequest(row);

                                            mvv.UpdateExtRequestRow(conPK_OSP, row, lLogger);
                                            //UpdateExtRequestThrowLegalList(row);

                                            //UpdateExtRequestRow(row);
                                            //prbWritingDBF.PerformStep();
                                            //prbWritingDBF.Refresh();
                                            //System.Windows.Forms.Application.DoEvents();
                                        }
                                    }
                                    if (cnt > 0)
                                    {
                                        lLogger.UpdateLLogCount(Convert.ToInt32(cnt));
                                        lLogger.UpdateLLogStatus(2);
                                        lLogger.UpdateLLogFileName(txtFileName);
                                    }

                                }

                            }
                            # endregion

                            #region "RunPotdIn2017"
                            // провести автозагрузку ответов из ИЦ МВД
                            // 20160331 - сделать разбивку ответа с учетом txtListID - разделитель ;
                            if (bRunPotdIn2017)
                            {
                                txtAgreementCode = "ПФР_10";
                                nPack_type = 2; // 2 - простой ответ


                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка ответов из ПФР_10 (место работы 2017).");
                                lLogger.ErrMessage += txtErr;
                                //ShowLoggerError(lLogger); // вывод сообщения об ошибке, если она вдруг случилась внутри вызванной функции

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                //iTotalFiles = mvv.AutoLoadSberReport(constrGIBDD, txtConStrPKOSP, txtUploadDirSberReport, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);

                                iTotalFiles = mvv.AutoLoad_PFR10_Otvet(constrGIBDD, txtConStrPKOSP, txtConStrED, txtPotdInPath, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, txtAgreementCode, nPack_type, lLogger);


                                lLogger.UpdateLLogCount(iTotalFiles);


                                // ShowLoggerError(lLogger);
                                if (lLogger.ErrMessage.Length > 0)
                                {
                                    // если есть лог с ошибками - отправить соответствующий отчет
                                    SendEmail(lLogger.ErrMessage, "Ошибки при загрузке ответов от ПФР_10 2017 в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                    lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                    lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                    ShowLoggerError(lLogger);
                                }

                                Console.WriteLine("Итого загружено ответов из ПФР_10 2017: " + iTotalFiles.ToString());
                            }

                            #endregion

                            #region RunPfrRepIn
                            // провести автозагрузку отчетов о взятии в обработку постановлений ПИЭВ, направленных в ПФР
                            if (bRunPfrRepIn)
                            {
                                txtAgreementCode = "ПФР_10_ПИЭВ";
                                //string txtPfrAgentCOde = "ПФР_10";
                                nPack_type = 15; // Загрузка уведомлений о принятии в обработку из Сбербанка
                                int nlogPack_type = 15; // общий лог автозагрузки Автозагрузка отчетов о взятии в обработку постановлений ПИЭВ, направленных в ПФ

                                // этот лог основной, отдельно для каждого файла будет свой лог, но он будет только для учета имени файла
                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nlogPack_type, txtAgreementCode, 0, nOspNum, "Автозагрузка отчетов о взятии в обработку постановлений ПИЭВ, направленных в ПФ.");
                                lLogger.ErrMessage += txtErr;


                                //ShowLoggerError(lLogger); // вывод сообщения об ошибке, если она вдруг случилась внутри вызванной функции

                                string txtOspEmail = "";
                                string txtAdminEmail = "";

                                if (OspEmails.ContainsKey(0)) txtAdminEmail = OspEmails[0];

                                if (OspEmails.ContainsKey(iDiv)) txtOspEmail = OspEmails[iDiv];
                                else txtOspEmail = txtAdminEmail;

                                // iTotalFiles = mvv.AutoLoadGibddReestrs(constrGIBDD, txtUploadDirGibdd, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                //iTotalFiles = mvv.AutoLoadSberReport(constrGIBDD, txtConStrPKOSP, txtUploadDirSberReport, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);
                                //iTotalFiles = mvv.AutoLoadIcMvdOtvet(constrGIBDD, txtConStrPKOSP, txtIC_MVD_Path_in, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, txtAgreementCode, nPack_type, lLogger);
                                iTotalFiles = mvv.AutoLoadPfrRepIn(constrGIBDD, txtConStrPKOSP, txtPfrRepIn, txtLogFileName, iDiv, txtOspEmail, txtAdminEmail, lLogger);

                                lLogger.UpdateLLogCount(iTotalFiles);
                                // добавляю отправку письма, т.к. письма по отдельным файлам не отправлял
                                string txtMessage = "Загрузка квитанций от ПФР. Всего загружено квитанций: " + iTotalFiles.ToString() + " в ОСП " + iDiv.ToString().PadLeft(2, '0');
                                ff.SendEmail(txtMessage, txtMessage, txtAdminEmail, txtAdminEmail, txtMailServ, "");
                                
                                if(iTotalFiles > 0) lLogger.UpdateLLogStatus(2);
                                // ShowLoggerError(lLogger);
                                if (lLogger.ErrMessage.Length > 0)
                                {
                                    // если есть лог с ошибками - отправить соответствующий отчет
                                    SendEmail(lLogger.ErrMessage, "Ошибки при загрузке квитанций от ПФР в ОСП " + lLogger.OspNum.ToString(), "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");

                                    lLogger.MemoryLLog("\nСообщения об ошибках: ");
                                    lLogger.MemoryLLog("\n" + lLogger.ErrMessage);
                                    ShowLoggerError(lLogger);
                                }

                                Console.WriteLine("Итого загружено квитанций от ПФР: " + iTotalFiles.ToString());
                            }

                            #endregion
                            

                            #region Sverka
                            if (bRunSverka)
                            {

                                //INSERT INTO AGREEMENTS (ID, AGREEMENT_CODE, NAME_AGREEMENT, DESCRIPTION) VALUES (250, '250', 'МВД Карелии (реестры)', NULL);
                                txtAgreementCode = "0"; // нет соглашения т.к. тут полная сверка

                                nPack_type = 100;
                                // INSERT INTO PACK_TYPE (ID, TYPE) VALUES (100, 'Сверка по GIBDD_PLATEZH');

                                Logger_ufssprk_tools lLogger2 = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, 0, nOspNum, "Сверка по базе GIBDD_PLATEZH.");
                                // !!! временно закомментировал 20170303
                                // mvv.Sverka2(txtConStrPKOSP, constrGIBDD, mvd_id, txtLogFileName, lLogger2);
                                // !!! не забудь включить обычную сверку после новой ГИБДД
                                // - чтобы там тоже было что-то
                                // !!! но внутри отключить то что касается МВД-ГИБДД iCintrSource = 1
                                mvv.Sverka(txtConStrPKOSP, constrGIBDD, mvd_id, txtLogFileName, lLogger2);

                                //Sverka(txtConStrPKOSP, constrGIBDD, mvd_id, txtLogFileName);
                            }
                            #endregion
                        
                        
                        # region "IC_MVD_10_FMS_SELFREQUEST"
                            // автоматический ответ на запрос через базу данных

                            if (bRunIcmvdFms10_selfrequest)
                            {
                                Int64 cnt = 0;

                                DataTable DT_doc_rtn10 = null;

                                string txtAgreementSber = "ИЦ_МВД_10_ФМС";
                                txtAgreementCode = "ИЦ_МВД_10_ФМС";
                                string txtAgentCode = "ИЦ_МВД_10";
                                string txtAgentDeptCode = "ИЦ_МВД_10";

                                Logger_ufssprk_tools lLoggerError = null;
                                Logger_ufssprk_tools lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, 1, txtAgreementSber, 0, nOspNum, "Пакет запросов в ИЦ_МВД_10_ФМС.");
                                DT_doc_rtn10 = mvv.ReadIcMvdFms10Zapros(conPK_OSP, lLogger);
                                // DT_doc_fiz_sber = mvv.ReadSberZaprosNewFormatTest(conPK_OSP, lLogger);
                                ShowLoggerError(lLogger);

                                // записать в таблицу что получили строльк-то запросов и все, переходим к обработке - то есть загрузке ответов
                                if (DT_doc_rtn10 != null)
                                {
                                    lLogger.WriteLLog("Выгружено запросов в ИЦ_МВД_10_ФМС: " + DT_doc_rtn10.Rows.Count.ToString() + "\n");
                                    lLogger.UpdateLLogCount(DT_doc_rtn10.Rows.Count);
                                    lLogger.UpdateLLogStatus(2);
                                }

                                // открыть лог ответа
                                nPack_type = 2; // 2 - простой ответ
                                decimal nParentID = lLogger.logID;
                                lLogger = new Logger_ufssprk_tools(constrGIBDD, 1, nPack_type, txtAgreementCode, nParentID, nOspNum, "Пакет загрузки ответов на запрос в ИЦ_МВД_10_ФМС");

                                if (DT_doc_rtn10 != null)
                                {
                                    // написать когда и что делаем
                                    lLogger.WriteLLog(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "Начало обработки запросов по базе в ИЦ_МВД_10_ФМС. Строк запросов всего: " + DT_doc_rtn10.Rows.Count.ToString() + "\r\n");

                                    OleDbConnection VUcon = new OleDbConnection(txtGibdd10ConString);

                                    // в цикле по всем строкам DT_doc_gibdd10 
                                    string txtId = "";
                                    decimal nId = 0;
                                    decimal nRespId = 0;
                                    decimal nStatus = 19; // ответ получен
                                    string txtAnswerType = "01";
                                    string txtEntityName = "МВД";


                                    foreach (DataRow row in DT_doc_rtn10.Rows)
                                    {
                                        // получить ID запроса
                                        // учитывать что если класс контрагента не физ. лицо - то смотреть на ИНН
                                        // классы физ. лиц 2, 71,95,96,97,666
                                        // инд. предприниматель  95
                                        // итого (2,71,95,96,97,666)
                                        txtId = Convert.ToString(row["ZAPROS"]);
                                        Decimal.TryParse(txtId, out nId);
                                        if (nId > 0)
                                        {
                                            
                                            string txtFioD = Convert.ToString(row["FIOVK"]).Trim().ToUpper();
                                            // убрать двойные пробелы
                                            txtFioD = RemoveDoubleSpaces(txtFioD, 200);

                                            string txtSnils = Convert.ToString(row["DEBTOR_SNILS"]).Trim().ToUpper();
                                            string txtInn = Convert.ToString(row["DEBTOR_INN"]).Trim().ToUpper();
                                            string txtAdr = Convert.ToString(row["DEBTOR_ADDRESS"]).Trim().ToUpper();


                                            DateTime dtBornD = DateTime.MinValue;
                                            string txtBornD = "";
                                            string txtDocNum = "";
                                            DateTime dtStart = DateTime.MinValue;
                                            DateTime dtDateVh = DateTime.MinValue;

                                            List<int> fiz = new List<int>(new[] { 2, 71, 95, 96, 97, 666 });
                                            int iDbtrCls = Convert.ToInt32(row["ID_DBTRCLS"]);
                                            // дату рождения проверяем только для физ лиц
                                            // для юр лиц нужно будет проверять ИНН и искать по нему

                                            if (fiz.Contains(iDbtrCls) && !row["DATROZHD"].Equals(DBNull.Value))
                                            {
                                                txtBornD = Convert.ToString(row["DATROZHD"]);
                                                if (DateTime.TryParse(txtBornD, out dtBornD))
                                                {
                                                    // TODO: получить ответ как DataTable и тут уже разбирать параметры и собирать строчку - иначе не вставить структурированные сведения
                                                    // string txtResp = mvv.FindVU(VUcon, txtGibdd10DataBase, txtFioD, dtBornD, lLogger);

                                                    // когда будем брать паспрта из ufssprk-tools - то будет VUcon
                                                    // DataTable tbl = mvv.FindDul(VUcon, txtFioD, dtBornD, lLogger);
                                                    
                                                    // а пока смотрим в базе того же ОСП - conPK_OSP
                                                    DataTable tbl = mvv.FindDul(conPK_OSP, txtFioD, dtBornD, txtSnils, txtInn, txtAdr, lLogger);
                                                    
                                                    string txtResp = "";
                                                    if (tbl != null && tbl.Rows.Count > 0)
                                                    {
                                                        txtResp += "Получены сведения об удостоверении личности должника " + txtFioD + " " + dtBornD.ToShortDateString() + " г.р.\r\n";

                                                        foreach (DataRow rrow in tbl.Rows)
                                                        {
                                                            string txtID_DBTR_ID_SERIAL = Convert.ToString(rrow["ID_DBTR_ID_SERIAL"]).Trim();
                                                            txtResp += "Паспорт РФ серия " + txtID_DBTR_ID_SERIAL;

                                                            string txtID_DBTR_ID_NUMBER = Convert.ToString(rrow["ID_DBTR_ID_NUMBER"]).Trim();
                                                            txtResp += " номер " + txtID_DBTR_ID_NUMBER;

                                                            string txtID_DBTR_ID_DATE = Convert.ToString(rrow["ID_DBTR_ID_DATE"]).Trim();
                                                            txtResp += " выдан " + txtID_DBTR_ID_DATE;

                                                            string txtID_DBTR_ID_OFFICE = Convert.ToString(rrow["ID_DBTR_ID_OFFICE"]).Trim();
                                                            txtResp += " " + txtID_DBTR_ID_OFFICE;

                                                            string txtID_DBTR_ID_CODE_DEP = Convert.ToString(rrow["ID_DBTR_ID_CODE_DEP"]).Trim();
                                                            txtResp += " (код подразделения: " + txtID_DBTR_ID_CODE_DEP + ").\r\n";

                                                            if (tbl.Rows.Count > 1) txtResp += "\r\n"; // перевод строки чтобы разделить записи если их больше 1
                                                        }
                                                    }
                                                    else txtResp = "Нет данных";

                                                    txtAnswerType = "01";
                                                    if (txtResp == "Нет данных") txtAnswerType = "02";
                                                    else if (txtResp.Length == 0) txtAnswerType = "03"; //  требует уточнения
                                                    // !!! TODO
                                                    // вставить ответ в ИТ
                                                    // nRespId = mvv.InsertResponseIntTable(conPK_OSP, nId, txtResp, DateTime.Today, nStatus, lLogger.logID, txtAgreementCode, txtAgreementCode, txtAgreementCode, txtEntityName, txtAnswerType, lLogger);
                                                    decimal nExtKey = mvv.InsertResponseIntTableNewExtKey(conPK_OSP, nId, txtResp, DateTime.Today, nStatus, lLogger.logID, txtAgentCode, txtAgentDeptCode, txtAgreementCode, txtEntityName, txtAnswerType, lLogger);
                                                    if (nExtKey > 0)
                                                    {

                                                        // если ответ был с данными - вставить в ИТ EXT_IDENTIFICATION_DATA структурированные сведения
                                                        if (txtAnswerType == "01")
                                                        {
                                                            // пройти по всем строчкам и вставить в таблицу только те, которые актуальны (не сняты с учета)
                                                            if (tbl != null && tbl.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow rrow in tbl.Rows)
                                                                {
                                                                    // прочитать данные из таблички и добавить в ИТ
                                                                    // ID_DBTR_ID_SERIAL, ID_DBTR_ID_NUMBER, ID_DBTR_ID_DATE, ID_DBTR_ID_OFFICE, ID_DBTR_ID_CODE_DEP
                                                                    // txtDocuemntKey = nExtKey.ToString()
                                                                    decimal nExtIdentID = mvv.InsertExtIdentDataIntTable(conPK_OSP, nExtKey.ToString(), txtFioD, rrow, lLogger);

                                                                }
                                                            }
                                                        }
                                                        // теперь нужно обозначить запрос в ИТ как выгруженный
                                                        if (mvv.UpdateExtRequestRow(conPK_OSP, row, lLogger))
                                                        {
                                                            lLogger.WriteLLog("Запрос № " + txtId.ToString() + " успешно обработан. Получен ответ № " + nExtKey.ToString() + "\r\n");
                                                            cnt++;
                                                        }
                                                        else
                                                            lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Несмотря на это в базу загружен ответ № " + nExtKey.ToString() + "\r\n");
                                                    }
                                                    else
                                                    {
                                                        lLogger.WriteLLog("Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n");
                                                        lLogger.ErrMessage += "Ошибка! Запрос № " + txtId.ToString() + " не удалось обработать - запрос будет обработан повторно. Ответ не получен.\r\n";
                                                    }
                                                }
                                            }

                                        } // end if (nId > 0)
                                    } // end  foreach (DataRow row in DT_doc_rtn10.Rows)
                                } //end if (DT_doc_rtn10 != null)
                            } // end if (bRunIcmvdFms10_selfrequest)
# endregion                                            

                        }
                        else
                        {
                            SendEmail("Недоступна база ufssprk-tools по пути: " + constrGIBDD, "Ошибка. Недоступна база ufssprk-tools.", "mvv_report@r10.fssprus.ru", "mvv_report@r10.fssprus.ru", txtMailServ, "");
                            Console.WriteLine("Недоступна база ufssprk-tools по пути: " + constrGIBDD);
                            WriteTofile(DateTime.Now.ToString() + "Недоступна база ufssprk-tools по пути: " + constrGIBDD, txtLogFileName);


                        }
                        // КОНЕЦ ПРОВЕРКИ ЧТО ПОЛУЧЕН КОД осп БЕЗ ОШИБОК
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception Thrown: " + e.ToString());
                    WriteTofile(DateTime.Now.ToString() + " Ошибка приложения: " + e.ToString(), txtLogFileName);
                }
            }
            
        }

        

        static bool ShowLoggerError(Logger_ufssprk_tools lLogger)
        {
            if (lLogger.ErrMessage.Length > 0)
            {
                Console.WriteLine("\nОшибка приложения. Message: " + lLogger.ErrMessage);
                lLogger.WriteLLog("\nОшибка приложения. Message: " + lLogger.ErrMessage);
                lLogger.ErrMessage = "";
                return true;
            }
            return false;
        }


        static Decimal GetOSP_Num(OleDbConnection con)
        {
            Decimal res = 0;
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                //OleDbCommand cmd = new OleDbCommand("Select DEPARTMENT from OSP", con, tran);
                OleDbCommand cmd = new OleDbCommand("select first 1 osp.department from system_site join osp on osp.osp_system_site_id = system_site.system_site_id", con, tran);
                res = Convert.ToDecimal(cmd.ExecuteScalar());
                tran.Rollback();
                con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
            }
            return res;

        }


        static bool DeleteUsedGibddPlat(OleDbConnection conG, bool flDeleteUsed, bool flDeleteOld)
        {

            OleDbCommand cmdD, cmdOldD;
            OleDbTransaction tran = null;

            try
            {
                if ((conG == null) || (conG.State.Equals(ConnectionState.Closed)))
                {
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
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
                if (conG != null)
                {
                    conG.Close();
                }
            }
            return false;
        }


        static bool AlterIndxI_ID(OleDbConnection con, bool flActive)
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
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }

        static bool UpdateGibddPlatezh(OleDbConnection conG, string txtNumber, int iValue, int iContrSource)
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
                cmdU.CommandText = "update GIBDD_PLATEZH SET FL_USE = :FL_USE WHERE NUMBER = :NUMBER and SOURCE_ID = :SOURCE_ID";

                cmdU.Parameters.Add(new OleDbParameter(":FL_USE", iValue));
                cmdU.Parameters.Add(new OleDbParameter(":NUMBER", txtNumber));
                cmdU.Parameters.Add(new OleDbParameter(":NUMBER", iContrSource));


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
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
                if (conG != null)
                {
                    conG.Close();
                }
            }
            return false;
        }


        static string Money_ToStr(decimal nMoney)
        {
            string txtResult = "";
            txtResult = nMoney.ToString("N2").Replace(".", " руб. ");
            txtResult = txtResult.Replace(",", " руб. ") + " коп.";

            return txtResult;
        }

        static string Money_ToStr(double nMoney)
        {
            string txtResult = "";
            txtResult = nMoney.ToString("N2").Replace(".", " руб. ");
            txtResult = txtResult.Replace(",", " руб. ") + " коп.";

            return txtResult;
        }

        static decimal ID_InsertOtherIP_DocTo_PK_OSP(OleDbConnection con, decimal nStatus, decimal nUserID, DateTime dtIdate, decimal nIP_ID, DateTime dtExtDocDate, string txtExtDocNum, string txtContent, decimal nContrID)
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
                txtContrName = GetLegal_Name(con, nContrID);
                txtContrAdr = GetLegal_Adr(con, nContrID);

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
                cmdIP.CommandText = "select d.docstatusid, d.doc_number, d_ip.ip_exec_prist, d_ip.ip_exec_prist_name, d_ip_d.id_docdate, d_ip.id_debttext, d_ip.ipno_num  from doc_ip_doc d_ip_d left join doc_ip d_ip on d_ip_d.id = d_ip.id left join document d on d_ip_d.id = d.id where d_ip_d.id = :IP_ID";
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

                // cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCUMENTCLASSID", Convert.ToDecimal(11))); // класс документооборота для объекта I - Ыходящий документ
                cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCUMENTCLASSID", Convert.ToDecimal(272))); // класс документооборота для объекта I - Ыходящий документ

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
                if (b20reliz)
                {
                    cmdInsI.CommandText = "insert into I (ID, PAGECNT, APPNCNT, SECURTYPE, APPBNPAGECNT, I_IDATE, EXTDOCDATE, EXTDOCNO, CONTR, CONTR_NAME, ADR, EXECUTOR, EXECUTOR_NAME, OFF_SPECIAL_CONTROL, CONTR_IS_INITIATOR)";
                    cmdInsI.CommandText += "  VALUES (:ID, 1, 0, 2, 0, :I_IDATE, :EXTDOCDATE, :EXTDOCNO, :CONTR, :CONTR_NAME, :ADR, :EXECUTOR, :EXECUTOR_NAME, :OFF_SPECIAL_CONTROL, :CONTR_IS_INITIATOR)";
                }
                else
                {
                    cmdInsI.CommandText = "insert into I (ID, PAGECNT, APPNCNT, SECURTYPE, APPBNPAGECNT, I_IDATE, EXTDOCDATE, EXTDOCNO, CONTR, CONTR_NAME, ADR, EXECUTOR, EXECUTOR_NAME, OFF_SPECIAL_CONTROL)";
                    cmdInsI.CommandText += "  VALUES (:ID, 1, 0, 2, 0, :I_IDATE, :EXTDOCDATE, :EXTDOCNO, :CONTR, :CONTR_NAME, :ADR, :EXECUTOR, :EXECUTOR_NAME, :OFF_SPECIAL_CONTROL)";
                }

                cmdInsI.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                //cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", dtIdate));
                cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", DateTime.Today));
                cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCDATE", dtExtDocDate));
                cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCNO", txtExtDocNum));
                cmdInsI.Parameters.Add(new OleDbParameter(":CONTR", nContrID));
                cmdInsI.Parameters.Add(new OleDbParameter(":CONTR_NAME", txtContrName));
                cmdInsI.Parameters.Add(new OleDbParameter(":ADR", txtContrAdr));
                cmdInsI.Parameters.Add(new OleDbParameter(":EXECUTOR", Convert.ToDecimal(nSUSER_ID)));
                cmdInsI.Parameters.Add(new OleDbParameter(":EXECUTOR_NAME", Convert.ToString(txtSUSER)));
                // новое поле - ставим 1 чтобы не было на контроле
                cmdInsI.Parameters.Add(new OleDbParameter(":OFF_SPECIAL_CONTROL", 1));
                if (b20reliz)
                {
                    cmdInsI.Parameters.Add(new OleDbParameter(":CONTR_IS_INITIATOR", 1));
                }



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
                    // MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
                if (con != null)
                {
                    con.Close();
                }
            }
            return -1;
        }
        

        static decimal ID_InsertPlatDocTo_PK_OSP(OleDbConnection con, decimal nStatus, decimal nUserID, double nAmount, DateTime dtIdate, decimal nIP_ID, DateTime dtExtDocDate, string txtExtDocNum, decimal nContrID, string txtFIO_D)
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
                cmdIP.CommandText = "select d_ip_d.id_dbtr_name, d_ip_d.id_dbtr_adr, d.docstatusid, d.doc_number, d_ip.id_dbtr, d_ip.ip_exec_prist, d_ip.ip_exec_prist_name, d_ip_d.id_docdate, d_ip.id_debttext, d_ip.ipno_num  from doc_ip_doc d_ip_d left join doc_ip d_ip on d_ip_d.id = d_ip.id left join document d on d_ip_d.id = d.id where d_ip_d.id = :IP_ID";
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
                if (b20reliz)
                {
                    cmdInsI.CommandText = "insert into I (ID, PAGECNT, APPNCNT, SECURTYPE, APPBNPAGECNT, I_IDATE, EXTDOCDATE, EXTDOCNO, CONTR, CONTR_NAME, ADR, OFF_SPECIAL_CONTROL, CONTR_IS_INITIATOR)";
                    cmdInsI.CommandText += "  VALUES (:ID, 1, 0, 2, 0, :I_IDATE, :EXTDOCDATE, :EXTDOCNO, :CONTR, :CONTR_NAME, :ADR, :OFF_SPECIAL_CONTROL, :CONTR_IS_INITIATOR)";
                }
                else
                {
                    cmdInsI.CommandText = "insert into I (ID, PAGECNT, APPNCNT, SECURTYPE, APPBNPAGECNT, I_IDATE, EXTDOCDATE, EXTDOCNO, CONTR, CONTR_NAME, ADR, OFF_SPECIAL_CONTROL)"; //, CONTR_IS_INITIATOR)";
                    cmdInsI.CommandText += "  VALUES (:ID, 1, 0, 2, 0, :I_IDATE, :EXTDOCDATE, :EXTDOCNO, :CONTR, :CONTR_NAME, :ADR, :OFF_SPECIAL_CONTROL)"; //, :CONTR_IS_INITIATOR)";
                }
                cmdInsI.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
                //cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", dtIdate));
                cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", DateTime.Today));
                cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCDATE", dtExtDocDate));
                cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCNO", txtExtDocNum));
                cmdInsI.Parameters.Add(new OleDbParameter(":CONTR", nContrID));
                cmdInsI.Parameters.Add(new OleDbParameter(":CONTR_NAME", txtContrName));
                cmdInsI.Parameters.Add(new OleDbParameter(":ADR", txtContrAdr));
                // новое поле - ставим 1 чтобы не было на контроле
                cmdInsI.Parameters.Add(new OleDbParameter(":OFF_SPECIAL_CONTROL", 1));
                if (b20reliz)
                {
                    cmdInsI.Parameters.Add(new OleDbParameter(":CONTR_IS_INITIATOR", 1));
                }



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

                cmdInsI_DEPOSIT = new OleDbCommand();
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
                // тут есть параметр - причина.
                // штраф ГИБДД не подходит для всего под
                // 101
                // Погашение взыскателю (РИЦ ЖХ)

                // 102
                // Погашение взыскателю (КРЦ)


                cmdInsI_OP_CS = new OleDbCommand();
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

                cmdInsI_OP_CS_ENDDBT = new OleDbCommand();
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
                    //MessageBox.Show("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, "Внимание!", MessageBoxButtons.OK);
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
                if (con != null)
                {
                    con.Close();
                }
            }
            return -1;
        }




        // добавить депозитный документ о погашении долга 
        //static decimal ID_InsertPlatDocTo_PK_OSP(OleDbConnection con, decimal nStatus, decimal nUserID, double nAmount, DateTime dtIdate, decimal nIP_ID, DateTime dtExtDocDate, string txtExtDocNum, decimal nContrID, string txtFIO_D)
        //{

        //    OleDbCommand cmdIP, cmd, cmdInsDoc, cmdInsI, cmdInsI_IP, cmdInsI_DEPOSIT, cmdInsI_OP_CS, cmdInsI_OP_CS_ENDDBT;
        //    Decimal newID, prevID;
        //    OleDbTransaction tran = null;
        //    DataSet dsIP_params;
        //    DataTable dtIP_params;
        //    decimal nIPNO_NUM = 0;
        //    string txtIP_DocNumber = "";
        //    decimal nSUSER_ID;
        //    string txtSUSER = "";
        //    string txtID_DEBTCLS_NAME = "";
        //    string txtContrName = "";
        //    string txtContrAdr = "";
        //    decimal id_dbtr = 0;

        //    try
        //    {
        //        // самое лучшее решение - это выбрать contr_id из i_id и передать сюда

        //        dsIP_params = new DataSet();
        //        dtIP_params = dsIP_params.Tables.Add("IP_params");
        //        newID = 0;
        //        prevID = 0;
        //        id_dbtr = 0;
        //        con.Open();
        //        tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

        //        // узнать параметры ИП по ИД

        //        cmdIP = new OleDbCommand();
        //        cmdIP.Connection = con;
        //        cmdIP.Transaction = tran;
        //        cmdIP.CommandText = "select d_ip_d.id_dbtr_name, d_ip_d.id_dbtr_adr, d.docstatusid, d.doc_number, d_ip.id_dbtr, d_ip.ip_exec_prist, d_ip.ip_exec_prist_name, d_ip_d.id_docdate, d_ip_d.id_debttext, d_ip.ipno_num  from doc_ip_doc d_ip_d left join doc_ip d_ip on d_ip_d.id = d_ip.id left join document d on d_ip_d.id = d.id where d_ip_d.id = :IP_ID";
        //        cmdIP.Parameters.Add(new OleDbParameter(":IP_ID", Convert.ToDecimal(nIP_ID)));
        //        using (OleDbDataReader rdr = cmdIP.ExecuteReader(CommandBehavior.Default))
        //        {
        //            dsIP_params.Load(rdr, LoadOption.OverwriteChanges, dtIP_params);
        //            rdr.Close();
        //        }

        //        if ((dsIP_params != null) && (dsIP_params.Tables.Count > 0))
        //        {
        //            txtIP_DocNumber = Convert.ToString(dsIP_params.Tables[0].Rows[0]["doc_number"]);
        //            nIPNO_NUM = Convert.ToDecimal(dsIP_params.Tables[0].Rows[0]["ipno_num"]);
        //            nSUSER_ID = Convert.ToDecimal(dsIP_params.Tables[0].Rows[0]["ip_exec_prist"]);
        //            txtSUSER = Convert.ToString(dsIP_params.Tables[0].Rows[0]["ip_exec_prist_name"]);
        //            txtID_DEBTCLS_NAME = Convert.ToString(dsIP_params.Tables[0].Rows[0]["id_debttext"]);
        //            txtContrName = Convert.ToString(dsIP_params.Tables[0].Rows[0]["id_dbtr_name"]); // теперь отправителем будет сам должник, а почему?
        //            txtContrAdr = Convert.ToString(dsIP_params.Tables[0].Rows[0]["id_dbtr_adr"]);

        //            if (nSUSER_ID > 0)
        //            {
        //                nUserID = nSUSER_ID;
        //            }
        //        }
        //        else
        //        {
        //            return -1;
        //        }

        //        // получить новый ключ
        //        cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
        //        newID = Convert.ToDecimal(cmd.ExecuteScalar());

        //        // вставить DOCUMENT
        //        cmdInsDoc = new OleDbCommand();
        //        cmdInsDoc.Connection = con;
        //        cmdInsDoc.Transaction = tran;
        //        cmdInsDoc.CommandText = "insert into DOCUMENT (ID, METAOBJECTNAME, DOCSTATUSID, DOCUMENTCLASSID, CREATE_DATE, SUSER_ID, AMOUNT)";
        //        cmdInsDoc.CommandText += " VALUES (:ID, 'I_OP_CS_ENDDBT', :DOCSTATUSID, :DOCUMENTCLASSID, :CREATE_DATE, :SUSER_ID, :AMOUNT)";

        //        cmdInsDoc.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));

        //        //cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(1)));
        //        cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(nStatus)));

        //        cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCUMENTCLASSID", Convert.ToDecimal(78)));
        //        //cmdInsDoc.Parameters.Add(new OleDbParameter(":PARENT_ID", Convert.ToDecimal(nID)));
        //        cmdInsDoc.Parameters.Add(new OleDbParameter(":CREATE_DATE", DateTime.Now));
        //        cmdInsDoc.Parameters.Add(new OleDbParameter(":SUSER_ID", Convert.ToDecimal(nUserID)));
        //        cmdInsDoc.Parameters.Add(new OleDbParameter(":AMOUNT", Convert.ToDouble(nAmount)));


        //        if (cmdInsDoc.ExecuteNonQuery() == -1)
        //        {
        //            Exception ex = new Exception("Error inserting new row to document table parent_id = " + newID.ToString());
        //            throw ex;
        //        }

        //        // вставить I
        //        // - Отправитель 	I.CONTR_NAME
        //        // - Адрес отправителя I.ADR


        //        cmdInsI = new OleDbCommand();
        //        cmdInsI.Connection = con;
        //        cmdInsI.Transaction = tran;
        //        cmdInsI.CommandText = "insert into I (ID, PAGECNT, APPNCNT, SECURTYPE, APPBNPAGECNT, I_IDATE, EXTDOCDATE, EXTDOCNO, CONTR, CONTR_NAME, ADR)";
        //        cmdInsI.CommandText += "  VALUES (:ID, 1, 0, 2, 0, :I_IDATE, :EXTDOCDATE, :EXTDOCNO, :CONTR, :CONTR_NAME, :ADR)";
        //        cmdInsI.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
        //        //cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", dtIdate));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", DateTime.Today));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCDATE", dtExtDocDate));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCNO", txtExtDocNum));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":CONTR", nContrID));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":CONTR_NAME", txtContrName));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":ADR", txtContrAdr));

        //        if (cmdInsI.ExecuteNonQuery() == -1)
        //        {
        //            Exception ex = new Exception("Error inserting new row to I table id = " + newID.ToString());
        //            throw ex;

        //        }


        //        // вставить I_IP


        //        cmdInsI_IP = new OleDbCommand();
        //        cmdInsI_IP.Connection = con;
        //        cmdInsI_IP.Transaction = tran;
        //        cmdInsI_IP.CommandText = "insert into I_IP (ID, IP_DOC_NUMBER,IP_ID, IPNO_NUM, ID_DEBTCLS_NAME, IP_EXEC_PRIST, IP_EXEC_PRIST_NAME)";
        //        cmdInsI_IP.CommandText += "  VALUES (:ID, :IP_DOC_NUMBER, :IP_ID, :IPNO_NUM, :ID_DEBTCLS_NAME, :IP_EXEC_PRIST, :IP_EXEC_PRIST_NAME)";

        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_DOC_NUMBER", Convert.ToString(txtIP_DocNumber)));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_ID", Convert.ToDecimal(nIP_ID)));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":IPNO_NUM", Convert.ToDecimal(nIPNO_NUM)));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":ID_DEBTCLS_NAME", txtID_DEBTCLS_NAME));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_EXEC_PRIST", Convert.ToDecimal(nSUSER_ID)));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_EXEC_PRIST_NAME", Convert.ToString(txtSUSER)));

        //        if (cmdInsI_IP.ExecuteNonQuery() == -1)
        //        {
        //            Exception ex = new Exception("Error inserting new row to I_IP table id = " + newID.ToString());
        //            throw ex;
        //        }


        //        // вставить I_DEPOSIT 

        //        cmdInsI_DEPOSIT = new OleDbCommand();
        //        cmdInsI_DEPOSIT.Connection = con;
        //        cmdInsI_DEPOSIT.Transaction = tran;
        //        cmdInsI_DEPOSIT.CommandText = "insert into I_DEPOSIT (ID)";
        //        cmdInsI_DEPOSIT.CommandText += "  VALUES (:ID)";
        //        cmdInsI_DEPOSIT.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));

        //        if (cmdInsI_DEPOSIT.ExecuteNonQuery() == -1)
        //        {
        //            Exception ex = new Exception("Error inserting new row to I_DEPOSIT  table id = " + newID.ToString());
        //            throw ex;

        //        }

        //        // вставить I_OP_CS

        //        cmdInsI_OP_CS = new OleDbCommand();
        //        cmdInsI_OP_CS.Connection = con;
        //        cmdInsI_OP_CS.Transaction = tran;
        //        cmdInsI_OP_CS.CommandText = "insert into I_OP_CS (ID, CHANGEDBT_REASON_ID, CHANGEDBT_REASON_DESCR, I_OP_CS_CHANGESUM)";
        //        cmdInsI_OP_CS.CommandText += "  VALUES (:ID, 3, 'Оплата штрафа в ГИБДД', :I_OP_CS_CHANGESUM)";
        //        cmdInsI_OP_CS.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
        //        cmdInsI_OP_CS.Parameters.Add(new OleDbParameter(":I_OP_CS_CHANGESUM", Convert.ToDouble(nAmount)));

        //        if (cmdInsI_OP_CS.ExecuteNonQuery() == -1)
        //        {
        //            Exception ex = new Exception("Error inserting new row to I_OP_CS  table id = " + newID.ToString());
        //            throw ex;

        //        }


        //        // вставить I_OP_CS_ENDDBT

        //        cmdInsI_OP_CS_ENDDBT = new OleDbCommand();
        //        cmdInsI_OP_CS_ENDDBT.Connection = con;
        //        cmdInsI_OP_CS_ENDDBT.Transaction = tran;
        //        cmdInsI_OP_CS_ENDDBT.CommandText = "insert into I_OP_CS_ENDDBT (ID)";
        //        cmdInsI_OP_CS_ENDDBT.CommandText += "  VALUES (:ID)";
        //        cmdInsI_OP_CS_ENDDBT.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));

        //        if (cmdInsI_OP_CS_ENDDBT.ExecuteNonQuery() == -1)
        //        {
        //            Exception ex = new Exception("Error inserting new row to I_OP_CS_ENDDBT  table id = " + newID.ToString());
        //            throw ex;

        //        }

        //        tran.Commit();
        //        con.Close();

        //        return newID;

        //    }
        //    catch (OleDbException ole_ex)
        //    {
        //        if (tran != null)
        //        {
        //            tran.Rollback();
        //        }
        //        if (con != null)
        //        {
        //            con.Close();
        //        }
        //        foreach (OleDbError err in ole_ex.Errors)
        //        {
        //            Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
        //        if (con != null)
        //        {
        //            con.Close();
        //        }
        //    }
        //    return -1;
        //}

        static decimal FindIDNumRicZH(OleDbConnection con, string txtNomID, double nSumID, DateTime dtDatID, string txtNameDolg)
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
                txtSql = "select i_id.ip_id from i_id join document d on i_id.ip_id = d.id where i_id.crdrcontr_name = 'РИЦ ЖХ ООО' and (d.docstatusid != -1) and (d.docstatusid != 7) and (d.docstatusid != 10) and i_id.debtsum = " + nSumID.ToString().Replace(',', '.') + " and i_id.id_docdate = '" + dtDatID.ToShortDateString() + "' and i_id.id_docno = '" + Convert.ToString(txtNomID) + "' and UPPER(i_id.dbtrcontr_name) = '" + txtNameDolg + "'";
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
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
            }


            if (con != null)
            {
                con.Close();
            }

            return res;
        }


        static decimal GetContrID(OleDbConnection con, int iSourceID)
        {
            OleDbTransaction tran;
            string txtSql = "";
            decimal res = -1;
            try
            {

                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                txtSql = "select entt_id from CONTR WHERE ID = " + Convert.ToString(iSourceID);

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
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
            }


            if (con != null)
            {
                con.Close();
            }

            return res;
        }



        static decimal FindIDNum(OleDbConnection con, string txtNomID, double nSumID, DateTime dtDatID)
        {
            OleDbTransaction tran;
            string txtSql = "";
            decimal res = -1;
            try
            {

                con.Open();
                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                //txtSql = "select d_ip_d.id from doc_ip_doc d_ip_d left join doc_ip d_ip on d_ip_d.id = d_ip.id left join document d on d_ip_d.id = d.id where ((d_ip_d.id_debtcls = 22) or (d_ip_d.id_debtcls = 37) or (d_ip_d.id_debtcls = 38) or (d_ip_d.id_debtcls = 45)) and (d.docstatusid != -1) and (d.docstatusid != 7) and (d.docstatusid != 10) and d_ip.id_debtsum = " + nSumID.ToString() + " and d_ip_d.id_docdate = '" + dtDatID.ToShortDateString() + "' and d_ip_d.id_docno = '" + Convert.ToString(txtNomID) + "'";
                // txtSql = "select i_id.ip_id from i_id left join document d on i_id.ip_id = d.id where ((i_id.debtcls = 22) or (i_id.debtcls = 37) or (i_id.debtcls = 38) or (i_id.debtcls = 45)) and (d.docstatusid != -1) and (d.docstatusid != 7) and (d.docstatusid != 10) and i_id.debtsum = " + nSumID.ToString().Replace(',', '.') + " and i_id.id_docdate = '" + dtDatID.ToShortDateString() + "' and i_id.id_docno = '" + Convert.ToString(txtNomID) + "'";
                txtSql = "select first 1 i_id.ip_id from i_id join document d on i_id.ip_id = d.id where      ((i_id.debtcls = 22) or (i_id.debtcls = 37) or (i_id.debtcls = 38) or (i_id.debtcls = 45)) and (d.docstatusid != -1) and i_id.debtsum = " + nSumID.ToString().Replace(',', '.') + " and i_id.id_docdate = '" + dtDatID.ToShortDateString() + "' and i_id.id_docno = '" + Convert.ToString(txtNomID) + "'";
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
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
            }


            if (con != null)
            {
                con.Close();
            }

            return res;
        }

        static string GetLegal_Name(OleDbConnection con, Decimal code)
        {
            String res = "нет значения в базе данных";
            try
            {
                if (con != null && con.State.Equals(ConnectionState.Closed)) con.Open();
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
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
            }
            return res;
        }


        static string GetLegal_Adr(OleDbConnection con, Decimal code)
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
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
            }
            return res;
        }

        // добавить прочий документ по ИП
        //static decimal ID_InsertOtherIP_DocTo_PK_OSP(OleDbConnection con, decimal nStatus, decimal nUserID, DateTime dtIdate, decimal nIP_ID, DateTime dtExtDocDate, string txtExtDocNum, string txtContent, decimal nContrID)
        //{

        //    OleDbCommand cmdIP, cmd, cmdInsDoc, cmdInsI, cmdInsI_IP, cmdInsI_IP_OTHER;
        //    Decimal newID, prevID;
        //    OleDbTransaction tran = null;
        //    DataSet dsIP_params;
        //    DataTable dtIP_params;
        //    decimal nIPNO_NUM = 0;
        //    string txtIP_DocNumber = "";
        //    decimal nSUSER_ID;
        //    string txtSUSER = "";
        //    string txtID_DEBTCLS_NAME = "";
        //    string txtContrName = "";
        //    string txtContrAdr = "";

        //    try
        //    {
        //        txtContrName = GetLegal_Name(con, nContrID);
        //        txtContrAdr = GetLegal_Adr(con, nContrID);

        //        dsIP_params = new DataSet();
        //        dtIP_params = dsIP_params.Tables.Add("IP_params");
        //        newID = 0;
        //        prevID = 0;
        //        con.Open();
        //        tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

        //        // узнать параметры ИП по ИД

        //        cmdIP = new OleDbCommand();
        //        cmdIP.Connection = con;
        //        cmdIP.Transaction = tran;
        //        cmdIP.CommandText = "select d.docstatusid, d.doc_number, d_ip.ip_exec_prist, d_ip.ip_exec_prist_name, d_ip_d.id_docdate, d_ip_d.id_debttext, d_ip.ipno_num  from doc_ip_doc d_ip_d left join doc_ip d_ip on d_ip_d.id = d_ip.id left join document d on d_ip_d.id = d.id where d_ip_d.id = :IP_ID";
        //        cmdIP.Parameters.Add(new OleDbParameter(":IP_ID", Convert.ToDecimal(nIP_ID)));
        //        using (OleDbDataReader rdr = cmdIP.ExecuteReader(CommandBehavior.Default))
        //        {
        //            dsIP_params.Load(rdr, LoadOption.OverwriteChanges, dtIP_params);
        //            rdr.Close();
        //        }

        //        if ((dsIP_params != null) && (dsIP_params.Tables.Count > 0))
        //        {
        //            txtIP_DocNumber = Convert.ToString(dsIP_params.Tables[0].Rows[0]["doc_number"]);
        //            nIPNO_NUM = Convert.ToDecimal(dsIP_params.Tables[0].Rows[0]["ipno_num"]);
        //            nSUSER_ID = Convert.ToDecimal(dsIP_params.Tables[0].Rows[0]["ip_exec_prist"]);
        //            txtSUSER = Convert.ToString(dsIP_params.Tables[0].Rows[0]["ip_exec_prist_name"]);
        //            txtID_DEBTCLS_NAME = Convert.ToString(dsIP_params.Tables[0].Rows[0]["id_debttext"]);
        //            if (nSUSER_ID > 0)
        //            {
        //                nUserID = nSUSER_ID;
        //            }
        //        }
        //        else
        //        {
        //            return -1;
        //        }

        //        // получить новый ключ
        //        cmd = new OleDbCommand("SELECT gen_id(seq_document, 1) FROM RDB$DATABASE", con, tran);
        //        newID = Convert.ToDecimal(cmd.ExecuteScalar());

        //        // вставить DOCUMENT
        //        cmdInsDoc = new OleDbCommand();
        //        cmdInsDoc.Connection = con;
        //        cmdInsDoc.Transaction = tran;
        //        cmdInsDoc.CommandText = "insert into DOCUMENT (ID, METAOBJECTNAME, DOCSTATUSID, DOCUMENTCLASSID, CREATE_DATE, SUSER_ID)";
        //        cmdInsDoc.CommandText += " VALUES (:ID, 'I_IP_OTHER', :DOCSTATUSID, :DOCUMENTCLASSID, :CREATE_DATE, :SUSER_ID)";

        //        cmdInsDoc.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));

        //        //cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(1)));
        //        cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCSTATUSID", Convert.ToDecimal(nStatus)));

        //        cmdInsDoc.Parameters.Add(new OleDbParameter(":DOCUMENTCLASSID", Convert.ToDecimal(11))); // класс документооборота для объекта I - Ыходящий документ
        //        //cmdInsDoc.Parameters.Add(new OleDbParameter(":PARENT_ID", Convert.ToDecimal(nID)));
        //        cmdInsDoc.Parameters.Add(new OleDbParameter(":CREATE_DATE", DateTime.Now));
        //        cmdInsDoc.Parameters.Add(new OleDbParameter(":SUSER_ID", Convert.ToDecimal(nUserID)));
        //        //cmdInsDoc.Parameters.Add(new OleDbParameter(":AMOUNT", Convert.ToDouble(nAmount)));


        //        if (cmdInsDoc.ExecuteNonQuery() == -1)
        //        {
        //            Exception ex = new Exception("Error inserting new row to document table parent_id = " + newID.ToString());
        //            throw ex;
        //        }

        //        // вставить I

        //        cmdInsI = new OleDbCommand();
        //        cmdInsI.Connection = con;
        //        cmdInsI.Transaction = tran;

        //        if (b20reliz)
        //        {
        //            cmdInsI.CommandText = "insert into I (ID, PAGECNT, APPNCNT, SECURTYPE, APPBNPAGECNT, I_IDATE, EXTDOCDATE, EXTDOCNO, CONTR, CONTR_NAME, ADR, EXECUTOR, EXECUTOR_NAME, OFF_SPECIAL_CONTROL, CONTR_IS_INITIATOR)";
        //            cmdInsI.CommandText += "  VALUES (:ID, 1, 0, 2, 0, :I_IDATE, :EXTDOCDATE, :EXTDOCNO, :CONTR, :CONTR_NAME, :ADR, :EXECUTOR, :EXECUTOR_NAME, :OFF_SPECIAL_CONTROL), :CONTR_IS_INITIATOR)";
        //        }
        //        else
        //        {
        //            cmdInsI.CommandText = "insert into I (ID, PAGECNT, APPNCNT, SECURTYPE, APPBNPAGECNT, I_IDATE, EXTDOCDATE, EXTDOCNO, CONTR, CONTR_NAME, ADR, EXECUTOR, EXECUTOR_NAME, OFF_SPECIAL_CONTROL)";
        //            cmdInsI.CommandText += "  VALUES (:ID, 1, 0, 2, 0, :I_IDATE, :EXTDOCDATE, :EXTDOCNO, :CONTR, :CONTR_NAME, :ADR, :EXECUTOR, :EXECUTOR_NAME, :OFF_SPECIAL_CONTROL)";
        //        }


        //        cmdInsI.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
        //        //cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", dtIdate));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":I_IDATE", DateTime.Today));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCDATE", dtExtDocDate));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":EXTDOCNO", txtExtDocNum));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":CONTR", nContrID));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":CONTR_NAME", txtContrName));
        //        cmdInsI.Parameters.Add(new OleDbParameter(":ADR", txtContrAdr));
                
        //        if (b20reliz)
        //        {
        //            cmdInsI.Parameters.Add(new OleDbParameter(":CONTR_IS_INITIATOR", 1));
        //        }


        //        if (cmdInsI.ExecuteNonQuery() == -1)
        //        {
        //            Exception ex = new Exception("Error inserting new row to I table id = " + newID.ToString());
        //            throw ex;

        //        }


        //        // вставить I_IP


        //        cmdInsI_IP = new OleDbCommand();
        //        cmdInsI_IP.Connection = con;
        //        cmdInsI_IP.Transaction = tran;
        //        cmdInsI_IP.CommandText = "insert into I_IP (ID, IP_DOC_NUMBER,IP_ID, IPNO_NUM, ID_DEBTCLS_NAME, IP_EXEC_PRIST, IP_EXEC_PRIST_NAME)";
        //        cmdInsI_IP.CommandText += "  VALUES (:ID, :IP_DOC_NUMBER, :IP_ID, :IPNO_NUM, :ID_DEBTCLS_NAME, :IP_EXEC_PRIST, :IP_EXEC_PRIST_NAME)";

        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_DOC_NUMBER", Convert.ToString(txtIP_DocNumber)));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_ID", Convert.ToDecimal(nIP_ID)));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":IPNO_NUM", Convert.ToDecimal(nIPNO_NUM)));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":ID_DEBTCLS_NAME", txtID_DEBTCLS_NAME));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_EXEC_PRIST", Convert.ToDecimal(nSUSER_ID)));
        //        cmdInsI_IP.Parameters.Add(new OleDbParameter(":IP_EXEC_PRIST_NAME", Convert.ToString(txtSUSER)));

        //        if (cmdInsI_IP.ExecuteNonQuery() == -1)
        //        {
        //            Exception ex = new Exception("Error inserting new row to I_IP table id = " + newID.ToString());
        //            throw ex;
        //        }


        //        cmdInsI_IP_OTHER = new OleDbCommand();
        //        cmdInsI_IP_OTHER.Connection = con;
        //        cmdInsI_IP_OTHER.Transaction = tran;
        //        cmdInsI_IP_OTHER.CommandText = "insert into I_IP_OTHER (ID, INDOC_TYPE, INDOC_TYPE_NAME, I_IP_OTHER_CONTENT)";
        //        cmdInsI_IP_OTHER.CommandText += "  VALUES (:ID, :INDOC_TYPE, :INDOC_TYPE_NAME, :I_IP_OTHER_CONTENT)";
        //        cmdInsI_IP_OTHER.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(newID)));
        //        cmdInsI_IP_OTHER.Parameters.Add(new OleDbParameter(":INDOC_TYPE", Convert.ToInt32(37)));
        //        cmdInsI_IP_OTHER.Parameters.Add(new OleDbParameter(":INDOC_TYPE_NAME", Convert.ToString("Сопроводительное письмо")));
        //        cmdInsI_IP_OTHER.Parameters.Add(new OleDbParameter(":I_IP_OTHER_CONTENT", txtContent));

        //        if (cmdInsI_IP_OTHER.ExecuteNonQuery() == -1)
        //        {
        //            Exception ex = new Exception("Error inserting new row to I_IP_OTHER  table id = " + newID.ToString());
        //            throw ex;

        //        }

        //        tran.Commit();
        //        con.Close();

        //        return newID;

        //    }
        //    catch (OleDbException ole_ex)
        //    {
        //        if (tran != null)
        //        {
        //            tran.Rollback();
        //        }
        //        if (con != null)
        //        {
        //            con.Close();
        //        }
        //        foreach (OleDbError err in ole_ex.Errors)
        //        {
        //            Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
        //        if (con != null)
        //        {
        //            con.Close();
        //        }
        //    }
        //    return -1;
        //}

        static ArrayList DirSearch(string sDir, int nCycle, int nMax)
        {
            ArrayList alFiles = new ArrayList();
            if (nCycle <= nMax)
            {
                try
                {
                    // почему-то не собираются файлы из первого запуска
                    foreach (string d in Directory.GetDirectories(sDir))
                    {
                        foreach (string f in Directory.GetFiles(d))
                        {
                            alFiles.Add(f);
                        }
                        //DirSearch(d);
                        alFiles.AddRange(DirSearch(d, nCycle + 1, nMax));
                    }

                }
                catch (System.Exception excpt)
                {
                    Console.WriteLine(excpt.Message);
                }
            }
            return alFiles;
        }

        // рекурсивная функция по получению списка файлов
        static ArrayList GetReestrs(string sDir, int nCycle, int nMax)
        {
            ArrayList alFiles = new ArrayList();
            try
            {
                alFiles = DirSearch(sDir, nCycle, nMax);
                foreach (string f in Directory.GetFiles(sDir))
                {
                    alFiles.Add(f);
                }

            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
            return alFiles;
        }


        static ArrayList GetLoadedReestrs(OleDbConnection ConParam, int iContrSource)
        {

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
                cmdIP.CommandText = "select distinct ish_number from GIBDD_PLATEZH where SOURCE_ID = " + iContrSource.ToString();
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
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
            }
            ConParam.Close();
            return result;
        }

        static ArrayList GetLoadedReestrs(OleDbConnection ConParam, int iContrSource, string txtDate)
        {

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
                cmdIP.CommandText = "select distinct ish_number from GIBDD_PLATEZH where SOURCE_ID = " + iContrSource.ToString() + " and DATE_ISH = '" + txtDate + "'";
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
                    Console.WriteLine("Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка приложения. Message: " + ex.ToString());
            }
            ConParam.Close();
            return result;
        }

        
        
        static int InsertDataToPlatezhTable(OleDbConnection conGIBDD, DataTable tbl, int iContrSource, string txtIshNumber, DateTime dtDateIsh)
        {
            Int32 iCnt = 0;
            OleDbTransaction tran;
            OleDbCommand m_cmd;
            string txtId = "";
            
            if ((conGIBDD == null) || (conGIBDD.State == ConnectionState.Closed))
                conGIBDD.Open();

            tran = conGIBDD.BeginTransaction(IsolationLevel.ReadCommitted);
            try
            {

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

                    m_cmd.CommandText = "INSERT INTO GIBDD_PLATEZH (NUMBER, NOMID, DATID, SUMM, SUMM_DOC, FIO_D, DATE_DOC, ISH_NUMBER, DATE_ISH, FL_USE, NUM_DOC, BORN_D, DATE_VH, SOURCE_ID)";
                    m_cmd.CommandText += " VALUES(:NUMBER, :NOMID, :DATID, :SUMM, :SUMM_DOC, :FIO_D, :DATE_DOC, :ISH_NUMBER, :DATE_ISH, :FL_USE, :NUM_DOC, :BORN_D, :DATE_VH, :SOURCE_ID)";


                    string txtNumber = Convert.ToString(row["Number"]).TrimEnd();

                    txtId = txtNumber;// для вывода ID строки в отладку ошибки

                    string txtGibddIDNumber = "";

                    // если это не реестр МВД, то ничего обрезать не нужно
                    if ((iContrSource == 1) && (txtNumber.Length > 7))
                    {
                        txtGibddIDNumber = txtNumber.Substring(1, 7);
                    }
                    else
                    {
                        txtGibddIDNumber = txtNumber;
                        // а это почему это?
                        //txtGibddIDNumber = "-1";
                    }

                    string txtDatID = Convert.ToString(row["Date_exec"]);
                    DateTime dtDatID;

                    Double nSum;
                    string txtSum = Convert.ToString(row["Summa"]);

                    string txtFioD = Convert.ToString(row["Plat_name"]).TrimEnd();
                    // TODO: нужно убрать все двойные пробелы из txtFioD, т.к там может быть бардак
                    txtFioD = RemoveDoubleSpaces(txtFioD, 200);
                    // и в верхний регистр букв
                    txtFioD = txtFioD.ToUpper();


                    string txtDateDoc = Convert.ToString(row["Date_doc"]);
                    DateTime dtDateDoc;

                    string txtNumDoc = Convert.ToString(row["Num_doc"]).TrimEnd().ToUpper();

                    string txtBornD = Convert.ToString(row["Date_plat"]);
                    DateTime dtBornD;

                    if (!DateTime.TryParse(txtBornD, out dtBornD))
                    {
                        dtBornD = Convert.ToDateTime("01.01.1800");
                    }

                    Double nSumDoc;
                    string txtSumDoc = Convert.ToString(row["Summa_doc"]);

                    //  если это не реестр из МВД, то никакого номера обрезать не нужно
                    if ((iContrSource == 1) && (txtGibddIDNumber[0] == '0'))
                    {
                        txtGibddIDNumber = txtGibddIDNumber.Substring(1, 6);
                    }

                    if (!DateTime.TryParse(txtDatID, out dtDatID))
                    {
                        dtDatID = DateTime.MinValue;
                    }

                    if (!Double.TryParse(txtSum, out nSum))
                    {
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
                    m_cmd.Parameters.Add(new OleDbParameter(":SOURCE_ID", iContrSource));   // источник - 1 МВД РК




                    int result = m_cmd.ExecuteNonQuery();

                    if (result != -1)
                    {
                        iCnt++;
                        // prbWritingDBF.PerformStep();
                    }
                }
                tran.Commit();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    Console.WriteLine("Ошибка при работе с данными. Номер строки = " + txtId + ". Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
                tran.Rollback();
                iCnt = -1;
                //return false;
            }
            catch (Exception ex)
            {
                //if (DBFcon != null) DBFcon.Close();
                Console.WriteLine("Ошибка приложения. Номер строки = " + txtId + ". Message: " + ex.ToString());
                iCnt = -1;
                //return false;
            }

            conGIBDD.Close();
            return iCnt;
        }

        static string RemoveDoubleSpaces(string txtString, int iMaxIter){
            
            int i = 0;
            while (txtString.IndexOf("  ") != -1)
            {
                txtString = txtString.Replace("  ", " ");
                i++;
                if (i > iMaxIter)
                {
                    break;
                }
            }
            return txtString;
        }

        static DataTable GetDbfTable(string txtSql, string tblName, string txtFilePath, string txtFileDir, string tablename){
            OleDbConnection DBFcon, DbaseCon;
            OleDbCommand m_cmd;

            bool bError = false;
            DataSet ds = new DataSet();
            DataTable tbl = ds.Tables.Add(tblName);

            try
            {
                //  ChangeByte(openFileDialog1.FileName, 0x65, 30);

                DBFcon = new OleDbConnection();
                //DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + openFileDialog1.FileName + ";Mode=Read;Collating Sequence=RUSSIAN;CODEPAGE=1251");
                DBFcon.ConnectionString = string.Format("Provider=VFPOLEDB.1;Data Source=" + txtFilePath + ";Mode=Read;Collating Sequence=RUSSIAN;CODEPAGE=1251");
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
                    Console.WriteLine("Ошибка при работе с данными. Будет предпринята повторная попытка обработать файл. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                }
                bError = true; // попробовать обработать через DBASE
            }

            if (bError)
            {
                try
                {

                    // если имя файла больше 8 символов - то копировать и обработат меньшее
                    string txtShortFileName = tablename;
                    if (tablename.Length > 8)
                    {
                        txtShortFileName = tablename.Substring(0, 8) + ".dbf";
                        if (!File.Exists(txtFileDir + txtShortFileName))
                            File.Copy(txtFilePath, txtFileDir + txtShortFileName);
                    }
                    else
                    {
                        txtShortFileName += ".dbf";
                    }

                    DbaseCon = new OleDbConnection();
                    DbaseCon.ConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBASE IV;Data Source={0}", txtFileDir);
                    DbaseCon.Open();
                    m_cmd = new OleDbCommand();
                    m_cmd.Connection = DbaseCon;
                    // это не совсем верно, но пока пусть будет как план B
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
                        Console.WriteLine("Ошибка при работе с данными. Файл обработать не удалось драйвером Dbase. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                    }
                }
            }

            return tbl;


        }
        static bool AutoLoadGibddReestrs(string constrGIBDD, string txtLogFileName, int iDiv, string txtOspEmail, string txtAdminEmail)
        {
            int iContrSource = 1;// источник - 1 МВД; 2 - РИЦ ЖХ
            decimal nUserID = 8992; // SYSDBA
            decimal nContrID = 0; // код контрагента (МВД, РИЦ ЖХ)

            OleDbConnection ConG, con;
            Int32 iCnt = 0;
            Int32 iPercent = 0;
            Int32 iFoundCnt = 0;
            Int32 iSourceID = 0;
            ArrayList alFiles;
            ArrayList alLoadedReestrs;
            string txtCurrPath = "";
            string txtMailServ = "mail10";

            try
            {
                ConG = new OleDbConnection(constrGIBDD);
                
                // 1. Получить список реестров
                //alFiles = DirSearch(txtUploadDirGibdd, 1, 10);
                alFiles =  GetReestrs(txtUploadDirGibdd, 1, 10);
                
                // 2.1 Получить список уже загруженных реестров ГИБДД
                alLoadedReestrs = GetLoadedReestrs(ConG, iContrSource); 
                
                // 2.2 Загрузить файлы из alFiles в GIBDD_PLATEZH, 
                foreach (string txtPath in alFiles)
                {
                    txtCurrPath = txtPath;
                    // вычленить имя файла
                    if (txtPath.Length > 0)
                    {
                        string tablename = "";
                        string txtExt = "";
                        string txtDateIsh = "";
                        DateTime dtDateIsh;
                        string txtFileDir = "";
                        string txtSql = "";


                        txtExt = txtPath.Substring(txtPath.LastIndexOf(".") + 1);
                        // если это dbf
                        if (txtExt.ToUpper().Equals("DBF"))
                        {

                            if (txtPath.Length > 4)
                                tablename = txtPath.Substring(0, txtPath.Length - 4);
                            tablename = tablename.Substring(tablename.LastIndexOf("\\") + 1);

                            // вычленить из имени файла дату исх. документа (реестра)
                            // проверить что имя правильной длины
                            tablename = tablename.Trim();
                            if (tablename.Length >= 8)
                                txtDateIsh = tablename.Substring(6, 2) + '.' + tablename.Substring(4, 2) + '.' + tablename.Substring(0, 4);

                            if (!DateTime.TryParse(txtDateIsh, out dtDateIsh))
                            {
                                dtDateIsh = DateTime.MinValue;
                            }

                            // получить каталог
                            txtFileDir = txtPath.Substring(0, txtPath.Length - tablename.Length - 4);

                            // вычленить из имени файла исх номер документа
                            string txtIshNumber = "";
                            //if (tablename.Length > 9)
                            //{
                            //    txtIshNumber = tablename.Substring(9, tablename.Length - 9);
                            //}

                            // новый формат txtIshNumber = дата_старый_исх.номер = tablename
                            txtIshNumber = tablename;

                            txtSql = "SELECT * FROM " + tablename + " WHERE LEN(RTRIM(NUMBER)) > 0";
                            
                            // проверить по имени файла что он не загружен
                            if (!alLoadedReestrs.Contains(txtIshNumber))
                            {
                                // если не загружен то загрузить bp DBF данные 
                                DataTable tblReestr = null;
                                tblReestr = GetDbfTable(txtSql, "GIBDD_PLATEZH", txtPath, txtFileDir, tablename);

                                // и вставить из в GIBDD_PLATEZH
                                if (tblReestr != null)
                                    iCnt = InsertDataToPlatezhTable(ConG, tblReestr, iContrSource, txtIshNumber, dtDateIsh);
                                if(iCnt >=0)
                                {
                                    // добавить в список alLoadedReestrs
                                    alLoadedReestrs.Add(txtIshNumber);
                                    string txtMessage = DateTime.Now.ToString() + " Загружено " + iCnt + "\tстрок реестра номер: " + txtIshNumber;
                                    txtMessage += "\n Реестр загружен из файла по пути: " + txtCurrPath;
                                    WriteTofile(txtMessage, txtLogFileName);

                                    SendEmail(txtMessage, "Загрузка реестра оплаченных штрафов МВД", txtOspEmail, txtAdminEmail, txtMailServ, "");
                                    SendEmail(txtMessage, "Загрузка реестра " + txtIshNumber + " оплаченных штрафов МВД ОСП " + iDiv.ToString().PadLeft(2, '0'), txtAdminEmail, txtAdminEmail, txtMailServ, "");
                                }
                            }
                        }
                    }
                }
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    Console.WriteLine(DateTime.Now.ToString() + " Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                    WriteTofile(DateTime.Now.ToString() + " Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, txtLogFileName);
                    string txtMessage = "Возникла ошибка при попытке загрузить реестр по пути " + txtCurrPath;
                    txtMessage += "\n" + DateTime.Now.ToString() + " Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState;
                    SendEmail(txtMessage, "Внимание! Ошибка загрузки реестра оплаченных штрафов МВД в ОСП " + iDiv.ToString().PadLeft(2, '0'), txtAdminEmail, txtAdminEmail, txtMailServ, "");
                }
                return false;
            }
            catch (Exception ex)
            {
                //if (DBFcon != null) DBFcon.Close();
                Console.WriteLine(DateTime.Now.ToString() + " Ошибка приложения. Message: " + ex.ToString());
                WriteTofile(DateTime.Now.ToString() + " Ошибка приложения. Message: " + ex.ToString(), txtLogFileName);
                string txtMessage = "Возникла ошибка при попытке загрузить реестр по пути " + txtCurrPath;
                txtMessage += "\n" + DateTime.Now.ToString() + " Ошибка приложения. Message: " + ex.ToString();
                SendEmail(txtMessage, "Внимание! Ошибка загрузки реестра оплаченных штрафов МВД в ОСП " + iDiv.ToString().PadLeft(2, '0'), txtAdminEmail, txtAdminEmail, txtMailServ, "");
                return false;
            }

            return true;
        }

        static int SendEmail(string txtContent, string txtSubj, string txtEmailTo, string txtEmailFrom, string txtSmtpAdr, string txtFileName)
        {
            try
            {
                //Авторизация на SMTP сервере
                SmtpClient Smtp = new SmtpClient(txtSmtpAdr, 25);
                //Smtp.Credentials = new NetworkCredential("login", "pass");
                //Smtp.EnableSsl = false;

                //Формирование письма
                MailMessage Message = new MailMessage();
                Message.From = new MailAddress(txtEmailFrom);
                Message.To.Add(new MailAddress(txtEmailTo));
                Message.Subject = txtSubj;
                Message.Body = txtContent;

                //Прикрепляем файл
                if (txtFileName.Trim() != "")
                {
                    Attachment attach = new Attachment(txtFileName, MediaTypeNames.Application.Octet);

                    // Добавляем информацию для файла
                    ContentDisposition disposition = attach.ContentDisposition;
                    disposition.CreationDate = System.IO.File.GetCreationTime(txtFileName);
                    disposition.ModificationDate = System.IO.File.GetLastWriteTime(txtFileName);
                    disposition.ReadDate = System.IO.File.GetLastAccessTime(txtFileName);

                    Message.Attachments.Add(attach);
                }

                Smtp.Send(Message);//отправка
            }
            catch (Exception ex)
            {
                return 0;
            }
            return 1;
        }


        static bool Sverka(string txtConStrPKOSP, string constrGIBDD, decimal mvd_id, string txtLogFileName)
        {
            // нужно проводить сверку в 2 этапа - сначала 1 - МВД, потом 2 - РИЦ ЖХ, теоретически, можно все объединить, нужен справочник только.
            // в справочнике пусть будет entt_id контрагента 
            int iContrSource = 2;// источник - 1 МВД; 2 - РИЦ ЖХ
            decimal nUserID = 8992; // SYSDBA
            decimal nContrID = 0; // код контрагента (МВД, РИЦ ЖХ)

            string tablename = "GIBDD_PLATEZH";
            OleDbConnection ConG, con;
            Int32 iCnt = 0;
            Int32 iPercent = 0;
            Int32 iFoundCnt = 0;
            Int32 iSourceID = 0;
            try
            {
                ConG = new OleDbConnection(constrGIBDD);
                con = new OleDbConnection(txtConStrPKOSP);

                // удалить все учтенные и все использованные
                // не нужно этого делать - пусть будут в базе - хоть откатить и проверить если что будет возможность
                DeleteUsedGibddPlat(ConG, true, true);

                DataSet ds = new DataSet();
                // основная выборка
                string txtSql = "SELECT * FROM " + tablename + " WHERE  FL_USE = 0 ORDER BY SOURCE_ID"; // выбрать только те, которые еще не были рассмотрены
                
                // дополнительная выборка
                // string txtSql = "SELECT * FROM " + tablename + " WHERE  FL_USE = 0 and nomid='027247' ORDER BY SOURCE_ID"; // выбрать только те, которые еще не были рассмотрены

                
                //string txtSql = "SELECT * FROM " + tablename + " WHERE  FL_USE = 0 AND SOURCE_ID = " + iContrSource.ToString(); // выбрать только те, которые еще не были рассмотрены
                DataTable tbl = GetDataTableFromFB(txtSql, tablename, constrGIBDD);
                if (tbl != null)
                {
                    Console.WriteLine(DateTime.Now.ToString() + " В реестре из ГИБДД для сверки содержится " + tbl.Rows.Count.ToString() + " строк.");
                }

                string txtContent = "";

                //AlterIndxI_ID(con, true);

                WriteTofile(DateTime.Now.ToString() + "проводится сверка.", txtLogFileName);

                foreach (DataRow row in tbl.Rows)
                {
                    iCnt++;
                    if (iPercent <= 100 * iCnt / tbl.Rows.Count)
                    {
                        //Console.Clear();
                        iPercent++;
                        Console.Write("\b\b\b" + iPercent.ToString() + "%");
                    }
                    
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
                    
                    string txtSourceID = Convert.ToString(row["SOURCE_ID"]).TrimEnd();
                    Int32.TryParse(txtSourceID, out iSourceID);
                    
                    // сделать запрос в базу ПК ОСП для поиска такого номера ИД

                    decimal id = 0;


                    if (iSourceID == 1) // если это МВД - то это один вид SQL select
                    {
                        id = FindIDNum(con, txtNomID, nSumID, dtDatID);
                    }
                    else if (iSourceID == 2) // если это РИЦ ЖХ - то это другой вид SQL select
                    {
                        id = FindIDNumRicZH(con, txtNomID, nSumID, dtDatID, txtFIO_D);
                    }

                    if (id > 0)
                    {
                        iFoundCnt++;

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
                        txtContent = "Должник " + txtFIO_D + " ";
                        if (!dtBornD.Equals(Convert.ToDateTime("01.01.1800")))
                        {
                            txtContent += "(дата рождения " + dtBornD.ToShortDateString() + ") ";
                        }
                        txtContent += dtDateDoc.ToShortDateString() + " оплатил " + Money_ToStr(nSumDoc) + " № документа об оплате " + txtNumKvit + " по ИД № " + txtNomID + " от " + dtDatID.ToShortDateString() + ".";
                        
                        //MessageBox.Show("Нашли ИД № " + txtNomID + ". IP_ID = " + id.ToString(), "Внимание!", MessageBoxButtons.OK);
                        string txtNumber = Convert.ToString(row["NUMBER"]);
                        nContrID = GetContrID(ConG, iSourceID);

                        // I_OP_CS_ENDDBT - Новый (1)
                        // I_IP_OTHER - Зарегистрирован (2)

                        ID_InsertPlatDocTo_PK_OSP(con, 1, nUserID, nSumDoc, DateTime.Today, id, dtDateDoc, txtNumKvit, nContrID, txtFIO_D);
                        UpdateGibddPlatezh(ConG, txtNumber, 1, iSourceID);
                        ID_InsertOtherIP_DocTo_PK_OSP(con, 2, nUserID, dtReestrVhodDate, id, dtExtDocDate, txtExtDocNum, txtContent, nContrID);
                        WriteTofile(DateTime.Now.ToString() + " " + txtContent + " IP_ID = " + id.ToString() + ".", txtLogFileName); // пишем в лог
                    }

                }

                //AlterIndxI_ID(con, false);

                Console.WriteLine(DateTime.Now.ToString() + " Данные успешно проверены. Найдено " + iFoundCnt.ToString() + " ИД.");
                
                WriteTofile(DateTime.Now.ToString() + " сверка окончена успешно. Найдено " + iFoundCnt.ToString() + " ИД.", txtLogFileName);

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    Console.WriteLine(DateTime.Now.ToString() + " Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState);
                    WriteTofile(DateTime.Now.ToString() + " Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState, txtLogFileName);
                }
                return false;
            }
            catch (Exception ex)
            {
                //if (DBFcon != null) DBFcon.Close();
                Console.WriteLine(DateTime.Now.ToString() + " Ошибка приложения. Message: " + ex.ToString());
                WriteTofile(DateTime.Now.ToString() + " Ошибка приложения. Message: " + ex.ToString(), txtLogFileName);
                return false;
            }

            return true;
        }

    }

    class FormUpdater
    {
        public static string my_version = "1.16.05";
        // 1.00 - первый вариант
        // 1.01 - для теста системы обновления
        // 1.02 - Автозагрузка реестров ГИБДД - тест
        // 1.03 - Автозагрузка реестров ГИБДД - релиз
        // 1.04 - + " WHERE LEN(RTRIM(NUMBER)) > 0"
        // 1.05 - переход на класс PKOSP_mvv
        // 1.06 - сверка ГИБДД многофакторная
        // 1.06.1 - исправление ошибки с обновлением и разноска по дням недели сверки ГИБДД
        // 1.06.2 - DeleteOld (удаление старше 2 лет несверенных)
        // 1.07 - Автозагрузка КРЦ + устранение ошибки с UpdatePlat
        // 1.07.01 - исправление ошибок сверки ГИБДД (ФИО теперь там где есть LIKE короткий номер), КРЦ автозагрузку ВЫКЛ
        // 1.07.02 - исправление ошибок сверки ГИБДД (в четверг самая быстрая сверка (только 1 и 2)). Поправлена загрузка новых реестров для номеров нового формата.
        // 1.07.03 - ускорение загрузки квитанций из Сбера за счет правки поиска restriction.id по acc.id
        // 1.07.04 - DateTime dtPeriod = DateTime.Today.AddMonths(iMonthPeriod);
        // 1.07.05 - автозагрузка сбер и КРЦ
        // 1.07.06 - автозагрузка  КРЦ отключена, т.к. она актуальна только для Петрозаводска, + исправлена описка в названии темы e-mail "их Сбербанка"
        // 1.07.07 - починил загрузку квитанций Сбербанка lLogger2 и update_filename 
        // 1.07.08 - заглушка при загрузке квитанций (статус 10 заменить на 1), если файл бинарный - записать в лог со статусом 0 и именем файла, чтобы больше его не пытаться проверять
        // 1.07.09 - убрал в PKOSP_mvv.RecFindIDNumMVD в txtSql лишние '% и лишние "" + "
        // 1.08 - автомат для выгрузки запросов в Сбер и загрузки ответов по XML форматам с пасп. данными
        // 1.08.01 - переставил выгрузку запросов и загрузку ответов сбер по XML форматам с пасп. данными до загрузки квитанций попост. сбера
        // 1.08.01 - поправил ошибку с выгрузкой серии паспорта в запросах в сбер по XML форматам с пасп. данными// 1.08.01 - поправил ошибку с выгрузкой серии паспорта в запросах в сбер по XML форматам с пасп. данными
        // 1.08.02 - AutoLoadSberKvit - улучшено логирование и добавлена отправка сообщения об ошибке на e-mail
        // 1.08.03 - packstatus заменил на pack_status в GetLoadedReestrs
        // 1.08.04 - теперь как EXT_REPORT грузятся только квитанции об отказе, ВСЕ квитанции грузятся в EXT_RECEIPT
        // 1.08.05 - реестры с 7 символами больше не обрабатываются, на все что касается ошибок имени файла уходят в почту ошибки файл помечается как обработанный
        // 1.08.06 - если квитанций в файле больше 0 а загружено 0, то написать email что Ошибка, нужно проверить данные.
        // 1.08.07 - уведомление о квитанции из Сбера с большим (более 50%) количеством отказов
        // 1.08.08 - в связи с тем, что много отказов и если в постановлении счета 2 раза указываются, то приходит отказ  временно убираю загрузку отказов совсем! ЛТП 1 № 16021
        // 1.08.09 - отказы не грузим, все квитанции грузим в EXT_REPORT, срабатываение уведомления о количестве отказов на уровене выше 70%
        // 1.09 - починили выборку параметров ИП в связи с релизом 1.21  и переносом поля ID_DEBTTEXT в DOC_IP
        // 1.09 - из сверки ГИБДД убрал все способы сверки кроме прямого попадания по длинному номеру или по коротокому + дата + сумма
        // 1.09.1 - платежный документ сверки ГИБДД теперь может регистрироваться
        // 1.10 - поменял алгоритм сверки ГИБДД, т.к. выяснилось что номера ИД в разных ИП могут совпадать и нужно учитывать ФИО и дату ИД
        // 1.10.01 - период поиска загруженных реестров сократил с 11 мес. до 6, иначе давал сбои
        // 1.10.02 - исправлена ошибка - когда нет подключения к ПК ОСП - сразу Ex без попыток что-то обработать.
        // 1.10.02 - pack_status = 1000 - если это бин. файл
        // 1.10.02 - pack_status = 2000 - если это ошибка в имени файла
        // 1.10.03 - убрал проверку IsBinary для квитанций Сбера - чтобы не было ложных отказов
        // 1.10.04 - исправил ID_InsertPlatDocTo_PK_OSP + ID_InsertOtherIP_DocTo_PK_OSP в части параметра :ID_DEBTCLS - вместо :ID_DEBTCLS_NAME
        // 1.11    - релиз с функциями bAutoLoadReport - отчеты о результатах обработки в Сбере,
        //           bAutoLoadVU - загрузка сведений о ВУ из ГИБДД
        // - при вводе в эксплуатацию не забыть при загрузке реестра ВУ из ГИБДД в одной транзакции сначала узнать какого года вставляем данные
        // потом удалить все записи с этим годом и потом вставить новый реестр

        // 1.12    - релиз с отключенной сверкой ГИБДД
        //  bRunGibdd10_selfrequest  - запрос о ВУ по базе ГИБДД
        //  bRunIC_MVD_out - выгрузка запросов ИЦ МВД
        //  bRunIC_MVD_in - загрузка ответов ИЦ МВД
        //  каталоги rsync делает ksv в рамках ЛТП 1 № 22092. папки для rsync инф. обмен IC_MVD

        // 1.13    - сверка с включенной обратно сверкой, но это уже новый алгоритм для МВД - Sverka2
        // 1.14 - включена функция автозагрузки сведений о ВУ в ufssprk-tools, вставка сведений о сверке ГИБДД в ИТ (Sverka2) временно отключена.
        // также в функции UpdateGibddPlatezhLong !!! 20160506 - убрал параметр ISH_NUMBER - чтобы повторы все за 1 раз ушли

        // 1.14.01 запрос о ВУ теперь через локальные базы, в UpdateGibddPlatezhLong  вернул параметр ISH_NUMBER - т.к. один раз без него отработало, больше не нужно
        // 1.14.02 - тестовая сборка для Медгоры и Сегежи.
        // включены вставка EXT_REPORT, EXT_DEBT_FIX
        //
        // для всех ОСП включена вставка EXT_IDENTIFICATION_DATA (на тесте не работает - точные причины выяснить пока не удалось)
        
        // 1.14.03 - добавлена загрузка прав на управление маломерным судном из ГИМС и автоответы на запросы в ГИМС 
        // отключены вставка EXT_REPORT, EXT_DEBT_FIX

        // 1.14.04 - добавлена дополнительная выборка при загрузке квитанций Сбера - если не найдено через GetActIdByAccID,
        // то ищем через GetActIdByAccIpNo (последниее по дате постановление из электронных где был нужный счет)
        // это позволит загрузить то, что перестало работать когда почистили ИТ в Петр1, Петр2
        // включены вставка EXT_REPORT, EXT_DEBT_FIX для всех навсегда

        // 1.14.05 - исправил select  в GetActIdByAccIpNo
        // добавил  o.doc_electron = 1  and o.contr_name = 'Центр сопровождения клиентских операций \"Волга-Сити\" ОАО \"Сбербанк России\"'

        // 1.14.06 - выгрузка запросов в ИЦ МВД только по понедельникам
        // 1.14.07 - переработана загрузка сведений из ГИБДД-МВД, проверяется поле Plat_name на наличие не кириллических символовю
        // если видим что символ не кириллический, то читаем эту строчку из файла зановоб но уже в другой (DOS866) кодировке.
        // 1.14.08 - в Sverka2 добавлен поиск по ИНН (2 способа внутри рекурсии). Лолжно искать штрафы на юриков, но и для физ. лиц может помочь, если вдруг ФИО неверно вбито,
        // потому что сначала ищет по ИНН, а потом по ФИО
        // 1.14.09 - WriteIcMvdReqRow
        // txtBirthPlace  поле место рождения не должно превышать 90 знаков - поэтому отрезаем запятые (они им не нужны) и режем с конца если все еще больше чем 90
        // 1.14.10 - в загрузке ответов IC_MVD найдена ошибка - не там был con.Dispose() - переставил после цикла обработки всех файлов
        // 1.15.00 - автоматическая выгрузка постановления в Сбер (рег.МВВ) + копирование zss файла
        // 1.15.01 -  выгрузка из ИТ в постановления Сбера - вместо acc.summa теперь выгрудается ext_restriction.ip_rest_debtsum
        // 1.15.02 -  исправлена выгрузка постановлений Сбера из ИТ - добавлен учет вынесения в сводном по должнику
        // 1.15.03 -  исправлена ошибка из 1.15.02 - забыл указать алиас old_regnumber после coalesce  в новой выборке - поэтому солбец с таким именем не был найден в таблице
        // 1.15.03 -  также закомментировал выгрузку ограничений со старым номером ИП в отдельный файл, т.к. больше сбер не хочет такое
        // 1.15.04 -  сбер вернулся к отдельному файлу для ограничений, но теперь это последний файл за день - внедрено в пром. эксплулатацию
        // 1.15.04 -  bRunCredOrgReqOut добавлена функция автоматического формирования запросов в кред.орги по понедельникам
        // 1.15.05 - все новые функции по-умолчанию отключены. В ReadGibdd10Zapros исправлена ошибка из-за которой не работали ответы на запросы о ВУ.
        // 1.15.06 - включена автовыгрузка запросов в cred_org и potd
        // 1.15.06 - включена автозагрузка ответов из cred_org и potd
        // 1.15.06 - для 10013, 10024 сделано автоформирование пути для //fs/Inf_obmen - теперь ненужно конфиги править, ура!
        // 1.15.06 - если код отдела 0 - то ничего из нового по-умолчанию не включается (=false);
        // 1.15.07 -  в загрузке ответов potd вместо ИЦ МВД исправлено верно - Console.WriteLine("Итого загружено ответов ПФ расширенные (XML): " + iTotalFiles.ToString());
        // 1.15.08 - добавлена проверка загружен ли fox622.exe и 2 dll к нему. если их нет - то загружаются.
        // 1.15.09 - загрузка fox622.exe и 2 dll к нему сделана обычной (не асинхронной)
        // 1.15.10 - проверка загружен ли fox622.exe и 2 dll к нему теперь происходит всегда при запуске программы
        // 1.15.10 -  выгрузка кред. орг запросов отдельно для вторника 28.02.2017 - т.к. пропустили понедельник
        // 1.15.11 - включена автозагрузка реестров ТГК-1
        // 1.15.12 - сверка ТГК-1 доработана - теперь работает не только поиск по номеру ИП
        // 1.15.12 - по-умолчанию включена загрузка реестров ТГК-1 если подключено к РБД. путь загрузки указан  = \\fs\inf_obmen_rozjsk\ответ\TGK-1
        // 1.15.13 - сделана замена суммы для постановлений в СБер - теперь это total_debt_sum
        // 1.15.13 - функционал по ГИМС и РТН недоделан (InsertExtTransportGimsIntTable) - поэтому пока по-умолчанию выключен
        // 1.15.14 - внедрен РТН и ГИМС РК Лодки bRunGimsLodka10_selfrequest = true;
        // 1.15.14  - забыл в прошлый раз - теперь и в ReadSberOldIpRestrictions сделана замена суммы для постановлений в СБер - теперь это total_debt_sum
        // 1.15.14  - в ReadSberOldIpRestrictions и ReadSberRestrictions // 20170425 - нужно добавить выборку отмены ареста и отмены обращения взыскания без галочки В электронном виде
        // 1.15.15  - для ОСП 10007 Лахд сделана замена имеин файл сервера на rdb-lahd
        // 1.15.15  - при загрузке ответов ПФ расширенных сделана корректировка поиска parent_id - замена в имени файла, т.к. не искал до этого
        // 1.15.16  - добавил вместо флага b20reliz строку - if (txtZaprosOut.Equals("1")) day = 1; чтобы выгружать через параметр в 5 строке как в понедельник запросы
        // 1.15.17  - автозагрузка реестров РИЦ ЖХ (почему-то ее не было раньше)
        // 1.15.18  - в Sverka добавил вставку документа об оплате долга через ИТ (вместо напрямую системных)
        // 1.15.19  - добавил в InsertExtDebtFix при вставке EXT_DEBT_FIX столбец  IS_PAYING_OFF = 1
        // 1.15.20  - добавил в Sverka вычисление nPackNumber для вставки в ИТ
        // 1.15.20 -  т.к. ext_input_header.pack_number(BIGINT) д.б. уникальным для каждого КПС
        // 1.15.20 -  то нужно менять pack_id при вставке - предлагаю добавлять разряд на основе iSourceID
        // 1.15.20 -  по формуле log_id + 1000000000*iSourceID  (iSourceID встанет первыми цифрами)
        // 1.15.21 -  добавил в текст EXT_REPORT для Sverka информацию о коде контрагента
        // 1.15.22 -  ReadSberOldIpRestrictions, ReadSberRestrictions (20170616) - готовимся к фед. эдо - будем только отмены выгружать в этой сборке
        // 1.15.23 - прописал fs-pryag для пряжи т.к. fs не разрешался
        // 1.15.24 - изменил в mvv.Sverka проверку для КРЦ и ТГК-1 в части:
        // - Проверки даты ИД, кроме i_id.id_docdate (даты выдачи ИД) еще нужно смотреть i_id.id_des_date (дата принятия решения по делу)
        // - При поиске ИП смотрим на статус ИП - берем только ИП на статусах ЖЦ в исполнении ip_d.docstatusid in (4, 9, 22, 23, 24)
        // - в вывод отчета EXT_REPORT в текст вставлять исх номер реестра, из которого получена информация.
        // 1.15.25 -  убрал проверку (iRowsW >= 0), т.к. получались ложные срабатывания если 0 был обработан, а была ошибка.
        // сделал if ((iRowsW > 0) || ((iRowsW == 0) && (dtXLS != null) && (dtXLS.Rows.Count == 0)))
        // 1.15.26 - RecFindIDNumKRC, RecFindIDNumTGK1  добавляем DateTime dtDatePlat
        // дату оплаты чтобы потом искать только те ИП, которые были оплачены до регистрации ИД
        // 1.15.27 - выгрузка 04.10.2017 специальной выгрузки в ПФ Колчанову
        // 1.15.28 - сделал изменение в выгрузке в ПФР колчанову - чтобы логи обновлялись один раз а не 300000 раз на каждую строку
        // 1.15.29 - настройки fs для Бел(2) и Пуд(15). Начал писать AutoLoad_PFR10_Otvet но не окончил - пока не включено.
        // 1.15.30 - сделано AutoLoad_PFR10_Otvet c 23 по 29.10.2017 будет загружать ответы ПФР о месте работы 2017
        // 1.15.30 - отчеты по выгрузке 1ss файлов в ОСП будут копией идти на nadezhda.smirnova1@r10.fssprus.ru
        // 1.15.31 - txtConStrED вместо ed10 вписал ip-адрес 10.10.4.243
        // 1.15.31 - исправил загрузку ответ на ПФР - теперь если не найден ЕГРЮЛ то адрес пустой а имя как в ПФР
        // 1.15.31 - при загрузке проверяется через ИТ был ли ответ на запрос уже загружен
        // 1.15.32 - исправил обнуление счетчиков iCnt = 0; iECnt = 0; в AutoLoadCredOrgOtvet (была ошибка - суммировалось по всем новым файлам)
        // 1.15.33 - загрузка реестров VU теперь с проверкой повторов раз в 120 мес - а если быть умнее - то я сам имя генерируюб можно вобще убрать это дело
        // 1.15.33 - также отключены загрузка и выгрузка по запросам о месте работы ПФР_10 2017
        // 1.15.34 - для ПФ расширенного запроса и запроса в кред.органищации добавлена выгрузка по вторникам
        // 1.15.34 - если в понедельник не было ничего выгружено if (nTomorrowUnloaded.Equals(0)) day = 1;
        // 1.15.35 - версию ПО записываю в поле FILENAME лога с conv_code SverkaGibdd
        // 1.15.35 - выгрузку по вторникам также и для ИЦ МВД (было только кред_орг и ПФ расширенный (potd))
        // 1.15.35 - выгрузку по вторникам if (nTomorrowUnloaded.Equals(0)) day = 1; исправлено чтобы только во вторник запускать - было в любой день
        // 1.15.36 - для ALL_CRED_ORG запросов сделать UpdateLLogStatus(2) если pack_count > 0
        // 1.15.37 (исправил то что было в 1.15.30) - отчеты по выгрузке 1ss файлов в ОСП будут копией идти на nadezhda.smirnova1@r10.fssprus.ru, т.к. не работало
        // 1.15.38 - В 10024 сделали обычный \\fs\Inf_obmen - исправил путь который зашит в код программы
        // 1.15.38 - В 10013 сделали обычный \\fs\Inf_obmen - исправил путь который зашит в код программы
        // 1.15.39 - В 10013 вернул как было т.к, \\fs\Inf_obmen - не настроено
        // 1.15.40 - Program.cs - (2486, 29) : if (bRunCredOrgReqOut):        if ((day == 1) || (DateTime.Today.ToShortDateString().Equals("13.12.2017")))
        // 1.15.40 - выгрузить запросы кред. орг  "13.12.2017"
        // 1.15.41 - выгрузить запросы кред. орг  "14.12.2017"
        // 1.15.42 - выгрузить запросы кред. орг  "21.12.2017" и "21.12.2017"
        // if ((day == 1) || (DateTime.Today.ToShortDateString().Equals("21.12.2017")) || (DateTime.Today.ToShortDateString().Equals("22.12.2017")))
        // 1.15.43 - case 13: txtPathBase = "\\\\fb-petr13\\Inf_obmen\\"; break;
        // 1.15.44 - исправлена загрузка ответов от кредю орг-ий. 
        // 1.15.45 - убрал временное переназначение txtPathBase в выгрузке постановлений в Сбер
        // 1.15.46 - string txtCistomLegalIds убрали 86200999999007 - БаренцБанк, в 21 новый путь case 21: txtPathBase = "\\\\rdb-petr3.petr3.karelai.ssp\\Inf_obmen\\"
        // 1.15.47 - //txtAgreementCode = "170"; - > txtAgreementCode = "190"; 170 - Баренц, 190 - Связьбанк
        // 1.15.47 - т.к. больше ни от МВД ни от ЖКХ у нас нет информации - то сверку больше не делаем bRunSverka = false;
        // 1.15.48 - больше не грузим ничего из ГИБДД - поэтому блокируем bAutoLoadGibdd = false;
        // 1.15.49 - в 1 новый путь case case 1: txtPathBase = "\\\\rdb-petr1.petr1.karelia.ssp\\Inf_obmen\\"; break;
        // 1.15.50 - в 1 новый путь case case 1: txtPathBase = "\\\\rdb-petr1.petr1.karelia.ssp\\Inf_obmen\\"; break;
        // 1.15.51 - AutoLoadPotdOtvet int nMonthPeriod = -100000; // имена уникальные - не нужно ничего чистить
        // 1.15.52 - AutoLoadPotdOtvet int nMonthPeriod = -1000; // -100000 слишком много для датыю сделали -1000 
        // 1.15.53 - // txtCistomLegalIds = "86200999999009"; оставили только Связь-Банк (Баренц убрали); для bRunCredOrgAnsIn int nMonthPeriod = -1000; // имена уникальные - не нужно ничего чистить
        // 1.15.54 - case 20: txtPathBase = "\\\\rdb-petr2.petr2.karelia.ssp\\Inf_obmen\\"; break;
        // 1.15.55 - в if (bRunRtn10_selfrequest) неверно вставлялся ответ для юлиц и ИП с ИНН - ошибка с внешним ключом была - есть скрипт для исправления
        // 1.15.56 - string txtMailServ = "mail10"; + убрал .karelia.ssp
        // 1.15.56 - убираем домен из conStr txtConStrED = ff.RemoveDomainFromConString(txtConStrED);
        // 1.15.57- исправил ff.RemoveDomainFromConString
        // 1.15.58- исправил если не указан порт в ff.RemoveDomainFromConString - if (!txtSrc.Contains("/3051:")) txtSrc = txtSrc.Replace(":", "/3051:");
        // 1.15.59 - в 11 muez новый путь case (case 11: txtPathBase = "\\\\rdb-muez\\Inf_obmen\\"; break;)
        /* 1.15.60 - 
            bAutoLoadVU = false; отключаем т.к по ВУ теперь Фед. МВВ и не нужно ничего загружать 
            bAutoLoadKRC = false; // отключено т.к. нет больше обмена с КРЦ
            bAutoLoadRicZH = false; // отключено т.к. нет больше обмена с РИЦ ЖХ
            bAutoLoadKESK = false; // отключено т.к. нет больше обмена с КЭСК
            bAutoLoadTGK1 = false; // отключено т.к. нет больше обмена с ТГК-1
            bAutoSberReqOut = false; // не выгружаем запросы в Сбер по Рег.МВВ
            bAutoSberRespIn = false;  // не загружаем ответы в Сбер по Рег.МВВ
            bRunGibdd10_out = false; // этого обмена нет вобще
            bRunCredOrgReqOut = false; // в феврале 2019 вывели Связь-Банк из обмена т.к. перешли на ФЕД.Мвв, больше нет Рег.МВВ банков
            bRunCredOrgAnsIn = false; // в феврале 2019 вывели Связь-Банк из обмена т.к. перешли на ФЕД.Мвв, больше нет Рег.МВВ банков;
        */
        // 1.15.61 - в 06 B 24  новый путь case (case 6: txtPathBase = "\\\\rdb-kost\\Inf_obmen\\"; break;) (case 24: txtPathBase = "\\\\rdb-mosp24\\Inf_obmen\\"; break;)
        // 1.15.62 - в 13 новый путь case 13: txtPathBase = "H:\\"; break; net use H: \\fb-mosp13\Inf_obmen /user:staskevich@r10 "<пароль>" /persistent:Yes
        // 1.15.63 - Добавил ACT_DATE в FindRtnFlDt и в  # region "RTN_10_SELFREQUEST" txtResp += "Дата актуальности сведений: " + txtAct_date + "\r\n";
        // 1.15.63 - Добавил ACT_DATE в FindRtnUlDt и в  # region "RTN_10_SELFREQUEST" txtResp += "Дата актуальности сведений: " + txtAct_date + "\r\n";
        // 1.15.63 - Добавил ACT_DATE в # region "ГИМС_10_SELFREQUEST" для ФЛ и для ЮЛ txtResp += "Дата актуальности сведений: " + txtAct_date + "\r\n";
        // 1.15.64 - проверка доступности базы ufssprk-tools
        // 1.15.65 - case 21: txtPathBase = "U:\\"; break; net use U: \\rdb-petr3\Inf_obmen /user:guest@r10 "1" /persistent:Yes
        // 1.15.66 - case 20: txtPathBase = "T:\\"; break; net use T: \\rdb-petr2\Inf_obmen /user:guest@r10 "1" /persistent:Yes
        // 1.15.67 - case 24: txtPathBase = "W:\\"; break;
        // 1.15.68 - в запросе Перс данных убрал join на doc_ip_doc - чтобы выбирало и сводные тоже 
        // 1.15.69 - case 1: txtPathBase = "X:\\"; break;
        // 1.15.70 - включили ПФР_10_ПИЭВ для отчетов bRunPfrRepIn = true
        // 1.15.71 - в обработчике RunPfrRepIn добавил учет префикса opfrm_86 (добавлена буква m) - это для отказов в ручном режиме от операторов.
        // 1.15.72 - получить отмены фед.МВВ которые не ушли dtFedMvvSberEndGaccount = mvv.ReadFedMvvSberEndGaccount(conPK_OSP, lLogger);
        // 1.15.72 - все отмены с 01.11.2019
        // 1.15.73 - вернул обратно dtRestrictionsSberOldIp = mvv.ReadSberOldIpRestrictions(conPK_OSP, lLogger);
        // 1.15.74 - отключаем dtFedMvvSberEndGaccount = mvv.ReadFedMvvSberEndGaccount(
        // 1.15.74 - т.к. возникли проблемы - почему-то повторно выгружается, наверное update не работает
        // 1.15.75 - поменял даты для выгрузки фед. отмен в Сбер с 01.09.2019 по 01.11.2019
        // 1.15.76 - исправил синтаксическую ошибку в select в ReadFedMvvSberEndGaccount
        // 1.15.77 - в select в ReadFedMvvSberEndGaccount добавил учет доп. сведений о ДС на статусе Ошибочные и убрал верхнюю границу даты, нижняя 01.09.2019
        // 1.15.78 - убрал лишний and в select в ReadFedMvvSberEndGaccount и добавил костыли для O_IP_ACT_ZP в обработке квитанций от ОПФР
        // 1.15.79 - снял с выборки записи ReadFedMvvSberEndGaccount где ID_DOCDATE is null, D.DOC_DATE > '04.07.2019' - это дата внедрения релиза 19.1 в большинстве ОСП
        // 86211548439898
        // 1.15.80 - в InsertSberDocRowToTxt2 изменил вывод номера ИП. теперь если есть старый номер ИП то он пишется в поле Обычного номера ИП, т.к. СБЕР не принимает доп. параметр
        // 1.15.81 - в InsertSberDocRowToTxt2 исправил ошибки
        // 1.15.82 - вернул выгрузку запросов в ИЦ МВД с группировкой по ФИО, дате, месту рождения
        // 1.15.83  - вернул временно на место bool bMultiRequest = false;
        // 1.15.84 - выгрузка в ИЦ МВД не более 4 ID в списке
        // 1.15.85 - отключение обработки запросов
        // bRunGims10_selfrequest = false;  // 20210422 - отмена загрузки ответов на запросы о вод. удостоверений ГИМС
        // bRunRtn10_selfrequest = false; // 20210422 - отмена обработки запросов в РТН
        // bRunGimsLodka10_selfrequest = false; // 20210422 - отмена обработки запросов в ГИМС о лодках
        // bRunGibdd10_selfrequest = false; // 20210422 - отмена обработки запросов о ВУ в ГИБДД
        // 1.15.86 - ГИБДД ВУ оставил работать
        // 1.15.87 - отключено все кроме ГИБДД ВУ, ИЦ МВД запросы-ответы и квитанции на постановления ПФР РегМВВ
        // 1.15.88 - для AutoLoadIcMvdOtvet nMonthPeriod = -12000; // чистим раз в 1000 лет - то есть никогда
        // 1.15.89 - добавлен костыль - чтобы не обходить районы в связи со сменой пароля добавлены 2 параметра (старый, новый пароли) и в ConnectionString делается замена
        // 1.15.90 - исправил ошибку - вызова функции GetOSP_Num.Ошибка при работе с данными. Message: Ошибка подключения к базе данных. пароль SYSDBA нужно менять до вызова этой функции
        // 1.16.00 - добавлена функциональность автоответов на запросы Рег. МВВ в ИЦМВД 10 ФНС о паспортных даных
        // 1.16.01 - исправлена ошибка в select при поиске ДУЛ
        // 1.16.02 - // 20230830 - временно отключил т.к. нужно широфвание и авторизация // Smtp.Send(Message);//отправка
        // 1.16.03 - при поиске паспорта проверять длину серии = 4 и длину номера = 6
        // 1.16.04 - в FindDul вместо DOCUMENT.DOCSTATUSID = 9
        // добавлено (DOCUMENT.DOCSTATUSID = 9  or (DOCUMENT.DOCSTATUSID = 10 and dateadd(year, -2, current_date) < DOC_IP.IP_DATE_FINISH))
        // 1.16.05 - в FindDul вместо DOCUMENT.DOCSTATUSID = 9
        // добавлено (DOCUMENT.DOCSTATUSID = 9  or (DOCUMENT.DOCSTATUSID in (10, 12) and dateadd(year, -2, current_date) < DOC_IP.IP_DATE_FINISH))
               

        // Ссылки для скачивания
        private string url_version = "http://web/updater/sverka_gibdd/version.txt";
        private string url_program = "http://web/updater/sverka_gibdd/PickupBases.exe";
        private string url_foruser = "http://web/updater/sverka_gibdd/index.php";
        private string url_fox622 = "http://web/updater/sverka_gibdd/fox622.exe";
        private string url_VFP6RENU = "http://web/updater/sverka_gibdd/VFP6RENU.DLL";
        private string url_VFP6R = "http://web/updater/sverka_gibdd/VFP6R.DLL";
        private string file_fox622 = "fox622.exe";
        private string file_VFP6RENU = "VFP6RENU.DLL";
        private string file_VFP6R = "VFP6R.DLL";

        private string my_filename;
        private string up_filename;
        private string txtLogFileName;

        // Признак, что началось скачивание обновления, требуется ожидание завершения процесса
        private bool is_download; public bool download() { return is_download; }

        // Признак, что обновление не требуется или закончено, можно запускать программу.
        private bool is_skipped; public bool skipped() { return is_skipped; }


        public FormUpdater(string[] argc, string txtLogFileNameParam)
        {
            // проверить закачку 3-х файлов для fox622.exe
            check_download_file(file_fox622, url_fox622);
            check_download_file(file_VFP6R, url_VFP6R);
            check_download_file(file_VFP6RENU, url_VFP6RENU);

            txtLogFileName = txtLogFileNameParam;
            my_filename = get_exec_filename(); // имя запущенной программы
            if(my_filename.EndsWith("\" ")){
                my_filename = my_filename.Substring(0, my_filename.Length-2);
            }
            up_filename = "new." + my_filename; // имя временного файла для новой версии

            //string[] keys = Environment.GetCommandLineArgs();
            string[] keys = argc;
            

            // временно
            //string[] keys2 = new string[3];

            //keys2[0] = keys[0];
            //keys2[1] = "/d";
            //keys2[2] = "\"PickupBases3.exe\"";

            //keys = keys2;
            
            // временно

            //if (keys.Length < 3)
            if (keys.Length < 2)
                do_check_update();
            else
            {
                WriteTofile("Argc1 = " + argc[0] + "; Argc2 = " + argc[1], txtLogFileName);

                string txtFileName = keys[1];

                if (txtFileName.StartsWith("\""))
                {
                    if (txtFileName.Length > 1) txtFileName = txtFileName.Substring(1);
                }

                if (txtFileName.EndsWith("\""))
                {
                    txtFileName = txtFileName.Substring(0, txtFileName.Length - 1);
                }
                keys[1] = txtFileName;
                
                if (keys[0].ToLower() == "/u")
                {
                    do_copy_downloaded_program(keys[1]);
                }

                if (keys[0].ToLower() == "/d")
                {
                    do_delete_old_program(keys[1]);
                }
            }
        }

        public string GetVersion() {
            return my_version;
        }

        public void do_check_update()
        {
            string new_version = get_server_version(); // Получаем номер последней версии

            if (my_version == new_version) // Если обновление не нужно
            {
                is_download = false;
                is_skipped = true; // Пропускаем модуль обновления
            }
            else
                do_download_update(); // Запускаем скачивание новой версии
        }

        public void do_download_update()
        {
            // InitializeComponent();
            // label_status.Text = "Скачивается файл: " + url_program;

            Console.WriteLine("Скачивается файл: " + url_program);
            download_file();
            is_download = true; // Будем ждать завершения процесса
            is_skipped = false; // Основную форму не нужно запускать
        }

        private void download_file()
        {
            try
            {
                WebClient webClient = new WebClient();
                webClient.DownloadFileCompleted += new AsyncCompletedEventHandler(Completed);
                webClient.DownloadProgressChanged += new DownloadProgressChangedEventHandler(ProgressChanged);
                webClient.DownloadFileAsync(new Uri(url_program), up_filename);

                // тут и надо ждать 1 минуту
                // почему тут?
                // скачивание происходит когда и как узнать что оно закончилось?
                // подождать 60000 милисек = 60 сек пока программа будет качаться
                int iLoop = 600;
                while (--iLoop > 0 && !this.skipped())
                {
                    Thread.Sleep(100);
                }

            }
            catch (Exception ex)
            {
                error(ex.Message + " функция download_file() " + up_filename);
            }
        }


        private void Completed(object sender, AsyncCompletedEventArgs e)
        {
            // Запускаем второй этап обновления
            //run_program(up_filename, "/u \"" + my_filename + "\"");
            run_program(up_filename, "/u " + my_filename);
            //this.Close();
            Environment.Exit(0);
        }


        private void ProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            // формат F2
            Console.Write("\b\b\b" + Convert.ToString(e.ProgressPercentage) + "%");
        }

        public void run_program(string filename, string keys)
        {
            try
            {
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                //proc.StartInfo.WorkingDirectory = Application.StartupPath;
                proc.StartInfo.WorkingDirectory = Environment.CurrentDirectory;
                proc.StartInfo.FileName = filename;
                proc.StartInfo.Arguments = keys; // Аргументы командной строки
                proc.Start(); // Запускаем!
            }
            catch (Exception ex)
            {
                error(ex.Message + " функция run_program() " + filename);
            }
        }

        private void check_download_file(string txtFilename, string txtUrl)
        {
            try
            {
                if (!File.Exists(txtFilename))
                {
                    WebClient webClient = new WebClient();
                    // webClient.DownloadFileCompleted += new AsyncCompletedEventHandler(Completed);
                    //webClient.DownloadProgressChanged += new DownloadProgressChangedEventHandler(ProgressChanged);
                    webClient.DownloadFile(new Uri(txtUrl), txtFilename);

                    // тут и надо ждать 1 минуту
                    // почему тут?
                    // скачивание происходит когда и как узнать что оно закончилось?
                    // подождать 60000 милисек = 60 сек пока программа будет качаться
                    int iLoop = 600;
                    while (--iLoop > 0 && !this.skipped())
                    {
                        Thread.Sleep(100);
                    }
                }

            }
            catch (Exception ex)
            {
                error(ex.Message + " функция download_file() " + up_filename);
            }
        }


        public void do_copy_downloaded_program(string filename)
        {
            try_to_delete_file(filename);
            try
            {
                WriteTofile("Копирование файла: file1 = " + my_filename + "; file2 = " + filename, txtLogFileName);

                File.Copy(my_filename, filename);
                // Запускаем последний этап обновления
                //run_program(filename, "/d \"" + my_filename + "\"");
                
                WriteTofile("Запуск процесса " + filename + " с ключом /d " + my_filename, txtLogFileName);
                run_program(filename, "/d " + my_filename);
                is_download = false;
                is_skipped = false;
            }
            catch (Exception ex)
            {
                error(ex.Message + " функция do_copy_downloaded_program() " + filename);
            }
        }

        public void do_delete_old_program(string filename)
        {
            Console.WriteLine("Удаление файла: file = " + filename);
            try_to_delete_file(filename);
            is_download = false;
            is_skipped = true;
        }

        private void try_to_delete_file(string filename)
        {
            int loop = 10;
            while (--loop > 0 && File.Exists(filename))
                try
                {
                    File.Delete(filename);
                }
                catch
                {
                    Thread.Sleep(200);
                }
        }

        private string get_server_version()
        {
            try
            {
                WebClient webClient = new WebClient();
                return webClient.DownloadString(url_version).Trim();
            }
            catch
            {
                // Если номер версии не можем получить, 
                // то программу даже и не надо пытаться.
                return my_version;
            }
        }

        private void error(string message)
        {
            /*
             if (DialogResult.Yes == MessageBox.Show(
                "Ошибка обновления: " + message +
                    "\n\nСкачать программу вручную?",
                    "Ошибка", MessageBoxButtons.YesNo))
                OpenLink(url_foruser);
             */
            Console.WriteLine("Ошибка обновления: " + message + "\n\nНеобходимо  установить обновление вручную. Оформите заявку в ОИиОИБ http://help");
            WriteTofile("Ошибка обновления: " + message + "\n", txtLogFileName);
            OpenLink("http://help");
            is_download = false;
            is_skipped = false; // в случае ошибки ничего не запускаем
        }

        private void OpenLink(string sUrl)
        {
            try
            {
                System.Diagnostics.Process.Start(sUrl);
            }
            catch (Exception exc1)
            {
                if (exc1.GetType().ToString() != "System.ComponentModel.Win32Exception")
                    try
                    {
                        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo("IExplore.exe", sUrl);
                        System.Diagnostics.Process.Start(startInfo);
                        startInfo = null;
                    }
                    catch (Exception)
                    {
                        //MessageBox.Show("Запустить обозреватель, к сожалению, не удалось.\n\nОткройте страницу ручным способом:\n" + sUrl, "ОШИБКА");
                        Console.WriteLine("Запустить обозреватель, к сожалению, не удалось.\n\nОткройте страницу ручным способом:\n" + sUrl);
                    }
            }
        }

        private string get_exec_filename()
        {
            //string fullname = Application.ExecutablePath;
            // каталог - string fullname = Environment.CurrentDirectory;
            string fullname = Environment.CommandLine;
            string txtRes = "";

            string[] split = { "\\" };
            string[] parts = fullname.Split(split, StringSplitOptions.None);
            if (parts.Length > 0)
                txtRes =  parts[parts.Length - 1];

            // если есть параметры комндной строки
            if(txtRes.Contains(".exe\"")){
                txtRes = txtRes.Substring(0, txtRes.IndexOf(".exe\"")+4);
            }

            return txtRes;
            
        }

        private static bool WriteTofile(string txtText, string outfile)
        {
            using (StreamWriter sw = new StreamWriter(outfile, true))
            {
                sw.WriteLine(txtText);
                sw.Close();
            }
            return true;

        }
    }

}
