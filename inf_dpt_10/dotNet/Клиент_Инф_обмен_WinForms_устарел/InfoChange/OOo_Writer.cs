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

using Microsoft.Win32;

using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.bridge;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.text;
using unoidl.com.sun.star.table;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.style;

namespace InfoChange
{
    public class OOo_Writer
    {
        private OleDbConnection FBcon;
        private OleDbConnection DBFcon;
        private OleDbCommand m_cmd;

        XComponentContext m_xContext;
        XMultiServiceFactory mxMSFactory;

        #region "OOo"
        private void InitOO3Env()
        {
            string baseKey;
            //Для 64 битной версии
            if (Marshal.SizeOf(typeof(IntPtr)) == 8) baseKey = @"SOFTWARE\Wow6432Node\OpenOffice.org\";
            else
                baseKey = @"SOFTWARE\OpenOffice.org\";

            // Get the URE directory
            string key = baseKey + @"Layers\URE\1";
            RegistryKey reg = Registry.CurrentUser.OpenSubKey(key);
            if (reg == null) reg = Registry.LocalMachine.OpenSubKey(key);
            string urePath = reg.GetValue("UREINSTALLLOCATION") as string;
            reg.Close();
            urePath = Path.Combine(urePath, "bin");

            // Get the UNO Path
            key = baseKey + @"UNO\InstallPath";
            reg = Registry.CurrentUser.OpenSubKey(key);
            if (reg == null) reg = Registry.LocalMachine.OpenSubKey(key);
            string unoPath = reg.GetValue(null) as string;
            reg.Close();

            string path;

            string sysPath = System.Environment.GetEnvironmentVariable("PATH");
            path = string.Format("{0};{1}", System.Environment.GetEnvironmentVariable("PATH"), urePath);
            System.Environment.SetEnvironmentVariable("PATH", path);
            System.Environment.SetEnvironmentVariable("UNO_PATH", unoPath);
        }
        //Проверяем установлен ли?
        public bool isOOoInstalled()
        {
            try
            {
                string baseKey;
                //if ()

                baseKey = @"SOFTWARE\OpenOffice.org\";

                // Get the URE directory
                string key = baseKey + @"Layers\URE\1";
                RegistryKey reg = Registry.CurrentUser.OpenSubKey(key);
                if (reg == null) reg = Registry.LocalMachine.OpenSubKey(key);
                string urePath = reg.GetValue("UREINSTALLLOCATION") as string;
                reg.Close();
                if (urePath != null) return true;
                else
                    return false;
            }
            catch 
            {
                return false;
            }
        }

        //Соединияемся с фабрикой (чтобы создавать \ открывать документы)
        private unoidl.com.sun.star.lang.XMultiServiceFactory uno_connect(String[] args)
        {
            InitOO3Env();
            m_xContext = uno.util.Bootstrap.bootstrap();
            

            if (m_xContext != null)
                return (unoidl.com.sun.star.lang.XMultiServiceFactory)m_xContext.getServiceManager();
            else
                return null;
        }

        //Создаем \ загружаем документ Calc
        private unoidl.com.sun.star.sheet.XSpreadsheetDocument OOo3_initCalcDocument(string filePath, bool newDoc)
        {

            XComponentLoader aLoader;
            XComponent xComponent = null;
            string url = newDoc ? "private:factory/scalc" : @"file:///" + filePath.Replace(@"\", @"/");
            try
            {
                aLoader = (XComponentLoader)
                mxMSFactory.createInstance("com.sun.star.frame.Desktop");

                xComponent = aLoader.loadComponentFromURL(
                    /*"private:factory/scalc"*/ url, "_blank", 0,
                new unoidl.com.sun.star.beans.PropertyValue[0]);
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                //iLog.WriteLog("OOo3 Exception in OOo3_initDocument(string filePath, bool newDoc):-> " + ex.Message + url);
            }
            return (unoidl.com.sun.star.sheet.XSpreadsheetDocument)xComponent;

        }

        //Создаем \ загружаем документ Writer
        private unoidl.com.sun.star.text.XTextDocument OOo3_initWriterDocument(string filePath, bool newDoc)
        {

            XComponentLoader aLoader;
            XComponent xComponent = null;
            string url = newDoc ? "private:factory/swriter" : @"file:///" + filePath.Replace(@"\", @"/");
            try
            {
                aLoader = (XComponentLoader)
                mxMSFactory.createInstance("com.sun.star.frame.Desktop");

                xComponent = aLoader.loadComponentFromURL(
                url, "_blank", 0,
                new unoidl.com.sun.star.beans.PropertyValue[0]);
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                //iLog.WriteLog("OOo3 Exception in OOo3_initDocument(string filePath, bool newDoc):-> " + ex.Message + url);
            }
            return (unoidl.com.sun.star.text.XTextDocument)xComponent;
        }

        private static string PathConverter(string file)
        {
            try
            {
                file = file.Replace(@"\", "/");

                return "file:///" + file;
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }

        # endregion

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

        public void OOo_Sber(string org, string type, DataSet ds, OleDbConnection con, Form1 mainF)
        {
            FBcon = con;

            if (type == "nofind")
            {
            #region "NOFIND"
                try
                {
                    XComponent doc;
                    XText myText;
                    XTextCursor myTextCursor;
                    XPropertySet myCursorProps;
                    XParagraphCursor myParagCursor;
                    XPropertySet myParagProps;
                    XTextDocument myTextDocument;

                    uno.Any myEnum;

                    string[] par = new string[1];
                    par[0] = "";
                    if (isOOoInstalled())
                    {
                        mxMSFactory = uno_connect(par);

                        doc = OOo3_initWriterDocument("", true);

                        myTextDocument = (XTextDocument)doc;
                        myText = myTextDocument.getText();

                        // create a paragraph cursor  
                        XParagraphCursor xParagraphCursor = (XParagraphCursor)(myText.createTextCursor());
                        XPropertySet myPropertySet = (XPropertySet)xParagraphCursor;


                        //*********
                        XMultiServiceFactory mxDocFactory = (XMultiServiceFactory)myTextDocument;
                        XText xText = myTextDocument.getText();
                        XTextCursor xTextCursor = xText.createTextCursor();
                        XPropertySet xTextCursorProps = (XPropertySet)xTextCursor;

                        String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").Value.ToString();
                        XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier)myTextDocument;
                        XNameAccess xFamilies = (XNameAccess)xSupplier.getStyleFamilies();
                        Object Families = xFamilies.getByName("PageStyles").Value;
                        XNameContainer xFamily = (XNameContainer)Families;

                        Object Family = xFamily.getByName(pageStyleName).Value;
                        XStyle xStyle = (XStyle)Family;
                        // Get the property set of the TextCursor 
                        XPropertySet xStyleProps = (XPropertySet)xStyle;

                        xStyleProps.setPropertyValue("LeftMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("RightMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("TopMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("BottomMargin", new uno.Any(1000));

                        //***************
                        //Шрифт и размер
                        myPropertySet.setPropertyValue("CharFontName", new uno.Any("Courier"));
                        myPropertySet.setPropertyValue("CharHeight", new uno.Any(8));

                        int spi;
                        int sch_line;
                        int fl_fst = 1;

                        foreach (DataRow drspi in ds.Tables["SPI"].Rows)
                        {
                            sch_line = 0;
                            if (fl_fst == 1)
                            {
                                sch_line = 1;
                                fl_fst = 0;
                                //par = doc.Paragraphs[1];
                            }
                            else
                            {
                                myPropertySet.setPropertyValue("BreakType", new uno.Any(typeof(BreakType), BreakType.PAGE_BEFORE));
                            }

                            myText.insertString(myText.getEnd(), "Реестр ответов на запросы СП-И о наличии ден. средств из банка\r\r", false);
                            myText.insertString(myText.getEnd(), "" + org + " от " + Convert.ToDateTime(ds.Tables["NOFIND"].Rows[0]["DATOTV"]).ToShortDateString() + "\r\r", false);
                            myText.insertString(myText.getEnd(), "За период с " + Convert.ToDateTime(ds.Tables["NOFIND"].Rows[0]["DATZPR1"]).ToShortDateString() + " по " + Convert.ToDateTime(ds.Tables["NOFIND"].Rows[0]["DATZPR2"]).ToShortDateString() + "\r\r", false);                      

                            spi = Convert.ToInt32(drspi["NOMSPI"]);

                            myText.insertString(myText.getEnd(), "СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "\r\r", false);

                            myText.insertString(myText.getEnd(), "Нет данных о наличии счетов у должников\r\r", false);

                            sch_line += 10;

                            foreach (DataRow row in ds.Tables["NOFIND"].Rows)
                            {
                                if (spi == Convert.ToInt32(row["NOMSPI"]))
                                {
                                    if (Convert.ToInt32(row["GODR"])==0)
                                        myText.insertString(myText.getEnd(), "" + Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FIO"]).TrimEnd() + "\r\r", false);
                                    else
                                        myText.insertString(myText.getEnd(), "" + Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FIO"]).TrimEnd() + " (" + Convert.ToInt32(row["GODR"]).ToString() + " г.р.) " + "\r\r", false);
                                    sch_line += 2;
                                }
                            }
                        }

                        //Сохранение файла
                        //((XStorable)doc).storeToURL(PathConverter("c:\\Temp\\1.odt"), new unoidl.com.sun.star.beans.PropertyValue[0]);

                        //Закрытие файла
                        //((unoidl.com.sun.star.text.XTextDocument)doc).dispose();
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
                # endregion
            }
            #region "FIND"

            else
            {
                try
                {
                    //**********Формирование**реестра**find************
                    XComponent doc;
                    XText myText;
                    XTextCursor myTextCursor;
                    XPropertySet myCursorProps;
                    XParagraphCursor myParagCursor;
                    XPropertySet myParagProps;
                    XTextDocument myTextDocument;

                    uno.Any myEnum;

                    string[] par = new string[1];
                    par[0] = "";
                    if (isOOoInstalled())
                    {
                        mxMSFactory = uno_connect(par);

                        doc = OOo3_initWriterDocument("", true);

                        myTextDocument = (XTextDocument)doc;
                        myText = myTextDocument.getText();

                        // create a paragraph cursor  
                        XParagraphCursor xParagraphCursor = (XParagraphCursor)(myText.createTextCursor());
                        XPropertySet myPropertySet = (XPropertySet)xParagraphCursor;


                        //*********
                        XMultiServiceFactory mxDocFactory = (XMultiServiceFactory)myTextDocument;
                        XText xText = myTextDocument.getText();
                        XTextCursor xTextCursor = xText.createTextCursor();
                        XPropertySet xTextCursorProps = (XPropertySet)xTextCursor;

                        String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").Value.ToString();
                        XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier)myTextDocument;
                        XNameAccess xFamilies = (XNameAccess)xSupplier.getStyleFamilies();
                        Object Families = xFamilies.getByName("PageStyles").Value;
                        XNameContainer xFamily = (XNameContainer)Families;

                        Object Family = xFamily.getByName(pageStyleName).Value;
                        XStyle xStyle = (XStyle)Family;
                        // Get the property set of the TextCursor 
                        XPropertySet xStyleProps = (XPropertySet)xStyle;

                        xStyleProps.setPropertyValue("LeftMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("RightMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("TopMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("BottomMargin", new uno.Any(1000));

                        //***************
                        //Шрифт и размер
                        myPropertySet.setPropertyValue("CharFontName", new uno.Any("Courier"));
                        myPropertySet.setPropertyValue("CharHeight", new uno.Any(8));

                        int spi;
                        int sch_line;
                        int fl_fst = 1;

                        foreach (DataRow drspi in ds.Tables["SPI"].Rows)
                        {
                            sch_line = 0;
                            if (fl_fst == 1)
                            {
                                sch_line = 1;
                                fl_fst = 0;
                            }
                            else
                            {
                                myPropertySet.setPropertyValue("BreakType", new uno.Any(typeof(BreakType), BreakType.PAGE_BEFORE));
                            }

                            myText.insertString(myText.getEnd(), "Реестр ответов на запросы СП-И о наличии ден. средств из банка\r\r", false);
                            myText.insertString(myText.getEnd(), "" + org + " от " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATOTV"]).ToShortDateString() + "\r\r", false);
                            myText.insertString(myText.getEnd(), "За период с " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATZPR1"]).ToShortDateString() + " по " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATZPR2"]).ToShortDateString() + "\r\r", false);

                            spi = Convert.ToInt32(drspi["NOMSPI"]);

                            myText.insertString(myText.getEnd(), "СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "\r\r", false);
                            myText.insertString(myText.getEnd(), "НОМЕР ИП             ДОЛЖНИК                            АДРЕС                      СЧЕТ       ОСТАТОК\r\r", false);

                            sch_line += 10;

                            foreach (DataRow row in ds.Tables["FIND"].Rows)
                            {
                                if (spi == Convert.ToInt32(row["NOMSPI"]))
                                {
                                    string txtLs = Convert.ToString(row["NOMLS"]).TrimEnd();
                                    string txtResponse = Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FIO"]).TrimEnd() + " (" + Convert.ToInt32(row["GODR"]).ToString() + " г.р.) " + Convert.ToString(row["ADRES"]).TrimEnd() + " " + txtLs + " остаток = " + Convert.ToString(row["OSTAT"]).Trim() + " " + getValuteByCod(txtLs) + " " + Convert.ToString(row["PRIZ"]).TrimEnd();
                                    if (txtResponse.Length > 112) sch_line += 3;
                                    else sch_line += 2;

                                    myText.insertString(myText.getEnd(), txtResponse + "\r\r", false);
                                }
                            }

                            //foreach (DataRow row in tbl.Rows)
                            //{
                            //    if (spi == Convert.ToInt32(row["NOMSPI"]))
                            //    {
                            //        string txtResponse = Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FIO"]).TrimEnd() + " " + Convert.ToString(row["ADRES"]).TrimEnd() + " " + Convert.ToString(row["NOMLS"]).TrimEnd() + " остаток = " + Money_ToStr(Convert.ToDecimal(row["OSTAT"])).TrimEnd() + " " + Convert.ToString(row["PRIZ"]).TrimEnd();
                            //        if (txtResponse.Length > 112) sch_line += 3;
                            //        else sch_line += 2;
                            //        par.Range.Text += txtResponse + "\n";
                            //    }
                            //}
                        }

                        //Сохранение файла
                        //((XStorable)doc).storeToURL(PathConverter("c:\\Temp\\1.odt"), new unoidl.com.sun.star.beans.PropertyValue[0]);

                        //Закрытие файла
                        //((unoidl.com.sun.star.text.XTextDocument)doc).dispose();
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
            # endregion
            }
        }

        public bool ReestrOutOOo(DataTable dtReg, string dir, string ospadr, string ospnam)
        {
            Decimal nYear = DateTime.Today.Year;
            DateTime dtDate;
            string bankname = "Карельское ОСБ N 8628 АК СБ РФ";
            //string bankadres = "г.Петрозаводск, ул.Антикайнена, д.2";
            string bankadres = "";
            string ospadres = ospadr;
            string ospname = ospnam;

            DataRow[] FizRows = dtReg.Select("LITZDOLG LIKE '/1/*'", "FIOVK");

            //if (File.Exists(string.Format(@"{0}\{1}", dir, DateTime.Today.ToShortDateString() + ".odt")))
            //    File.Delete(string.Format(@"{0}\{1}", dir, DateTime.Today.ToShortDateString() + ".odt"));
 
            try
            {
                XComponent doc;
                XText myText;
                XTextCursor myTextCursor;
                XPropertySet myCursorProps;
                XParagraphCursor myParagCursor;
                XPropertySet myParagProps;
                XTextDocument myTextDocument;

                uno.Any myEnum;

                string[] par = new string[1];
                par[0] = "";
                if (isOOoInstalled())
                {
                    mxMSFactory = uno_connect(par);

                    doc = OOo3_initWriterDocument("", true);

                    myTextDocument = (XTextDocument)doc;
                    myText = myTextDocument.getText();

                    // create a paragraph cursor  
                    XParagraphCursor xParagraphCursor = (XParagraphCursor)(myText.createTextCursor());
                    XPropertySet myPropertySet = (XPropertySet)xParagraphCursor;


                    //*********
                    XMultiServiceFactory mxDocFactory = (XMultiServiceFactory)myTextDocument;
                    XText xText = myTextDocument.getText();
                    XTextCursor xTextCursor = xText.createTextCursor();
                    XPropertySet xTextCursorProps = (XPropertySet)xTextCursor;

                    String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").Value.ToString();
                    XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier)myTextDocument;
                    XNameAccess xFamilies = (XNameAccess)xSupplier.getStyleFamilies();
                    Object Families = xFamilies.getByName("PageStyles").Value;
                    XNameContainer xFamily = (XNameContainer)Families;

                    Object Family = xFamily.getByName(pageStyleName).Value;
                    XStyle xStyle = (XStyle)Family;
                    // Get the property set of the TextCursor 
                    XPropertySet xStyleProps = (XPropertySet)xStyle;

                    xStyleProps.setPropertyValue("LeftMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("RightMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("TopMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("BottomMargin", new uno.Any(1000));

                    //***************
                    //Шрифт и размер
                    myPropertySet.setPropertyValue("CharFontName", new uno.Any("Courier"));
                    myPropertySet.setPropertyValue("CharHeight", new uno.Any(8));

                    int spi;
                    int sch_line;
                    int fl_fst = 1;

                    myText.insertString(myText.getEnd(), "              МИНЮСТ                      " + bankname + "\r", false);  
                    myText.insertString(myText.getEnd(), "   ФЕДЕРАЛЬНАЯ СЛУЖБА СУДЕБНЫХ ПРИСТАВОВ  " + bankadres + "\r", false);
                    myText.insertString(myText.getEnd(), "   УПРАВЛЕНИЕ ФССП ПО РЕСПУБЛИКЕ КАРЕЛИЯ  \r\r", false);
                    myText.insertString(myText.getEnd(), "  " + ospname + "\r\r", false);
                    myText.insertString(myText.getEnd(), "  " + ospadres + "\r\r", false);
                    myText.insertString(myText.getEnd(), "    Исх.N ________от _________   ;\r\r", false);
                    myText.insertString(myText.getEnd(), "                                З А П Р О С\r", false);
                    myText.insertString(myText.getEnd(), "   На исполнении в " + ospname + "\r", false);
                    myText.insertString(myText.getEnd(), "   находятся исполнительные документы на должников : \r", false);

                    foreach (DataRow row in FizRows)
                    {
                        nYear = 0;
                        if (DateTime.TryParse(Convert.ToString(row["GOD"]), out dtDate))
                        {
                            nYear = dtDate.Year;
                        }
                        myText.insertString(myText.getEnd(), "   " + Convert.ToString(row["ZAPROS"]) + ", " + Convert.ToString(row["FIOVK"]) + ", " + Convert.ToString(row["ADDR"]) + ", " + Convert.ToString(nYear) + "\r", false);
                    }
                    myText.insertString(myText.getEnd(), "\r\r", false);
                    myText.insertString(myText.getEnd(), "  Просим Вас в семидневный срок сообщить о счетах должников,\r", false);
                    myText.insertString(myText.getEnd(), "  зарегистрированных в Вашем банке.\r\r\r", false);
                    myText.insertString(myText.getEnd(), "  Исполнитель:  \r", false);
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
            
            return true;
        }


        public void OOo_Cred(string org, string type, DataSet ds, OleDbConnection con, Form1 mainF)
        {
            FBcon = con;

            if (type == "nofind")
            {
            #region "NOFIND"
                try
                {
                    XComponent doc;
                    XText myText;
                    XTextCursor myTextCursor;
                    XPropertySet myCursorProps;
                    XParagraphCursor myParagCursor;
                    XPropertySet myParagProps;
                    XTextDocument myTextDocument;

                    uno.Any myEnum;

                    string[] par = new string[1];
                    par[0] = "";
                    if (isOOoInstalled())
                    {
                        mxMSFactory = uno_connect(par);

                        doc = OOo3_initWriterDocument("", true);

                        myTextDocument = (XTextDocument)doc;
                        myText = myTextDocument.getText();

                        // create a paragraph cursor  
                        XParagraphCursor xParagraphCursor = (XParagraphCursor)(myText.createTextCursor());
                        XPropertySet myPropertySet = (XPropertySet)xParagraphCursor;


                        //*********
                        XMultiServiceFactory mxDocFactory = (XMultiServiceFactory)myTextDocument;
                        XText xText = myTextDocument.getText();
                        XTextCursor xTextCursor = xText.createTextCursor();
                        XPropertySet xTextCursorProps = (XPropertySet)xTextCursor;

                        String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").Value.ToString();
                        XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier)myTextDocument;
                        XNameAccess xFamilies = (XNameAccess)xSupplier.getStyleFamilies();
                        Object Families = xFamilies.getByName("PageStyles").Value;
                        XNameContainer xFamily = (XNameContainer)Families;

                        Object Family = xFamily.getByName(pageStyleName).Value;
                        XStyle xStyle = (XStyle)Family;
                        // Get the property set of the TextCursor 
                        XPropertySet xStyleProps = (XPropertySet)xStyle;

                        xStyleProps.setPropertyValue("LeftMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("RightMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("TopMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("BottomMargin", new uno.Any(1000));

                        //***************
                        //Шрифт и размер
                        myPropertySet.setPropertyValue("CharFontName", new uno.Any("Courier"));
                        myPropertySet.setPropertyValue("CharHeight", new uno.Any(8));

                        int spi;
                        int sch_line;
                        int fl_fst = 1;

                        foreach (DataRow drspi in ds.Tables["SPI"].Rows)
                        {
                            sch_line = 0;
                            if (fl_fst == 1)
                            {
                                sch_line = 1;
                                fl_fst = 0;
                                //par = doc.Paragraphs[1];
                            }
                            else
                            {
                                myPropertySet.setPropertyValue("BreakType", new uno.Any(typeof(BreakType), BreakType.PAGE_BEFORE));
                            }

                            myText.insertString(myText.getEnd(), "Реестр ответов на запросы СП-И о наличии ден. средств из банка\r\r", false);
                            myText.insertString(myText.getEnd(), "" + org + " от " + Convert.ToDateTime(ds.Tables["NOFIND"].Rows[0]["DATOTV"]).ToShortDateString() + "\r\r", false);
                            myText.insertString(myText.getEnd(), "За период с " + Convert.ToDateTime(ds.Tables["NOFIND"].Rows[0]["DATZPR1"]).ToShortDateString() + " по " + Convert.ToDateTime(ds.Tables["NOFIND"].Rows[0]["DATZPR2"]).ToShortDateString() + "\r\r", false);
                            myText.insertString(myText.getEnd(), "Нет данных о наличии счетов у должников\r\r", false);

                            spi = Convert.ToInt32(drspi["NOMSPI"]);

                            myText.insertString(myText.getEnd(), "СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "\r\r", false);

                            sch_line += 10;

                            foreach (DataRow row in ds.Tables["NOFIND"].Rows)
                            {
                                if (spi == Convert.ToInt32(row["NOMSPI"]))
                                {
                                    myText.insertString(myText.getEnd(), "" + GetIPNum(con, Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIO"]).TrimEnd() + "\r\r", false);
                                    sch_line += 2;
                                }
                            }
                        }

                        //Сохранение файла
                        //((XStorable)doc).storeToURL(PathConverter("c:\\Temp\\1.odt"), new unoidl.com.sun.star.beans.PropertyValue[0]);

                        //Закрытие файла
                        //((unoidl.com.sun.star.text.XTextDocument)doc).dispose();
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
                # endregion
            }
            #region "FIND"

            else if(type == "find")
            {
                try
                {
                    //**********Формирование**реестра**find************
                    XComponent doc;
                    XText myText;
                    XTextCursor myTextCursor;
                    XPropertySet myCursorProps;
                    XParagraphCursor myParagCursor;
                    XPropertySet myParagProps;
                    XTextDocument myTextDocument;

                    uno.Any myEnum;

                    string[] par = new string[1];
                    par[0] = "";
                    if (isOOoInstalled())
                    {
                        mxMSFactory = uno_connect(par);

                        doc = OOo3_initWriterDocument("", true);

                        myTextDocument = (XTextDocument)doc;
                        myText = myTextDocument.getText();

                        // create a paragraph cursor  
                        XParagraphCursor xParagraphCursor = (XParagraphCursor)(myText.createTextCursor());
                        XPropertySet myPropertySet = (XPropertySet)xParagraphCursor;


                        //*********
                        XMultiServiceFactory mxDocFactory = (XMultiServiceFactory)myTextDocument;
                        XText xText = myTextDocument.getText();
                        XTextCursor xTextCursor = xText.createTextCursor();
                        XPropertySet xTextCursorProps = (XPropertySet)xTextCursor;

                        String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").Value.ToString();
                        XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier)myTextDocument;
                        XNameAccess xFamilies = (XNameAccess)xSupplier.getStyleFamilies();
                        Object Families = xFamilies.getByName("PageStyles").Value;
                        XNameContainer xFamily = (XNameContainer)Families;

                        Object Family = xFamily.getByName(pageStyleName).Value;
                        XStyle xStyle = (XStyle)Family;
                        // Get the property set of the TextCursor 
                        XPropertySet xStyleProps = (XPropertySet)xStyle;

                        xStyleProps.setPropertyValue("LeftMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("RightMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("TopMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("BottomMargin", new uno.Any(1000));

                        //***************
                        //Шрифт и размер
                        myPropertySet.setPropertyValue("CharFontName", new uno.Any("Courier"));
                        myPropertySet.setPropertyValue("CharHeight", new uno.Any(8));

                        int spi;
                        int sch_line;
                        int fl_fst = 1;
                        int year = 0;

                        foreach (DataRow drspi in ds.Tables["SPI"].Rows)
                        {
                            sch_line = 0;
                            if (fl_fst == 1)
                            {
                                sch_line = 1;
                                fl_fst = 0;
                            }
                            else
                            {
                                myPropertySet.setPropertyValue("BreakType", new uno.Any(typeof(BreakType), BreakType.PAGE_BEFORE));
                            }

                            myText.insertString(myText.getEnd(), "Реестр ответов на запросы СП-И о наличии ден. средств из банка\r\r", false);
                            myText.insertString(myText.getEnd(), "" + org + " от " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATOTV"]).ToShortDateString() + "\r\r", false);
                            myText.insertString(myText.getEnd(), "За период с " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATZPR1"]).ToShortDateString() + " по " + Convert.ToDateTime(ds.Tables["FIND"].Rows[0]["DATZPR2"]).ToShortDateString() + "\r\r", false);

                            spi = Convert.ToInt32(drspi["NOMSPI"]);

                            myText.insertString(myText.getEnd(), "СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "\r\r", false);
                            myText.insertString(myText.getEnd(), "НОМЕР ИП             ДОЛЖНИК                            АДРЕС                      СЧЕТ       ОСТАТОК\r\r", false);

                            sch_line += 10;

                            foreach (DataRow row in ds.Tables["FIND"].Rows)
                            {
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


                                    myText.insertString(myText.getEnd(), txtResLine + "\r\r", false);
                                }
                            }
                        }

                        //Сохранение файла
                        //((XStorable)doc).storeToURL(PathConverter("c:\\Temp\\1.odt"), new unoidl.com.sun.star.beans.PropertyValue[0]);

                        //Закрытие файла
                        //((unoidl.com.sun.star.text.XTextDocument)doc).dispose();
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
            # endregion
            }
            #region "E_TOFIND"
            else if(type == "e_tofind")
            {
                try
                {
                    //**********Формирование**реестра**e_tofind************
                    XComponent doc;
                    XText myText;
                    XTextCursor myTextCursor;
                    XPropertySet myCursorProps;
                    XParagraphCursor myParagCursor;
                    XPropertySet myParagProps;
                    XTextDocument myTextDocument;

                    uno.Any myEnum;

                    string[] par = new string[1];
                    par[0] = "";
                    if (isOOoInstalled())
                    {
                        mxMSFactory = uno_connect(par);

                        doc = OOo3_initWriterDocument("", true);

                        myTextDocument = (XTextDocument)doc;
                        myText = myTextDocument.getText();

                        // create a paragraph cursor  
                        XParagraphCursor xParagraphCursor = (XParagraphCursor)(myText.createTextCursor());
                        XPropertySet myPropertySet = (XPropertySet)xParagraphCursor;


                        //*********
                        XMultiServiceFactory mxDocFactory = (XMultiServiceFactory)myTextDocument;
                        XText xText = myTextDocument.getText();
                        XTextCursor xTextCursor = xText.createTextCursor();
                        XPropertySet xTextCursorProps = (XPropertySet)xTextCursor;

                        String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").Value.ToString();
                        XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier)myTextDocument;
                        XNameAccess xFamilies = (XNameAccess)xSupplier.getStyleFamilies();
                        Object Families = xFamilies.getByName("PageStyles").Value;
                        XNameContainer xFamily = (XNameContainer)Families;

                        Object Family = xFamily.getByName(pageStyleName).Value;
                        XStyle xStyle = (XStyle)Family;
                        // Get the property set of the TextCursor 
                        XPropertySet xStyleProps = (XPropertySet)xStyle;

                        xStyleProps.setPropertyValue("LeftMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("RightMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("TopMargin", new uno.Any(1000));
                        xStyleProps.setPropertyValue("BottomMargin", new uno.Any(1000));

                        //***************
                        //Шрифт и размер
                        myPropertySet.setPropertyValue("CharFontName", new uno.Any("Courier"));
                        myPropertySet.setPropertyValue("CharHeight", new uno.Any(8));

                        int spi;
                        int sch_line;
                        int fl_fst = 1;
                        int year = 0;

                        foreach (DataRow drspi in ds.Tables["SPI"].Rows)
                        {
                            sch_line = 0;
                            if (fl_fst == 1)
                            {
                                sch_line = 1;
                                fl_fst = 0;
                            }
                            else
                            {
                                myPropertySet.setPropertyValue("BreakType", new uno.Any(typeof(BreakType), BreakType.PAGE_BEFORE));
                            }

                            myText.insertString(myText.getEnd(), "Реестр не принятых в обработку запросов СП-И о наличии ден. средств из банка\r\r", false);
                            myText.insertString(myText.getEnd(), "" + org + " от " + DateTime.Today.ToShortDateString() + "\r\r", false);
                            myText.insertString(myText.getEnd(), "За период с " + Convert.ToDateTime(ds.Tables["E_TOFIND"].Rows[0]["DATZAPR1"]).ToShortDateString() + " по " + Convert.ToDateTime(ds.Tables["E_TOFIND"].Rows[0]["DATZAPR1"]).ToShortDateString() + "\r\r", false);

                            spi = Convert.ToInt32(drspi["NOMSPI"]);

                            myText.insertString(myText.getEnd(), "СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "\r\r", false);
                            myText.insertString(myText.getEnd(), "НОМЕР ИП             ДОЛЖНИК\r\r", false);

                            sch_line += 10;

                            foreach (DataRow row in ds.Tables["E_TOFIND"].Rows)
                            {
                                if (spi == Convert.ToInt32(row["NOMSPI"]))
                                {
                                    string txtResLine = GetIPNum(con, Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FIOVK"]).TrimEnd();
                                    if (Convert.ToInt32(row["LITZDOLG"]) == 2)
                                    {
                                        txtResLine += " (" + Convert.ToInt32(row["GOD"]).ToString() + " г.р.)";
                                        //txtResLine += " (" + Convert.ToInt32(row["GODR"]).ToString() + " г.р.)";
                                    }
                                    myText.insertString(myText.getEnd(), txtResLine + "\r\r", false);
                                }
                            }
                        }

                        //Сохранение файла
                        //((XStorable)doc).storeToURL(PathConverter("c:\\Temp\\1.odt"), new unoidl.com.sun.star.beans.PropertyValue[0]);

                        //Закрытие файла
                        //((unoidl.com.sun.star.text.XTextDocument)doc).dispose();
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
            # endregion
            }
        }

        public void OOo_Ktfoms(string tblname, DataSet ds, OleDbConnection con, Form1 mainF)
        {
            FBcon = con;

            try
            {
                //**********Формирование**реестра**find************
                XComponent doc;
                XText myText;
                XTextCursor myTextCursor;
                XPropertySet myCursorProps;
                XParagraphCursor myParagCursor;
                XPropertySet myParagProps;
                XTextDocument myTextDocument;

                uno.Any myEnum;

                string[] par = new string[1];
                par[0] = "";
                if (isOOoInstalled())
                {
                    mxMSFactory = uno_connect(par);

                    doc = OOo3_initWriterDocument("", true);

                    myTextDocument = (XTextDocument)doc;
                    myText = myTextDocument.getText();

                    // create a paragraph cursor  
                    XParagraphCursor xParagraphCursor = (XParagraphCursor)(myText.createTextCursor());
                    XPropertySet myPropertySet = (XPropertySet)xParagraphCursor;


                    //*********
                    XMultiServiceFactory mxDocFactory = (XMultiServiceFactory)myTextDocument;
                    XText xText = myTextDocument.getText();
                    XTextCursor xTextCursor = xText.createTextCursor();
                    XPropertySet xTextCursorProps = (XPropertySet)xTextCursor;

                    String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").Value.ToString();
                    XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier)myTextDocument;
                    XNameAccess xFamilies = (XNameAccess)xSupplier.getStyleFamilies();
                    Object Families = xFamilies.getByName("PageStyles").Value;
                    XNameContainer xFamily = (XNameContainer)Families;

                    Object Family = xFamily.getByName(pageStyleName).Value;
                    XStyle xStyle = (XStyle)Family;
                    // Get the property set of the TextCursor 
                    XPropertySet xStyleProps = (XPropertySet)xStyle;

                    xStyleProps.setPropertyValue("LeftMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("RightMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("TopMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("BottomMargin", new uno.Any(1000));

                    //***************
                    //Шрифт и размер
                    myPropertySet.setPropertyValue("CharFontName", new uno.Any("Courier"));
                    myPropertySet.setPropertyValue("CharHeight", new uno.Any(8));

                    int spi;
                    int sch_line;
                    int fl_fst = 1;
                    string priz = "";
                    DataTable tbl = ds.Tables[tblname];

                    foreach (DataRow drspi in ds.Tables["SPI"].Rows)
                    {
                        sch_line = 0;
                        if (fl_fst == 1)
                        {
                            sch_line = 1;
                            fl_fst = 0;
                        }
                        else
                        {
                            myPropertySet.setPropertyValue("BreakType", new uno.Any(typeof(BreakType), BreakType.PAGE_BEFORE));
                        }

                        myText.insertString(myText.getEnd(), "Реестр ответов на запросы СП-И в КТФОМС\r", false);
                        myText.insertString(myText.getEnd(), "Ответ из КТФОМС от " + Convert.ToDateTime(tbl.Rows[0]["DATZAPR"]).ToShortDateString() + "\r\r", false);
                       
                        spi = Convert.ToInt32(drspi["NOMSPI"]);

                        myText.insertString(myText.getEnd(), "СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "\r\r", false);
                        sch_line += 9;

                        foreach (DataRow row in tbl.Rows)
                        {
                            priz = Convert.ToString(row["PRIZ"]).Trim();
                            if (priz == "T")
                            {
                                if (spi == Convert.ToInt32(row["NOMSPI"]))
                                {
                                    //myText.insertString(myText.getEnd(), "" + Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FIO"]).TrimEnd() + "\r", false);
                                    //sch_line += 2;

                                    string txtResponse = "";
                                    if (row.Table.Columns.Contains("NAME"))
                                    {
                                        txtResponse = GetIPNum(con, Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["NAME"]).TrimEnd() + " " + Convert.ToString(row["FNAME"]).TrimEnd() + " " + Convert.ToString(row["SNAME"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd();
                                    }
                                    else
                                    {
                                        txtResponse = GetIPNum(con, Convert.ToString(row["ZAPROS"]).TrimEnd()) + " " + Convert.ToString(row["FAM"]).TrimEnd() + " " + Convert.ToString(row["IM"]).TrimEnd() + " " + Convert.ToString(row["OT"]).TrimEnd() + " " + Convert.ToDateTime(row["DD_R"]).ToShortDateString().TrimEnd();
                                    }
                                    myText.insertString(myText.getEnd(), txtResponse + "\r", false);

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

                                    if (txtResponse.Length > 112)
                                        sch_line += 3;
                                    else sch_line += 2;

                                    myText.insertString(myText.getEnd(), txtResponse + "\r\r", false);
                                }
                            }
                        }
                    }

                    //Сохранение файла
                    //((XStorable)doc).storeToURL(PathConverter("c:\\Temp\\1.odt"), new unoidl.com.sun.star.beans.PropertyValue[0]);

                    //Закрытие файла
                    //((unoidl.com.sun.star.text.XTextDocument)doc).dispose();
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
        }

        public void OOo_Pens(string tblname, DataSet ds, OleDbConnection con, Form1 mainF)
        {
            FBcon = con;
            try
            {
                string str;
                str = "";
                

                ReportMaker Rep = new ReportMaker();
                Rep.StartReport();
                Rep.AddToReport("<h2>Привет, я отчет-тест!</h2><br />Съешь еще этих мягких булочек!<br />Съешь еще этих мягких булочек!<br />Съешь еще этих мягких булочек!<br />Съешь еще этих мягких булочек!<br />Съешь еще этих мягких булочек!");
                Rep.EndReport();
                Rep.ShowReport();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", System.Windows.Forms.MessageBoxButtons.OK);
            }
            /*
            try
            {
                //**********Формирование**реестра**find************
                XComponent doc;
                XText myText;
                XTextCursor myTextCursor;
                XPropertySet myCursorProps;
                XParagraphCursor myParagCursor;
                XPropertySet myParagProps;
                XTextDocument myTextDocument;

                uno.Any myEnum;

                string[] par = new string[1];
                par[0] = "";
                if (isOOoInstalled())
                {
                    mxMSFactory = uno_connect(par);

                    doc = OOo3_initWriterDocument("", true);

                    myTextDocument = (XTextDocument)doc;
                    myText = myTextDocument.getText();

                    // create a paragraph cursor  
                    XParagraphCursor xParagraphCursor = (XParagraphCursor)(myText.createTextCursor());
                    XPropertySet myPropertySet = (XPropertySet)xParagraphCursor;


                    //*********
                    XMultiServiceFactory mxDocFactory = (XMultiServiceFactory)myTextDocument;
                    XText xText = myTextDocument.getText();
                    XTextCursor xTextCursor = xText.createTextCursor();
                    XPropertySet xTextCursorProps = (XPropertySet)xTextCursor;

                    String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").Value.ToString();
                    XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier)myTextDocument;
                    XNameAccess xFamilies = (XNameAccess)xSupplier.getStyleFamilies();
                    Object Families = xFamilies.getByName("PageStyles").Value;
                    XNameContainer xFamily = (XNameContainer)Families;

                    Object Family = xFamily.getByName(pageStyleName).Value;
                    XStyle xStyle = (XStyle)Family;
                    // Get the property set of the TextCursor 
                    XPropertySet xStyleProps = (XPropertySet)xStyle;

                    xStyleProps.setPropertyValue("LeftMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("RightMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("TopMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("BottomMargin", new uno.Any(1000));

                    //***************
                    //Шрифт и размер
                    myPropertySet.setPropertyValue("CharFontName", new uno.Any("Courier"));
                    myPropertySet.setPropertyValue("CharHeight", new uno.Any(8));

                    int spi;
                    int sch_line;
                    int fl_fst = 1;
                    int priz;
                    DataTable tbl = ds.Tables[tblname];

                    foreach (DataRow drspi in ds.Tables["SPI"].Rows)
                    {
                        sch_line = 0;
                        if (fl_fst == 1)
                        {
                            sch_line = 1;
                            fl_fst = 0;
                        }
                        else
                        {
                            myPropertySet.setPropertyValue("BreakType", new uno.Any(typeof(BreakType), BreakType.PAGE_BEFORE));
                        }

                        myText.insertString(myText.getEnd(), "Реестр ответов на запросы СП-И в ПФР о наличии пенсии\r\r", false);
                        myText.insertString(myText.getEnd(), "Ответ из ПФР от " + DateTime.Today.ToShortDateString() + "\r\r", false);

                        spi = Convert.ToInt32(drspi["NOMSPI"]);

                        myText.insertString(myText.getEnd(), "СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "\r\r", false);
                        sch_line += 6;

                        foreach (DataRow row in tbl.Rows)
                        {                            
                            if (spi == Convert.ToInt32(row["NOMSPI"]))
                            {
                                //priz = Convert.ToInt32(row["PRIZ"]);

                                priz = 0;

                                if (!(int.TryParse(Convert.ToString(row["PRIZ"]), out priz)))
                                {
                                    priz = 2;
                                }

                                if (priz == 1)
                                {
                                    myText.insertString(myText.getEnd(), Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["NAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["FNAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["SNAMEDOL"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd() + "\r", false);
                                    myText.insertString(myText.getEnd(), Convert.ToString(row["ADRES"]).TrimEnd() + "\r", false);
                                    myText.insertString(myText.getEnd(), "Должник является получателем пенсии. Cумма дохода, на которую можно обратить взыскание: " + Convert.ToString(row["SUMMA"]).TrimEnd() + "\r\r", false);
                                    sch_line += 5;
                                }
                            }
                            
                        }

                        // если ничего положительного в ответах нет, то так и пишем
                        //if (sch_line == 6)
                        //{
                        //    par.Range.Text += "Нет положительных ответов по запросам о наличии пенсии у должников.";
                        //    sch_line++;
                        //    object oMissing = System.Reflection.Missing.Value;
                        //    par.Range.Delete(ref oMissing, ref oMissing);
                        //}

                    }

                    //Сохранение файла
                    //((XStorable)doc).storeToURL(PathConverter("c:\\Temp\\1.odt"), new unoidl.com.sun.star.beans.PropertyValue[0]);

                    //Закрытие файла
                    //((unoidl.com.sun.star.text.XTextDocument)doc).dispose();
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
            }/**/
        }

        public void OOo_Potd(string tblname, DataSet ds, OleDbConnection con, Form1 mainF)
        {
            FBcon = con;
            
            try
            {
                //**********Формирование**реестра**find************
                XComponent doc;
                XText myText;
                XTextCursor myTextCursor;
                XPropertySet myCursorProps;
                XParagraphCursor myParagCursor;
                XPropertySet myParagProps;
                XTextDocument myTextDocument;

                uno.Any myEnum;

                string[] par = new string[1];
                par[0] = "";
                if (isOOoInstalled())
                {
                    mxMSFactory = uno_connect(par);

                    doc = OOo3_initWriterDocument("", true);

                    myTextDocument = (XTextDocument)doc;
                    myText = myTextDocument.getText();

                    // create a paragraph cursor  
                    XParagraphCursor xParagraphCursor = (XParagraphCursor)(myText.createTextCursor());
                    XPropertySet myPropertySet = (XPropertySet)xParagraphCursor;


                    //*********
                    XMultiServiceFactory mxDocFactory = (XMultiServiceFactory)myTextDocument;
                    XText xText = myTextDocument.getText();
                    XTextCursor xTextCursor = xText.createTextCursor();
                    XPropertySet xTextCursorProps = (XPropertySet)xTextCursor;

                    String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").Value.ToString();
                    XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier)myTextDocument;
                    XNameAccess xFamilies = (XNameAccess)xSupplier.getStyleFamilies();
                    Object Families = xFamilies.getByName("PageStyles").Value;
                    XNameContainer xFamily = (XNameContainer)Families;

                    Object Family = xFamily.getByName(pageStyleName).Value;
                    XStyle xStyle = (XStyle)Family;
                    // Get the property set of the TextCursor 
                    XPropertySet xStyleProps = (XPropertySet)xStyle;

                    xStyleProps.setPropertyValue("LeftMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("RightMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("TopMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("BottomMargin", new uno.Any(1000));

                    //***************
                    //Шрифт и размер
                    myPropertySet.setPropertyValue("CharFontName", new uno.Any("Courier"));
                    myPropertySet.setPropertyValue("CharHeight", new uno.Any(8));

                    int spi;
                    int sch_line;
                    int fl_fst = 1;
                    string priz = "";
                    DataTable tbl = ds.Tables[tblname];

                    foreach (DataRow drspi in ds.Tables["SPI"].Rows)
                    {
                        sch_line = 0;
                        if (fl_fst == 1)
                        {
                            sch_line = 1;
                            fl_fst = 0;
                        }
                        else
                        {
                            myPropertySet.setPropertyValue("BreakType", new uno.Any(typeof(BreakType), BreakType.PAGE_BEFORE));
                        }

                        myText.insertString(myText.getEnd(), "Реестр ответов на запросы СП-И о персонифицированных данных в ПФР\r\r", false);
                        myText.insertString(myText.getEnd(), "Ответ из ПФР от " + DateTime.Today.ToShortDateString() + "\r\r", false);

                        spi = Convert.ToInt32(drspi["NOMSPI"]);

                        myText.insertString(myText.getEnd(), "СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "\r\r", false);
                        sch_line += 9;

                        foreach (DataRow row in tbl.Rows)
                        {                          
                            if (spi == Convert.ToInt32(row["NOMSPI"]))
                            {
                                myText.insertString(myText.getEnd(), Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FNAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["NAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["SNAMEDOL"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd() + "\r", false);
                                if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                                    myText.insertString(myText.getEnd(), "нет данных по должнику\r\r", false);    
                                else
                                {
                                    myText.insertString(myText.getEnd(), "Наименование страхователя: " + Convert.ToString(row["NAMEORG"]).TrimEnd() + "." + "\r", false);
                                    myText.insertString(myText.getEnd(), "Местонахождение страхователя: " + Convert.ToString(row["ADRORG"]).TrimEnd() + "." + "\r", false);
                                    myText.insertString(myText.getEnd(), "Дата начала периода работы: " + Convert.ToDateTime(row["DATST"]).ToShortDateString() + "." + "\r", false);
                                    myText.insertString(myText.getEnd(), "Дата окончания периода работы: " + Convert.ToDateTime(row["DATFN"]).ToShortDateString() + "."+"\r", false);
                                    myText.insertString(myText.getEnd(), "Комментарий: " + Convert.ToString(row["KOMMENT"]).TrimEnd() +"\r\r", false);
                                    sch_line++; 
                                }
                                sch_line += 3;
                            }                            
                        }
                    }

                    //Сохранение файла
                    //((XStorable)doc).storeToURL(PathConverter("c:\\Temp\\1.odt"), new unoidl.com.sun.star.beans.PropertyValue[0]);

                    //Закрытие файла
                    //((unoidl.com.sun.star.text.XTextDocument)doc).dispose();
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
        }

        public void OOo_Gbdd(DateTime DatZapr, DataTable DT_gibd_reg, OleDbConnection con, Form1 mainF)
        {
            FBcon = con;

            try
            {
                //**********Формирование**реестра**find************
                XComponent doc;
                XText myText;
                XTextCursor myTextCursor;
                XPropertySet myCursorProps;
                XParagraphCursor myParagCursor;
                XPropertySet myParagProps;
                XTextDocument myTextDocument;

                uno.Any myEnum;

                string[] par = new string[1];
                par[0] = "";
                if (isOOoInstalled())
                {
                    mxMSFactory = uno_connect(par);

                    doc = OOo3_initWriterDocument("", true);

                    myTextDocument = (XTextDocument)doc;
                    myText = myTextDocument.getText();

                    // create a paragraph cursor  
                    XParagraphCursor xParagraphCursor = (XParagraphCursor)(myText.createTextCursor());
                    XPropertySet myPropertySet = (XPropertySet)xParagraphCursor;


                    //*********
                    XMultiServiceFactory mxDocFactory = (XMultiServiceFactory)myTextDocument;
                    XText xText = myTextDocument.getText();
                    XTextCursor xTextCursor = xText.createTextCursor();
                    XPropertySet xTextCursorProps = (XPropertySet)xTextCursor;

                    String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").Value.ToString();
                    XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier)myTextDocument;
                    XNameAccess xFamilies = (XNameAccess)xSupplier.getStyleFamilies();
                    Object Families = xFamilies.getByName("PageStyles").Value;
                    XNameContainer xFamily = (XNameContainer)Families;

                    Object Family = xFamily.getByName(pageStyleName).Value;
                    XStyle xStyle = (XStyle)Family;
                    // Get the property set of the TextCursor 
                    XPropertySet xStyleProps = (XPropertySet)xStyle;

                    xStyleProps.setPropertyValue("LeftMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("RightMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("TopMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("BottomMargin", new uno.Any(1000));

                    //***************
                    //Шрифт и размер
                    myPropertySet.setPropertyValue("CharFontName", new uno.Any("Courier"));
                    myPropertySet.setPropertyValue("CharHeight", new uno.Any(8));

                    //int spi;
                    //int sch_line;
                    //int fl_fst = 1;
                    //string priz = "";
                    //DataTable tbl = ds.Tables[tblname];

                    //foreach (DataRow drspi in ds.Tables["SPI"].Rows)
                    //{
                    //    sch_line = 0;
                    //    if (fl_fst == 1)
                    //    {
                    //        sch_line = 1;
                    //        fl_fst = 0;
                    //    }
                    //    else
                    //    {
                    //        myPropertySet.setPropertyValue("BreakType", new uno.Any(typeof(BreakType), BreakType.PAGE_BEFORE));
                    //    }

                    //    myText.insertString(myText.getEnd(), "Реестр ответов на запросы СП-И о персонифицированных данных в ПФР\r\r", false);
                    //    myText.insertString(myText.getEnd(), "Ответ из ПФР от " + DateTime.Today.ToShortDateString() + "\r\r", false);

                    //    spi = Convert.ToInt32(drspi["NOMSPI"]);

                    //    myText.insertString(myText.getEnd(), "СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "\r\r", false);
                    //    sch_line += 9;

                    //    foreach (DataRow row in tbl.Rows)
                    //    {
                    //        if (spi == Convert.ToInt32(row["NOMSPI"]))
                    //        {
                    //            myText.insertString(myText.getEnd(), Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["FNAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["NAMEDOL"]).TrimEnd() + " " + Convert.ToString(row["SNAMEDOL"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd() + "\r", false);
                    //            if ((Convert.ToString(row["NAMEORG"]).TrimEnd()) == "")
                    //                myText.insertString(myText.getEnd(), "нет данных по должнику\r\r", false);
                    //            else
                    //            {
                    //                myText.insertString(myText.getEnd(), "Наименование страхователя: " + Convert.ToString(row["NAMEORG"]).TrimEnd() + "." + "\r", false);
                    //                myText.insertString(myText.getEnd(), "Местонахождение страхователя: " + Convert.ToString(row["ADRORG"]).TrimEnd() + "." + "\r", false);
                    //                myText.insertString(myText.getEnd(), "Дата начала периода работы: " + Convert.ToDateTime(row["DATST"]).ToShortDateString() + "." + "\r", false);
                    //                myText.insertString(myText.getEnd(), "Дата окончания периода работы: " + Convert.ToDateTime(row["DATFN"]).ToShortDateString() + "." + "\r", false);
                    //                myText.insertString(myText.getEnd(), "Комментарий: " + Convert.ToString(row["KOMMENT"]).TrimEnd() + "\r\r", false);
                    //                sch_line++;
                    //            }
                    //            sch_line += 3;
                    //        }
                    //    }
                    //}

                    //Microsoft.Office.Interop.Word.ApplicationClass app = new ApplicationClass();

                    //object s1 = "";
                    //object fl = false;
                    //object t = WdNewDocumentType.wdNewBlankDocument;
                    //object fl2 = true;

                    //Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref s1, ref fl, ref t, ref fl2);
                    //doc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                    //Paragraph par = doc.Content.Paragraphs[1];

                    //par.Range.Font.Name = "Courier";
                    //par.Range.Font.Size = 8;
                    //float a = par.Range.PageSetup.RightMargin;
                    //float b = par.Range.PageSetup.LeftMargin;
                    //float c = par.Range.PageSetup.TopMargin;
                    //float d = par.Range.PageSetup.BottomMargin;

                    //par.Range.PageSetup.RightMargin = 30;
                    //par.Range.PageSetup.LeftMargin = 30;
                    //par.Range.PageSetup.TopMargin = 20;
                    //par.Range.PageSetup.BottomMargin = 20;

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

                    mainF.prbWritingDBF.Value = 0;
                    mainF.prbWritingDBF.Maximum = DT_gibd_reg.Rows.Count;
                    //prbWritingDBF.Maximum = DT_gibd_rst.Rows.Count;
                    mainF.prbWritingDBF.Step = 1;

                    foreach (DataRow row in DT_gibd_reg.Rows)
                    //foreach (DataRow row in DT_gibd_rst.Rows)
                    {
                        if (spi < Convert.ToInt32(row["USCODE"]))
                        {
                            if (mainF.cb_prgibd.Checked == true)
                            {
                                while (sch_line > 61)
                                    sch_line = sch_line - 61;

                                nline = "";
                                for (int i = sch_line; i < 61; i++)
                                    myText.insertString(myText.getEnd(), "\r", false);
                                    //par.Range.Text += "";
                                //nline += "\n";
                                //par.Range.Text += Convert.ToString(i - 1);
                                //par.Range.Text += nline;
                            }
                            else
                            {
                                string bord = "";
                                for (int j = 0; j < 100; j++)
                                    bord += "*";
                                myText.insertString(myText.getEnd(), bord + "нет данных по должнику\r\r", false);
                                //par.Range.Text += bord;
                                //par.Range.Text += "";
                                sch_line++;
                                sch_line++;

                            }

                            spi = 999;
                        }
                        if (spi > Convert.ToInt32(row["USCODE"]))
                        {
                            if (mainF.cb_prgibd.Checked == false)
                            {
                                totline = sch_line;
                                while (totline > 61)
                                    totline = totline - 61;
                                idcnt = mainF.idCount(DT_gibd_reg, Convert.ToInt32(row["USCODE"]));
                                if ((idcnt + 11 + totline) > 61)
                                {
                                    for (int i = (totline); i < 61; i++)
                                    {
                                        //par.Range.Text += Convert.ToString(i - 1);
                                        myText.insertString(myText.getEnd(), "\r", false);
                                        //par.Range.Text += "";
                                        sch_line++;
                                    }
                                }
                            }

                            myText.insertString(myText.getEnd(), "РЕЕСТР ИСПОЛНИТЕЛЬНЫХ ДОКУМЕНТОВ, ДОЛГ ПО КОТОРЫМ ОПЛАЧЕН.\r", false);
                            myText.insertString(myText.getEnd(), "СФОРМИРОВАН НА ОСНОВЕ ДАННЫХ, ПОЛУЧЕННЫХ ИЗ ГИБДД.\r\r", false);
                            myText.insertString(myText.getEnd(), "ДАТА ФОРМИРОВАНИЯ " + DatZapr.ToShortDateString() + "\r\r", false);
                            //par.Range.Text += "РЕЕСТР ИСПОЛНИТЕЛЬНЫХ ДОКУМЕНТОВ, ДОЛГ ПО КОТОРЫМ ОПЛАЧЕН.";
                            //par.Range.Text += "СФОРМИРОВАН НА ОСНОВЕ ДАННЫХ, ПОЛУЧЕННЫХ ИЗ ГИБДД.\n";
                            //par.Range.Text += "ДАТА ФОРМИРОВАНИЯ " + DatZapr.ToShortDateString() + "\n";

                            spi = Convert.ToInt32(row["USCODE"]);

                            if (mainF.cb_prgibd.Checked == true)
                            {
                                sch_line = 0;
                                if (fl_fst == 1)
                                {
                                    sch_line = 1;
                                    fl_fst = 0;
                                }
                            }
                            //par.Range.Text += " ";
                            myText.insertString(myText.getEnd(), GetSpiName2(Convert.ToInt32(row["USCODE"])) + "\r\r", false);
                            myText.insertString(myText.getEnd(), "НОМЕР ИД       ДОЛЖНИК     ВЗЫСКАТЕЛЬ     НОМЕР ИП       ДАТА ВНЕСЕНИЯ В БАЗУ ГИБДД      ДАТА ЗАГРУЗКИ\r", false);
                            //par.Range.Text += GetSpiName2(Convert.ToInt32(row["USCODE"])) + "\n";
                            //par.Range.Text += "НОМЕР ИД       ДОЛЖНИК     ВЗЫСКАТЕЛЬ     НОМЕР ИП       ДАТА ВНЕСЕНИЯ В БАЗУ ГИБДД      ДАТА ЗАГРУЗКИ";
                            //par.Range.Text += GetOSP_Name();
                            sch_line += 8;
                        }
                        if (spi == Convert.ToInt32(row["USCODE"]))
                        {
                            // тупняк какой-то svn глючит!
                            string txtResponse = Convert.ToString(row["NOMID"]) + "  " + /*Money_ToStr(Convert.ToDecimal(row["summ"])) + "  " +*/ Convert.ToString(row["FIO_D"] + "  " + Convert.ToString(row["name_v"]) + "  " + Convert.ToString(row["NUM_IP"])) + "  " + Convert.ToString(Convert.ToDateTime(row["BASE_T"]).ToShortDateString()) + "  " + Convert.ToString(Convert.ToDateTime(row["DATE_Z"]).ToShortDateString());
                            //sch_line++;
                            //string txtResponse = Convert.ToString(row["BASE_T"]) + " " + Convert.ToString(row["DATE_Z"]);
                            myText.insertString(myText.getEnd(), txtResponse + "\r", false);
                            //par.Range.Text += txtResponse;
                            sch_line++;
                            if (txtResponse.Length > 200)
                            {
                                sch_line++; // если был перевод строки
                            }
                        }

                        mainF.prbWritingDBF.PerformStep();
                        mainF.prbWritingDBF.Refresh();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    //app.Visible = true;



                    //Сохранение файла
                    //((XStorable)doc).storeToURL(PathConverter("c:\\Temp\\1.odt"), new unoidl.com.sun.star.beans.PropertyValue[0]);

                    //Закрытие файла
                    //((unoidl.com.sun.star.text.XTextDocument)doc).dispose();
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
        }

        public void OOo_Krc(string tblname, DataSet ds, OleDbConnection con, Form1 mainF)
        {
            FBcon = con;

            try
            {
                //**********Формирование**реестра**find************
                XComponent doc;
                XText myText;
                XTextCursor myTextCursor;
                XPropertySet myCursorProps;
                XParagraphCursor myParagCursor;
                XPropertySet myParagProps;
                XTextDocument myTextDocument;

                uno.Any myEnum;

                string[] par = new string[1];
                par[0] = "";
                if (isOOoInstalled())
                {
                    mxMSFactory = uno_connect(par);

                    doc = OOo3_initWriterDocument("", true);

                    myTextDocument = (XTextDocument)doc;
                    myText = myTextDocument.getText();

                    // create a paragraph cursor  
                    XParagraphCursor xParagraphCursor = (XParagraphCursor)(myText.createTextCursor());
                    XPropertySet myPropertySet = (XPropertySet)xParagraphCursor;


                    //*********
                    XMultiServiceFactory mxDocFactory = (XMultiServiceFactory)myTextDocument;
                    XText xText = myTextDocument.getText();
                    XTextCursor xTextCursor = xText.createTextCursor();
                    XPropertySet xTextCursorProps = (XPropertySet)xTextCursor;

                    String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").Value.ToString();
                    XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier)myTextDocument;
                    XNameAccess xFamilies = (XNameAccess)xSupplier.getStyleFamilies();
                    Object Families = xFamilies.getByName("PageStyles").Value;
                    XNameContainer xFamily = (XNameContainer)Families;

                    Object Family = xFamily.getByName(pageStyleName).Value;
                    XStyle xStyle = (XStyle)Family;
                    // Get the property set of the TextCursor 
                    XPropertySet xStyleProps = (XPropertySet)xStyle;

                    xStyleProps.setPropertyValue("LeftMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("RightMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("TopMargin", new uno.Any(1000));
                    xStyleProps.setPropertyValue("BottomMargin", new uno.Any(1000));

                    //***************
                    //Шрифт и размер
                    myPropertySet.setPropertyValue("CharFontName", new uno.Any("Courier"));
                    myPropertySet.setPropertyValue("CharHeight", new uno.Any(8));

                    int spi;
                    int sch_line;
                    int fl_fst = 1;
                    int priz;
                    DataTable tbl = ds.Tables[tblname];

                    foreach (DataRow drspi in ds.Tables["SPI"].Rows)
                    {
                        sch_line = 0;
                        if (fl_fst == 1)
                        {
                            sch_line = 1;
                            fl_fst = 0;
                        }
                        else
                        {
                            myPropertySet.setPropertyValue("BreakType", new uno.Any(typeof(BreakType), BreakType.PAGE_BEFORE));
                        }

                        myText.insertString(myText.getEnd(), "Реестр оплаченных сумм из КРЦ\r\r", false);
                        myText.insertString(myText.getEnd(), "Ответ из КРЦ от " + DateTime.Today.ToShortDateString() + "\r\r", false);

                        spi = Convert.ToInt32(drspi["NOMSPI"]);

                        myText.insertString(myText.getEnd(), "СП-И: " + GetSpiName2(Convert.ToInt32(drspi["NOMSPI"])) + "\r\r", false);
                        sch_line += 6;

                        foreach (DataRow row in tbl.Rows)
                        {
                            if (spi == Convert.ToInt32(row["NOMSPI"]))
                            {
                                if (Convert.ToInt32(row["SUMPL"])!=0)
                                {
                                    myText.insertString(myText.getEnd(), Convert.ToString(row["ZAPROS"]).TrimEnd() + " " + Convert.ToString(row["NAMEDOL"]).TrimEnd() + " " + Convert.ToDateTime(row["BORN"]).ToShortDateString().TrimEnd() + "\r", false);
                                    myText.insertString(myText.getEnd(), Convert.ToString(row["ADRES"]).TrimEnd() + "\r", false);
                                    myText.insertString(myText.getEnd(), "Должник выплатил сумму задолженности: " + Convert.ToString(row["SUMPL"]) + " . Дата выплаты: " + Convert.ToString(row["DATPL"]) + "\r\r", false);
                                    sch_line += 5;
                                }
                            }

                        }

                        // если ничего положительного в ответах нет, то так и пишем
                        //if (sch_line == 6)
                        //{
                        //    par.Range.Text += "Нет положительных ответов по запросам о наличии пенсии у должников.";
                        //    sch_line++;
                        //    object oMissing = System.Reflection.Missing.Value;
                        //    par.Range.Delete(ref oMissing, ref oMissing);
                        //}

                    }

                    //Сохранение файла
                    //((XStorable)doc).storeToURL(PathConverter("c:\\Temp\\1.odt"), new unoidl.com.sun.star.beans.PropertyValue[0]);

                    //Закрытие файла
                    //((unoidl.com.sun.star.text.XTextDocument)doc).dispose();
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
        }



        private String GetSpiName2(decimal USCODE)
        {
            String res = "";
            try
            {
                FBcon.Open();
                OleDbTransaction tran = FBcon.BeginTransaction(IsolationLevel.ReadCommitted);
                //OleDbCommand cmd = new OleDbCommand("Select FULL_NAME from S_USERS WHERE USCODE = '" + Convert.ToString((int)USCODE) + "'", FBcon, tran);
                OleDbCommand cmd = new OleDbCommand("select suser_fio from spi left join sys_users on spi.suser_id = sys_users.suser_id where spi.SPI_ZONENUM = " + Convert.ToString((int)USCODE), FBcon, tran);
                res = Convert.ToString(cmd.ExecuteScalar());
                tran.Rollback();
                FBcon.Close();

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

        private string Money_ToStr(decimal nMoney)
        {
            string txtResult = "";
            txtResult = nMoney.ToString("N2").Replace(".", " руб. ");
            txtResult = txtResult.Replace(",", " руб. ") + " коп.";

            return txtResult;
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


    }
}
