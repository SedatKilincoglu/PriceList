using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using System.Reflection;
using System.IO;

namespace PriceList
{
    public class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        public static SAPbouiCOM.Application SBO_Application;
        public static SAPbobsCOM.Company diCompany;
        public static List<Models.Paym> paymList = new List<Models.Paym>();
        public static bool AuthorizedUser = false;
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }
                ConnectToUI();
                fillPaymList();
                OrganizeTables();
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                

                SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                oApp.Run();
               
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void ConnectToUI()
        {
            SAPbouiCOM.SboGuiApi SboGuiApi;
            string sConnectionString;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

            SboGuiApi.Connect(sConnectionString);
            SBO_Application = SboGuiApi.GetApplication();
            ConnectwithSharedMemory();
        }
        private static void ConnectwithSharedMemory()
        {
            diCompany = (SAPbobsCOM.Company)Program.SBO_Application.Company.GetDICompany();
        }

        private static void fillPaymList()
        {
            SAPbobsCOM.Recordset PaymentRecordSet;
            PaymentRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string strQuery = "SELECT REPLACE(GroupNum,'-','N') AS PCode, PymntGroup AS PName FROM OCTG ORDER BY ExtraMonth,ExtraDays";
            PaymentRecordSet.DoQuery(strQuery);
            PaymentRecordSet.MoveFirst();

            while (!PaymentRecordSet.EoF)
            {
                Models.Paym payment = new Models.Paym();
                payment.paymCode = "paym" + PaymentRecordSet.Fields.Item("PCode").Value.ToString();
                payment.paymName = PaymentRecordSet.Fields.Item("PName").Value.ToString();
                paymList.Add(payment);
                PaymentRecordSet.MoveNext();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(PaymentRecordSet);
            PaymentRecordSet = null;
            GC.Collect();
        }


        private static void OrganizeTables()
        {
            CreateTable("SML_PRCHEAD", "Fiyat Listeleri", SAPbobsCOM.BoUTBTableType.bott_Document);
            CreateField("SML_PRCHEAD", "ValidFrom", "Geçerlilik Başlangıcı", SAPbobsCOM.BoFieldTypes.db_Date, 100);
            CreateField("SML_PRCHEAD", "ValidUntil", "Geçerlilik Bitişi", SAPbobsCOM.BoFieldTypes.db_Date, 100);
            CreateField("SML_PRCHEAD", "Description", "Açıklama", SAPbobsCOM.BoFieldTypes.db_Alpha, 200);

            CreateTable("SML_PRCITEM", "Fiyat Listesi Detay", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            CreateField("SML_PRCITEM", "ItemCode", "Kalem Tanıtıcı", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            CreateField("SML_PRCITEM", "ItemName", "Kalem Açıklama", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
            CreateField("SML_PRCITEM", "Price", "Fiyat", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price);
            CreateField("SML_PRCITEM", "Currency", "Döviz Cinsi", SAPbobsCOM.BoFieldTypes.db_Alpha, 3); 

            CreateUDO("SML_PRCHEAD", "SML_PRCITEM","Fiyat Listeleri", SAPbobsCOM.BoUDOObjType.boud_Document);

            CreateTable("SML_DSCHEAD", "Dönemsel İndirimler", SAPbobsCOM.BoUTBTableType.bott_Document);
            CreateField("SML_DSCHEAD", "ValidFrom", "Geçerlilik Başlangıcı", SAPbobsCOM.BoFieldTypes.db_Date, 100);
            CreateField("SML_DSCHEAD", "ValidUntil", "Geçerlilik Bitişi", SAPbobsCOM.BoFieldTypes.db_Date, 100);
            CreateField("SML_DSCHEAD", "Description", "Açıklama", SAPbobsCOM.BoFieldTypes.db_Alpha, 200);

            CreateTable("SML_DSCITEM", "Dönemsel İndirim Detay", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            CreateField("SML_DSCITEM", "ItemCode", "Kalem Tanıtıcı", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            CreateField("SML_DSCITEM", "ItemName", "Kalem Açıklama", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
            CreateField("SML_DSCITEM", "AdditionalDiscount", "Ek İndirim Oranı Hakkı", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage);
            foreach(Models.Paym paym in paymList)
            {
                CreateField("SML_DSCITEM", paym.paymCode, "İndirim Oranı", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage);
            }

            CreateUDO("SML_DSCHEAD", "SML_DSCITEM","Dönemsel İndirimler", SAPbobsCOM.BoUDOObjType.boud_Document);

            CreateTable("SML_PRCAUTH", "Fiyat Listesi yetkilendirme", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            CreateUDO("SML_PRCAUTH", "", "İndirim Yetkilendirme Tablosu", SAPbobsCOM.BoUDOObjType.boud_MasterData);
            ExecuteSqlScripts();
        }

        private static void ExecuteSqlScripts()
        {
            try
            {
                SAPbobsCOM.Company bobsCompany;
                bobsCompany = (SAPbobsCOM.Company)Program.diCompany;
                SAPbobsCOM.Recordset defaultRecordSet;
                defaultRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Assembly assembly = Assembly.GetExecutingAssembly();
                string fileName = "PriceList.SQL.InsertScript.sql";
                string sqlScript;
                using (Stream stream = assembly.GetManifestResourceStream(fileName))
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        sqlScript = reader.ReadToEnd();
                    }
                }
                defaultRecordSet.DoQuery(sqlScript);
                SAPbobsCOM.Recordset userRecordSet;
                userRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sqlquery = "SELECT * FROM [@SML_PRCAUTH] WHERE Code = '" + bobsCompany.UserName + "'";
                defaultRecordSet.DoQuery(sqlquery);
                if (defaultRecordSet.RecordCount > 0)
                {
                    AuthorizedUser = true;
                }
            }



            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }


        }

        private static void CreateTable(string TableName, string TableDescription, SAPbobsCOM.BoUTBTableType tableType)
        {
            SAPbobsCOM.UserTablesMD oUDT;

            oUDT = (SAPbobsCOM.UserTablesMD)diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            try
            {
                if (oUDT.GetByKey(TableName) == false)
                {
                    oUDT.TableName = TableName;
                    oUDT.TableDescription = TableDescription;
                    oUDT.TableType = tableType;
                    int ret = oUDT.Add();
                    if (ret == 0)
                    {
                        SBO_Application.StatusBar.SetText("Add Table: " + oUDT.TableName + " successfull", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                    else
                    {
                        SBO_Application.StatusBar.SetText("Add Table error: " + diCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDT);
                oUDT = null;
                GC.Collect();
            }
        }

        private static void CreateField(string MyTableName, string MyFieldName, string MyFieldDescrition, SAPbobsCOM.BoFieldTypes MyFieldType, int MyFieldSize, SAPbobsCOM.BoFldSubTypes MyFieldSubType = SAPbobsCOM.BoFldSubTypes.st_None)
        {
            SAPbobsCOM.UserFieldsMD oUDF;
            oUDF = (SAPbobsCOM.UserFieldsMD)diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            try
            {
                oUDF.TableName = MyTableName;
                oUDF.Name = MyFieldName;
                oUDF.Description = MyFieldDescrition;
                oUDF.Type = MyFieldType;
                oUDF.SubType = MyFieldSubType;
                if (MyFieldSize > 0)
                    oUDF.EditSize = MyFieldSize;

                if (MyFieldSize > 0)
                {
                    oUDF.EditSize = MyFieldSize;
                }

                int ret = oUDF.Add();
                if (ret == 0)
                {
                    SBO_Application.StatusBar.SetText("UDF " + oUDF.Name + " added.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDF);
                oUDF = null;
                GC.Collect();
            }
        }

        private static void CreateUDO(String MainTable, String ChildTable,String MenuCaption, SAPbobsCOM.BoUDOObjType ObjectType)
        {
            String UdoName = "UDO" + MainTable;
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            SAPbobsCOM.UserObjectMD_FindColumns oUDOFind = null;
            SAPbobsCOM.UserObjectMD_FormColumns oUDOForm = null;
            SAPbobsCOM.UserObjectMD_EnhancedFormColumns oUDOEnhancedForm = null;
            GC.Collect();
            oUserObjectMD = diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
            oUDOFind = oUserObjectMD.FindColumns;
            oUDOForm = oUserObjectMD.FormColumns;
            oUDOEnhancedForm = oUserObjectMD.EnhancedFormColumns;
            var retval = oUserObjectMD.GetByKey(UdoName);
            if (!retval)
            {
                oUserObjectMD.Code = UdoName;
                oUserObjectMD.Name = UdoName;
                oUserObjectMD.TableName = MainTable;

                oUserObjectMD.ObjectType = ObjectType;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.MenuItem = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.MenuCaption = MenuCaption;
                if (ChildTable != "")
                {
                    oUserObjectMD.ChildTables.TableName = ChildTable;
                    oUserObjectMD.ChildTables.Add();
                }

                oUDOFind.ColumnAlias = "DocEntry";
                oUDOFind.ColumnDescription = "DocEntry";
                oUDOFind.Add();
                if (!retval)
                {
                    try
                    {
                        int rv = oUserObjectMD.Add();
                        if (rv < 0)
                        {
                            SBO_Application.StatusBar.SetText("Exception: " + diCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }

                }

            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = null;
            GC.Collect();
        }


        // Global bir bayrak değişkeni ekle
        private static bool isUpdating = false;
        private static readonly HashSet<string> SalesForms = new HashSet<string> { "139", "140", "133", "149" }; // 139: Satış Siparişi, 140: İrsaliye, 133: Fatura

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Sonsuz döngüyü önlemek için kontrol
                
                if (isUpdating) return;

                //Malzeme kodu değiştiğinde
                if (SalesForms.Contains(pVal.FormTypeEx) && pVal.ItemUID == "38" && pVal.ColUID == "1" &&
                    pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.BeforeAction)
                {
                    var oForm = SBO_Application.Forms.Item(FormUID);
                    var oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    var docDate = (SAPbouiCOM.EditText)oForm.Items.Item("46").Specific;
                    var itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific);
                    var paymCode = ((SAPbouiCOM.ComboBox)oForm.Items.Item("47").Specific).Value;
                    paymCode = paymCode.Trim();
                    paymCode = paymCode.Replace("-", "N");
                    paymCode = "U_Paym" + paymCode;
                    var price = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific);
                    var discount = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific);
                    // Fiyat sorgusu
                    SAPbobsCOM.Recordset prcRecordSet;
                    prcRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string strQuery = "SELECT OITM.ItemCode,OITM.ItemName,PRC.PrcValidFrom,PRC.PrcValidUntil,PRC.PrcDescription, PRC.PrcPrice,PRC.PrcCurrency,DSC.* FROM OITM ";
                    strQuery = strQuery + "LEFT JOIN ";
                    strQuery = strQuery + "(SELECT PRCHEAD.U_ValidFrom AS PrcValidFrom, PRCHEAD.U_ValidUntil AS PrcValidUntil, PRCHEAD.U_Description AS PrcDescription, PRCITEM.U_ItemCode AS PrcItemCode, PRCITEM.U_Price AS PrcPrice, PRCITEM.U_Currency AS PrcCurrency FROM[@SML_PRCHEAD] PRCHEAD INNER JOIN[@SML_PRCITEM] PRCITEM ON PRCITEM.DocEntry = PRCHEAD.DocEntry) PRC ";
                    strQuery = strQuery + " ON PRC.PrcItemCode = OITM.ItemCode ";
                    strQuery = strQuery + " LEFT JOIN ";
                    strQuery = strQuery + " (SELECT DSCHEAD.U_ValidFrom AS DscValidFrom, DSCHEAD.U_ValidUntil AS DscValidUntil,DSCHEAD.U_Description AS DscDescription,DSCITEM.* FROM [@SML_DSCHEAD] DSCHEAD INNER JOIN [@SML_DSCITEM] DSCITEM ON DSCITEM.DocEntry = DSCHEAD.DocEntry) DSC ";
                    strQuery = strQuery + " ON  DSC.U_ItemCode = OITM.ItemCode ";
                    strQuery = strQuery + $"where OITM.ItemCode = '{itemCode.Value}' AND PRC.PrcValidFrom <= '{docDate.Value}' AND PRC.PrcValidUntil >= '{docDate.Value}' ";

                    prcRecordSet.DoQuery(strQuery);

                    if (!prcRecordSet.EoF)
                    {
                        // Bayrağı ayarla (sonsuz döngüyü engellemek için)
                        isUpdating = true;
                        price.Value = prcRecordSet.Fields.Item("PrcPrice").Value.ToString() + " " + prcRecordSet.Fields.Item("PrcCurrency").Value.ToString();
                        var disc = prcRecordSet.Fields.Item(paymCode).Value.ToString();
                        disc = disc.Replace(',', '.');
                        discount.Value = disc;
                    }
                }

                //Belge Tarihi ya da Ödeme Koşulu değiştiğinde
                if ((SalesForms.Contains(pVal.FormTypeEx) && 
                    pVal.ItemUID == "47" && 
                    pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && 
                    !pVal.BeforeAction) 
                    ||
                    (SalesForms.Contains(pVal.FormTypeEx) && 
                    (pVal.ItemUID == "46" || pVal.ItemUID == "10") && 
                    pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && 
                    !pVal.BeforeAction))
                {
                    var oForm = SBO_Application.Forms.Item(FormUID);
                    var docDate = (SAPbouiCOM.EditText)oForm.Items.Item("46").Specific;
                    var paymCode = ((SAPbouiCOM.ComboBox)oForm.Items.Item("47").Specific).Value;
                    var beforePanelLevel = oForm.PaneLevel;
                    paymCode = paymCode.Trim();
                    if (paymCode == "-9")
                    {
                        return;
                    }
                    paymCode = paymCode.Replace("-", "N");
                    paymCode = "U_Paym" + paymCode;
                    isUpdating = true;
                    if (oForm.PaneLevel != 1)
                    {
                        oForm.PaneLevel = 1;
                    }
                    
                    var oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    oMatrix.Item.Enabled = true;
                    for (var i = 0; i < oMatrix.RowCount; i ++)
                    {
                        var itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i + 1).Specific);
                        var price = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(i + 1).Specific);
                        var discount = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(i + 1).Specific);
                        SAPbobsCOM.Recordset prcRecordSet;
                        prcRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string strQuery = "SELECT OITM.ItemCode,OITM.ItemName,PRC.PrcValidFrom,PRC.PrcValidUntil,PRC.PrcDescription, PRC.PrcPrice,PRC.PrcCurrency,DSC.* FROM OITM ";
                        strQuery = strQuery + "LEFT JOIN ";
                        strQuery = strQuery + "(SELECT PRCHEAD.U_ValidFrom AS PrcValidFrom, PRCHEAD.U_ValidUntil AS PrcValidUntil, PRCHEAD.U_Description AS PrcDescription, PRCITEM.U_ItemCode AS PrcItemCode, PRCITEM.U_Price AS PrcPrice, PRCITEM.U_Currency AS PrcCurrency FROM[@SML_PRCHEAD] PRCHEAD INNER JOIN[@SML_PRCITEM] PRCITEM ON PRCITEM.DocEntry = PRCHEAD.DocEntry) PRC ";
                        strQuery = strQuery + " ON PRC.PrcItemCode = OITM.ItemCode ";
                        strQuery = strQuery + " LEFT JOIN ";
                        strQuery = strQuery + " (SELECT DSCHEAD.U_ValidFrom AS DscValidFrom, DSCHEAD.U_ValidUntil AS DscValidUntil,DSCHEAD.U_Description AS DscDescription,DSCITEM.* FROM [@SML_DSCHEAD] DSCHEAD INNER JOIN [@SML_DSCITEM] DSCITEM ON DSCITEM.DocEntry = DSCHEAD.DocEntry) DSC ";
                        strQuery = strQuery + " ON  DSC.U_ItemCode = OITM.ItemCode ";
                        strQuery = strQuery + $"where OITM.ItemCode = '{itemCode.Value}' AND PRC.PrcValidFrom <= '{docDate.Value}' AND PRC.PrcValidUntil >= '{docDate.Value}' ";

                        prcRecordSet.DoQuery(strQuery);

                        if (!prcRecordSet.EoF)
                        {
                            // Bayrağı ayarla (sonsuz döngüyü engellemek için)
                            price.Item.Enabled = true;
                            discount.Item.Enabled = true;
                            price.Value = prcRecordSet.Fields.Item("PrcPrice").Value.ToString() + " " + prcRecordSet.Fields.Item("PrcCurrency").Value.ToString();
                            var a = prcRecordSet.Fields.Item(paymCode).Value.ToString();
                            discount.Value = prcRecordSet.Fields.Item(paymCode).Value.ToString();
                        }
                    }
                    if (oForm.PaneLevel != beforePanelLevel)
                    {
                        oForm.PaneLevel = beforePanelLevel;
                    }
                }
                if (SalesForms.Contains(pVal.FormTypeEx) && pVal.ItemUID == "38" && pVal.ColUID == "15" &&
                    pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.BeforeAction)
                {
                    var oForm = SBO_Application.Forms.Item(FormUID);
                    var oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    var docDate = (SAPbouiCOM.EditText)oForm.Items.Item("46").Specific;
                    var itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific);
                    var paymCode = ((SAPbouiCOM.ComboBox)oForm.Items.Item("47").Specific).Value;
                    var discount = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific);
                    paymCode = paymCode.Trim();
                    paymCode = paymCode.Replace("-", "N");
                    paymCode = "U_Paym" + paymCode;
                    SAPbobsCOM.Recordset prcRecordSet;
                    prcRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string strQuery = "SELECT DSCHEAD.U_ValidFrom AS DscValidFrom, DSCHEAD.U_ValidUntil AS DscValidUntil,DSCHEAD.U_Description AS DscDescription,DSCITEM.* FROM [@SML_DSCHEAD] DSCHEAD INNER JOIN [@SML_DSCITEM] DSCITEM ON DSCITEM.DocEntry = DSCHEAD.DocEntry ";
                    strQuery = strQuery + $"where DSCITEM.U_ItemCode = '{itemCode.Value}' AND DSCHEAD.U_ValidFrom <= '{docDate.Value}' AND DSCHEAD.U_ValidUntil >= '{docDate.Value}' ";
                    prcRecordSet.DoQuery(strQuery);
                    if (!prcRecordSet.EoF)
                    {
                        float maxDiscount = float.Parse(prcRecordSet.Fields.Item(paymCode).Value.ToString()) + float.Parse(prcRecordSet.Fields.Item("U_AdditionalDiscount").Value.ToString());
                        float activeDiscount = float.Parse(discount.String);
                        if (activeDiscount > maxDiscount)
                        {
                            SBO_Application.StatusBar.SetText($"Maksimum indirim oranı aşıldı", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                            return;
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText($"Hata: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                
                isUpdating = false;
            }
        }



        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
