using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using System.Reflection;
using System.IO;
using System.Threading;

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
        // Global bir bayrak değişkeni ekle
        private static bool isUpdating = false;
        private static readonly HashSet<string> SalesForms = new HashSet<string> { "149", "139","540000988" }; //149: Satış Teklifi, 139: Satın alma Siparişi,540000988: Satın alma teklifi
        // Form ID ve satır numarasına göre önceki ItemCode değerlerini tutacak dictionary
        private static Dictionary<string, Dictionary<int, string>> previousItemCodes = new Dictionary<string, Dictionary<int, string>>();
        public static Dictionary<string,int> colors = new Dictionary<string,int>();
        [STAThread]

        static int GetBGRColorValue(byte r, byte g, byte b)
        {
            return (b << 16) | (g << 8) | r;
        }
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
                FillPaymList();
                OrganizeTables();
                Initialize();
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;

                oApp.Run();

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public static void Initialize()
        {
            colors.Add("green", GetBGRColorValue(93, 155, 106));
            colors.Add("yellow", GetBGRColorValue(213, 218, 124));
            colors.Add("red", GetBGRColorValue(199, 119, 118));
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

        private static void FillPaymList()
        {
            SAPbobsCOM.Recordset PaymentRecordSet;
            PaymentRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string strQuery = "SELECT REPLACE(GroupNum,'-','N') AS PCode, PymntGroup AS PName FROM OCTG ORDER BY ExtraMonth,ExtraDays";
            PaymentRecordSet.DoQuery(strQuery);
            PaymentRecordSet.MoveFirst();

            while (!PaymentRecordSet.EoF)
            {
                Models.Paym payment = new Models.Paym
                {
                    paymCode = "paym" + PaymentRecordSet.Fields.Item("PCode").Value.ToString(),
                    paymName = PaymentRecordSet.Fields.Item("PName").Value.ToString()
                };
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

            CreateUDO("SML_PRCHEAD", "SML_PRCITEM", "Fiyat Listeleri", SAPbobsCOM.BoUDOObjType.boud_Document);

            CreateTable("SML_DSCHEAD", "Dönemsel İndirimler", SAPbobsCOM.BoUTBTableType.bott_Document);
            CreateField("SML_DSCHEAD", "ValidFrom", "Geçerlilik Başlangıcı", SAPbobsCOM.BoFieldTypes.db_Date, 100);
            CreateField("SML_DSCHEAD", "ValidUntil", "Geçerlilik Bitişi", SAPbobsCOM.BoFieldTypes.db_Date, 100);
            CreateField("SML_DSCHEAD", "Description", "Açıklama", SAPbobsCOM.BoFieldTypes.db_Alpha, 200);

            CreateTable("SML_DSCITEM", "Dönemsel İndirim Detay", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            CreateField("SML_DSCITEM", "ItemCode", "Kalem Tanıtıcı", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            CreateField("SML_DSCITEM", "ItemName", "Kalem Açıklama", SAPbobsCOM.BoFieldTypes.db_Alpha, 254);
            //CreateField("SML_DSCITEM", "AdditionalDiscount", "Ek İndirim Oranı Hakkı", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage);
            foreach (Models.Paym paym in paymList)
            {
                CreateField("SML_DSCITEM", paym.paymCode, "İndirim Oranı", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage);
                CreateField("SML_DSCITEM", "AddDisc" + paym.paymCode, "Ek İndirim Oranı Hakkı", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage);
            }

            CreateUDO("SML_DSCHEAD", "SML_DSCITEM", "Dönemsel İndirimler", SAPbobsCOM.BoUDOObjType.boud_Document);

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

        private static void CreateUDO(String MainTable, String ChildTable, String MenuCaption, BoUDOObjType ObjectType)
        {
            String UdoName = "UDO" + MainTable;
            GC.Collect();
            UserObjectsMD oUserObjectMD = diCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD) as UserObjectsMD;
            UserObjectMD_FindColumns oUDOFind = oUserObjectMD.FindColumns;
            var retval = oUserObjectMD.GetByKey(UdoName);
            if (!retval)
            {
                oUserObjectMD.Code = UdoName;
                oUserObjectMD.Name = UdoName;
                oUserObjectMD.TableName = MainTable;
                oUserObjectMD.ObjectType = ObjectType;
                oUserObjectMD.CanFind = BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = BoYesNoEnum.tYES;
                oUserObjectMD.CanCancel = BoYesNoEnum.tNO;
                oUserObjectMD.CanClose = BoYesNoEnum.tNO;
                oUserObjectMD.CanYearTransfer = BoYesNoEnum.tNO;
                oUserObjectMD.CanLog = BoYesNoEnum.tNO;
                oUserObjectMD.ManageSeries = BoYesNoEnum.tNO;
                oUserObjectMD.CanCreateDefaultForm = BoYesNoEnum.tYES;
                oUserObjectMD.MenuItem = BoYesNoEnum.tYES;
                oUserObjectMD.EnableEnhancedForm = BoYesNoEnum.tNO;
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



        private static Recordset GetRecordSet(SAPbouiCOM.Form oForm, string itemCode)
        {
            var docDate = ((SAPbouiCOM.EditText)oForm.Items.Item("46").Specific).Value;
            Recordset prcRecordSet;
            prcRecordSet = (Recordset)Program.diCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string strQuery = "SELECT OITM.ItemCode,OITM.ItemName,PRC.PrcValidFrom,PRC.PrcValidUntil,PRC.PrcDescription, PRC.PrcPrice,PRC.PrcCurrency,DSC.* FROM OITM ";
            strQuery += "LEFT JOIN ";
            strQuery += "(SELECT PRCHEAD.U_ValidFrom AS PrcValidFrom, PRCHEAD.U_ValidUntil AS PrcValidUntil, PRCHEAD.U_Description AS PrcDescription, PRCITEM.U_ItemCode AS PrcItemCode, PRCITEM.U_Price AS PrcPrice, PRCITEM.U_Currency AS PrcCurrency FROM[@SML_PRCHEAD] PRCHEAD INNER JOIN[@SML_PRCITEM] PRCITEM ON PRCITEM.DocEntry = PRCHEAD.DocEntry) PRC ";
            strQuery += " ON PRC.PrcItemCode = OITM.ItemCode ";
            strQuery += " LEFT JOIN ";
            strQuery += $" (SELECT DSCHEAD.U_ValidFrom AS DscValidFrom, DSCHEAD.U_ValidUntil AS DscValidUntil,DSCHEAD.U_Description AS DscDescription,DSCITEM.* FROM [@SML_DSCHEAD] DSCHEAD INNER JOIN [@SML_DSCITEM] DSCITEM ON DSCITEM.DocEntry = DSCHEAD.DocEntry AND DSCHEAD.U_ValidFrom <= '{docDate}' AND DSCHEAD.U_ValidUntil >= '{docDate}') DSC ";
            strQuery += " ON  DSC.U_ItemCode = OITM.ItemCode ";
            strQuery += $"where OITM.ItemCode = '{itemCode}' AND PRC.PrcValidFrom <= '{docDate}' AND PRC.PrcValidUntil >= '{docDate}' ";
            prcRecordSet.DoQuery(strQuery);
            return prcRecordSet;
        }


        private static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //SBO_Application.StatusBar.SetText($"Form Event: {BusinessObjectInfo.EventType} Before Action: {BusinessObjectInfo.BeforeAction}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

            if (SalesForms.Contains(BusinessObjectInfo.FormTypeEx))
            {
                if (previousItemCodes.ContainsKey(BusinessObjectInfo.FormUID))
                {
                    previousItemCodes[BusinessObjectInfo.FormUID].Clear();
                }


                var oForm = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                var oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                if (!previousItemCodes.ContainsKey(BusinessObjectInfo.FormUID))
                {
                    previousItemCodes[BusinessObjectInfo.FormUID] = new Dictionary<int, string>();
                }

                // Matrix'teki tüm satırları kontrol et
                for (int row = 1; row <= oMatrix.RowCount; row++)
                {
                    var itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(row).Specific).Value;
                    if (!string.IsNullOrEmpty(itemCode))
                    {
                        previousItemCodes[BusinessObjectInfo.FormUID][row] = itemCode;
                    }
                }
            }

        }

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormTypeEx == "PriceList.PriceListReport" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ItemUID == "gr_List" && pVal.ColUID == "Kayıt No" && pVal.BeforeAction == false)
            {
                SAPbouiCOM.Form repForm;
                repForm = SBO_Application.Forms.Item(FormUID);
                var grid = (SAPbouiCOM.Grid)repForm.Items.Item("gr_List").Specific;
                var DocEntry = grid.DataTable.GetValue("Kayıt No", pVal.Row).ToString();
                var objType = grid.DataTable.GetValue("Tip", pVal.Row).ToString();
                if (objType == "Fiyat Listesi")
                {
                    Form1 activeForm = new Form1(DocEntry);
                    activeForm.Show();
                }
                else
                {
                    DiscountForm activeForm = new DiscountForm(DocEntry);
                    activeForm.Show();
                }

            }
            if (!SalesForms.Contains(pVal.FormTypeEx) || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD && !pVal.BeforeAction))
            {
                return;
            }
            SAPbouiCOM.Form oForm;
            oForm = SBO_Application.Forms.Item(FormUID);
            oForm.Freeze(true);
            try
            {

                //CardCode focus aldığında yeni kayıt. Tablonu temizle
                if (pVal.ItemUID == "4" &&
                         pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && !pVal.BeforeAction)
                {
                    var cardCode = (SAPbouiCOM.EditText)oForm.Items.Item("4").Specific;
                    if (cardCode.Value == "")
                    {
                        if (previousItemCodes.ContainsKey(pVal.FormUID))
                        {
                            previousItemCodes[pVal.FormUID].Clear();
                        }
                    }

                }

                // Sonsuz döngüyü önlemek için kontrol
                if (isUpdating) return;

                // Form kapanışını yakala
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE &&
                    !pVal.BeforeAction)
                {
                    // Form kapanırken dictionary'den temizle
                    if (previousItemCodes.ContainsKey(FormUID))
                    {
                        previousItemCodes.Remove(FormUID);
                    }
                }

                //Malzeme kodu ya da fiyat değiştiğinde
                if (pVal.ItemUID == "38" && pVal.ColUID == "1" &&
                    pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.BeforeAction)
                {

                    var oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    var itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific);
                    // Form için dictionary yoksa oluştur
                    if (!previousItemCodes.ContainsKey(FormUID))
                    {
                        previousItemCodes[FormUID] = new Dictionary<int, string>();
                    }

                    // Önceki değeri al veya boş string kullan
                    string previousValue = previousItemCodes[FormUID].ContainsKey(pVal.Row) ? previousItemCodes[FormUID][pVal.Row] : "";
                    string currentValue = itemCode.Value;

                    // Sadece değer gerçekten değiştiyse işlem yap
                    if (currentValue != previousValue)
                    {
                        // Yeni değeri kaydet
                        previousItemCodes[FormUID][pVal.Row] = currentValue;
                        var docDate = (SAPbouiCOM.EditText)oForm.Items.Item("46").Specific;
                        var paymCode = ((SAPbouiCOM.ComboBox)oForm.Items.Item("47").Specific).Value;
                        paymCode = paymCode.Trim();
                        paymCode = paymCode.Replace("-", "N");
                        paymCode = "U_Paym" + paymCode;
                        var price = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific);
                        var discount = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific);
                        // Fiyat sorgusu
                        SAPbobsCOM.Recordset prcRecordSet;

                        prcRecordSet = GetRecordSet(oForm, itemCode.Value);
                        prcRecordSet.MoveFirst();
                        if (!prcRecordSet.EoF)
                        {
                            // Bayrağı ayarla (sonsuz döngüyü engellemek için)
                            isUpdating = true;
                            price.Value = prcRecordSet.Fields.Item("PrcPrice").Value.ToString() + " " + prcRecordSet.Fields.Item("PrcCurrency").Value.ToString();
                            var disc = prcRecordSet.Fields.Item(paymCode).Value.ToString();
                            disc = disc.Replace(',', '.');
                            if (pVal.FormTypeEx == "149")//Sadece satış teklifinde indirim yansımalı
                            {
                                discount.Value = disc;
                            }
                            
                        }
                        for (var i = 0; i < oMatrix.RowCount; i++)
                        {
                            if (!previousItemCodes[FormUID].ContainsKey(i + 1))
                            {
                                price = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(i + 1).Specific);
                                discount = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(i + 1).Specific);
                                var itemCodeTmp = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i + 1).Specific);
                                prcRecordSet = GetRecordSet(oForm, itemCodeTmp.Value);
                                prcRecordSet.MoveFirst();
                                price.Value = prcRecordSet.Fields.Item("PrcPrice").Value.ToString() + " " + prcRecordSet.Fields.Item("PrcCurrency").Value.ToString();
                                var disc = prcRecordSet.Fields.Item(paymCode).Value.ToString();
                                disc = disc.Replace(',', '.');
                                
                                if (pVal.FormTypeEx == "149")//Sadece satış teklifinde indirim yansımalı
                                {
                                    discount.Value = disc;
                                }
                            }

                        }

                    }
                }
                //Choose from listte toplu seçim yapıldığında doğrudan indirim alanı değişirse
                if (pVal.ItemUID == "38" && (pVal.ColUID == "14" || pVal.ColUID == "15") &&
                         pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.BeforeAction)
                {
                    var oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    var itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific);
                    previousItemCodes[FormUID][pVal.Row] = itemCode.Value;
                }

                //Belge Tarihi ya da Ödeme Koşulu değiştiğinde
                if ((pVal.ItemUID == "47" &&
                pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT &&
                !pVal.BeforeAction)
                ||
                ((pVal.ItemUID == "46" || pVal.ItemUID == "10") &&
                pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS &&
                !pVal.BeforeAction))
                {
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
                        foreach (SAPbouiCOM.Item item in oForm.Items)
                        {
                            if (item.Type == SAPbouiCOM.BoFormItemTypes.it_FOLDER)
                            {
                                var folder = (SAPbouiCOM.Folder)item.Specific;
                                folder.Select();
                                break;
                            }
                        }
                    }

                    var oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    oMatrix.Item.Enabled = true;
                    for (var i = 0; i < oMatrix.RowCount; i++)
                    {
                        var itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i + 1).Specific);
                        var price = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(i + 1).Specific);
                        var discount = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(i + 1).Specific);
                        SAPbobsCOM.Recordset prcRecordSet;
                        prcRecordSet = GetRecordSet(oForm, itemCode.Value);
                        prcRecordSet.MoveFirst();
                        if (!prcRecordSet.EoF)
                        {
                            // Bayrağı ayarla (sonsuz döngüyü engellemek için)
                            price.Item.Enabled = true;
                            discount.Item.Enabled = true;
                            price.Value = prcRecordSet.Fields.Item("PrcPrice").Value.ToString() + " " + prcRecordSet.Fields.Item("PrcCurrency").Value.ToString();
                            var disc = prcRecordSet.Fields.Item(paymCode).Value.ToString();
                            disc = disc.Replace(',', '.');
                            if (pVal.FormTypeEx == "149")//Sadece satış teklifinde indirim yansımalı
                            {
                                discount.Value = disc;
                            }
                        }
                    }
                }
                if (pVal.ItemUID == "38" && pVal.ColUID == "15" &&
                    pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.BeforeAction && pVal.FormTypeEx == "149")
                {

                    var oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    var docDate = (SAPbouiCOM.EditText)oForm.Items.Item("46").Specific;
                    var itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific);
                    var paymCode = ((SAPbouiCOM.ComboBox)oForm.Items.Item("47").Specific).Value;
                    var discount = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific);
                    paymCode = paymCode.Trim();
                    paymCode = paymCode.Replace("-", "N");
                    string paymAddCode = "U_AddDiscpaym" + paymCode;
                    paymCode = "U_Paym" + paymCode;
                    SAPbobsCOM.Recordset prcRecordSet;
                    prcRecordSet = GetRecordSet(oForm, itemCode.Value);
                    prcRecordSet.MoveFirst();
                    if (!prcRecordSet.EoF)
                    {
                        float maxDiscount = float.Parse(prcRecordSet.Fields.Item(paymCode).Value.ToString()) + float.Parse(prcRecordSet.Fields.Item(paymAddCode).Value.ToString());
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
            catch
            {
                // SBO_Application.StatusBar.SetText($"Hata: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
