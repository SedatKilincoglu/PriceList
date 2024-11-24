using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using SAPbouiCOM;
using EROPASAPLib;
using System.Linq;

namespace PriceList
{
    [FormAttribute("PriceList.DiscountForm", "DiscountForm.b1f")]

    class DiscountForm : UserFormBase
    {
        public string BaseFormUID { get; set; }

        public DiscountForm()
        { 
        }
        private SAPbouiCOM.Matrix mx_Disc;
        private SAPbouiCOM.StaticText lb_desc;
        private SAPbouiCOM.EditText tx_Desc;
        private SAPbouiCOM.StaticText lb_Vf;
        private SAPbouiCOM.StaticText lb_Vu;
        private SAPbouiCOM.EditText tx_VF;
        private SAPbouiCOM.EditText tx_VU;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button bt_Add;
        private SAPbouiCOM.EditText tx_Doc;
        private SAPbouiCOM.StaticText lb_Doc;
        private SAPbobsCOM.Recordset PaymentRecordSet;
        private SAPbouiCOM.Button bt_Del;
        private SAPbouiCOM.ComboBox cm_paym;
        private SAPbouiCOM.EditText tx_ch;
        private SAPbouiCOM.Button bt_ChAll;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.lb_desc = ((SAPbouiCOM.StaticText)(this.GetItem("lb_desc").Specific));
            this.tx_Desc = ((SAPbouiCOM.EditText)(this.GetItem("tx_Desc").Specific));
            this.lb_Vf = ((SAPbouiCOM.StaticText)(this.GetItem("lb_Vf").Specific));
            this.lb_Vu = ((SAPbouiCOM.StaticText)(this.GetItem("lb_Vu").Specific));
            this.tx_VF = ((SAPbouiCOM.EditText)(this.GetItem("tx_VF").Specific));
            this.tx_VU = ((SAPbouiCOM.EditText)(this.GetItem("tx_VU").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.bt_Add = ((SAPbouiCOM.Button)(this.GetItem("bt_Add").Specific));
            this.bt_Add.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.bt_Add_ClickBefore);
            this.tx_Doc = ((SAPbouiCOM.EditText)(this.GetItem("tx_Doc").Specific));
            this.lb_Doc = ((SAPbouiCOM.StaticText)(this.GetItem("lb_Doc").Specific));
            this.mx_Disc = ((SAPbouiCOM.Matrix)(this.GetItem("mx_Disc").Specific));
            this.bt_Del = ((SAPbouiCOM.Button)(this.GetItem("bt_Del").Specific));
            this.bt_Del.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.bt_Del_ClickBefore);
            this.cm_paym = ((SAPbouiCOM.ComboBox)(this.GetItem("cm_paym").Specific));
            this.tx_ch = ((SAPbouiCOM.EditText)(this.GetItem("tx_ch").Specific));
            this.bt_ChAll = ((SAPbouiCOM.Button)(this.GetItem("bt_ChAll").Specific));
            this.bt_ChAll.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.bt_ChAll_ClickBefore);
            this.OnCustomInitialize();

        }
        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.UnloadAfter += new UnloadAfterHandler(this.Form_UnloadAfter);
            Program.SBO_Application.FormDataEvent += this.SBO_Application_FormDataEvent;
        }

        private void OnCustomInitialize()
        {
            PaymentRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            AddContextMenu();
            AddMatrixColumns();
        }

        private void setMatrixData(SAPbouiCOM.Matrix mtName, String colName, int rowIndex, String Data)
        {
            var tmpval = (SAPbouiCOM.EditText)mtName.Columns.Item(colName).Cells.Item(rowIndex).Specific;
            tmpval.Value = Data;
        }

        private void setStatusBarText(string Message, string Type)
        {
            SAPbouiCOM.BoStatusBarMessageType messageType;
            messageType = SAPbouiCOM.BoStatusBarMessageType.smt_None;
            if (Type == "error")
            {
                messageType = SAPbouiCOM.BoStatusBarMessageType.smt_Error;
            }
            if (Type == "warning")
            {
                messageType = SAPbouiCOM.BoStatusBarMessageType.smt_Warning;
            }
            if (Type == "success")
            {
                messageType = SAPbouiCOM.BoStatusBarMessageType.smt_Success;
            }
            Program.SBO_Application.StatusBar.SetText(Message, SAPbouiCOM.BoMessageTime.bmt_Short, messageType);
        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = false;
            MainSAPForm = Program.SBO_Application.Forms.ActiveForm;
            if (MainSAPForm.Mode == BoFormMode.fm_OK_MODE)
            {
                BubbleEvent = true;
                return;
            }
            try
            {
                if (tx_VF.Value == "" || tx_VU.Value == "")
                {
                    setStatusBarText("Lütfen Geçerlilik Tarihi alanlarını doldurun", "error");
                    return;
                }
                if (int.Parse(tx_VF.Value) > int.Parse(tx_VU.Value))
                {
                    setStatusBarText("Geçerlilik Başlangıcı, Geçerlilik Bitişinden küçük olmalıdır", "error");
                    return;
                }
                if (mx_Disc.RowCount == 0)
                {
                    setStatusBarText("Lütfen Kalemleri girin", "error");
                    return;
                }
                setStatusBarText("Silinen satırlar ayıklandı", "none");
            }
            catch(Exception e)
            { 
                Program.SBO_Application.StatusBar.SetText("Hata: " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            BubbleEvent = true;

        }

        SAPbouiCOM.Form MainSAPForm;
        private void bt_Add_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            MainSAPForm = Program.SBO_Application.Forms.ActiveForm;
            var frmFilter = new FilterForm(Program.SBO_Application, PassFilterData);
            frmFilter.Show();
            BubbleEvent = true;
        }

        private bool isExistonMatrix(SAPbouiCOM.Matrix matrix, string ColName, string FindingValue)
        {

            for (var ix = 0; ix < matrix.RowCount; ix++)
            {
                SAPbouiCOM.EditText findingCell = (SAPbouiCOM.EditText)matrix.Columns.Item(ColName).Cells.Item(ix + 1).Specific;
                if (findingCell.Value == FindingValue)
                {
                    return true;
                }
            }
            return false;
        }

        private void PassFilterData(Models.ItemSet[] selectedMaterials)
        {
            bool isNewRowAdded = false;
            MainSAPForm.Freeze(true);
            Models.ItemSet[] matrixItems = new Models.ItemSet[] { };
            for (var i = 0; i < mx_Disc.RowCount; i++)
            {
                Models.ItemSet matrixItem = new Models.ItemSet();
                matrixItem.itemCode = ((SAPbouiCOM.EditText)mx_Disc.Columns.Item("ItemCode").Cells.Item(i + 1).Specific).Value;
                matrixItems = matrixItems.Append(matrixItem).ToArray();
            }
            for (var i = 0; i < selectedMaterials.Length; i++)
            {
                bool isFound = false;
                for (var j = 0; j < matrixItems.Length; j++)
                {
                    if (selectedMaterials[i].itemCode == matrixItems[j].itemCode)
                    {
                        isFound = true;
                        break;
                    }
                }
                if (isFound)
                {
                    continue;
                }
                mx_Disc.AddRow();
                mx_Disc.ClearRowData(mx_Disc.RowCount);
                
                SAPbouiCOM.EditText itemCodeCell = (SAPbouiCOM.EditText)mx_Disc.Columns.Item("ItemCode").Cells.Item(mx_Disc.VisualRowCount).Specific;
                SAPbouiCOM.EditText itemNameCell = (SAPbouiCOM.EditText)mx_Disc.Columns.Item("ItemName").Cells.Item(mx_Disc.VisualRowCount).Specific;
                SAPbouiCOM.EditText LineIdCell = (SAPbouiCOM.EditText)mx_Disc.Columns.Item("#").Cells.Item(mx_Disc.VisualRowCount).Specific;
                itemCodeCell.Value = selectedMaterials[i].itemCode;
                itemNameCell.Value = selectedMaterials[i].itemName;
                LineIdCell.Value = null;
                isNewRowAdded = true;
            }

            if (MainSAPForm.Mode != BoFormMode.fm_ADD_MODE && isNewRowAdded)
            {
                MainSAPForm.Mode = BoFormMode.fm_UPDATE_MODE;
            }
            mx_Disc.FlushToDataSource();
            MainSAPForm.Freeze(false);
            mx_Disc.AutoResizeColumns();

        }

        private void AddContextMenu()
        {
            // Yeni bir sağ tıklama menüsü ekle
            SAPbouiCOM.Menus menus = Program.SBO_Application.Menus;
            SAPbouiCOM.MenuItem menuItem = Program.SBO_Application.Menus.Item("1280"); // Sağ tıklama menüsünün ID'si

            if (!menus.Exists("Custom_DeleteRow"))
            {
                SAPbouiCOM.MenuCreationParams menuParams = (SAPbouiCOM.MenuCreationParams)Program.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                menuParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                menuParams.UniqueID = "Custom_DeleteRow";
                menuParams.String = "Satırı Sil";
                menuParams.Enabled = true;
                menuParams.Position = menuItem.SubMenus.Count + 1;
                menuItem.SubMenus.AddEx(menuParams);
            }
        }

        private void AddMatrixColumns()
        {
            SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.ActiveForm;
            SAPbouiCOM.Columns oColumns = mx_Disc.Columns;
            cm_paym.ValidValues.Add("E", "Ek İndirim Yetkisi");
            // Sorguyu çalıştır

            foreach (Models.Paym paym in Program.paymList)
            {
                {
                    string paymentCode = paym.paymCode;
                    string paymentName = paym.paymName;

                    // Matrix'e dinamik kolon ekle
                    string colId = "col" + paymentCode;
                    if (!ColumnExists(oColumns, colId)) // Eğer kolon zaten yoksa
                    {
                        SAPbouiCOM.Column oColumn = oColumns.Add(colId, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.RightJustified = true;
                        oColumn.TitleObject.Caption = $"İndirim({paymentName})";
                        oColumn.DataBind.SetBound(true, "@SML_DSCITEM", "U_" + paymentCode); // UserDataSource ile bağla
                        oColumn.Editable = true;
                    }

                    cm_paym.ValidValues.Add(paymentCode, paymentName);
                }
            }
        }

        // Kolonun varlığını kontrol eden yardımcı fonksiyon
        private bool ColumnExists(SAPbouiCOM.Columns oColumns, string colId)
        {
            try
            {
                var column = oColumns.Item(colId);
                return column != null;
            }
            catch
            {
                return false;
            }
        }

        private void Form_DataLoadAfter(ref BusinessObjectInfo pVal)
        {
            mx_Disc.AutoResizeColumns();
        }

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            mx_Disc.AutoResizeColumns();
        }

        private void Form_UnloadAfter(SBOItemEventArg pVal)
        {
            
        }
        

        public void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if ((businessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD ||
                businessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                ) &&
                !businessObjectInfo.BeforeAction && 
                businessObjectInfo.ActionSuccess) 
            {
                // Tetiklenen formu al
                SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(businessObjectInfo.FormUID);

                // Formu kontrol et
                if (oForm.TypeEx == "PriceList.DiscountForm") // Form tipini kontrol et
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                }
            }
        }
         
        

        private void bt_Del_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            MainSAPForm = Program.SBO_Application.Forms.ActiveForm;
            int selectedRow = mx_Disc.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

            if (selectedRow > 0)
            {
                DeleteMatrixRow(mx_Disc, selectedRow);
            }
            else
            {
                setStatusBarText("Silinecek satır seçilmedi.", "warning");
            }

        }

        private void DeleteMatrixRow(SAPbouiCOM.Matrix matrix, int rowIndex)
        {
            try
            {
                string ItemCode = ((SAPbouiCOM.EditText)matrix.Columns.Item("ItemCode").Cells.Item(rowIndex).Specific).Value;
                matrix.DeleteRow(rowIndex);
                matrix.FlushToDataSource();

                if (MainSAPForm == null)
                {
                    setStatusBarText("Aktif form alınamadı.", "error");
                    return;
                }
                if (MainSAPForm.Mode != BoFormMode.fm_ADD_MODE)
                {
                    MainSAPForm.Mode = BoFormMode.fm_UPDATE_MODE;
                }

            }
            catch (Exception ex)
            {
                setStatusBarText($"Satır silinirken bir hata oluştu: {ex.Message}", "error");
            }
        }

        private void Button2_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }

        private void bt_ChAll_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            string colId;
            if (cm_paym.Value == "" || float.Parse(tx_ch.Value) < 0)
            {
                return;
            }
            if (cm_paym.Value == "E")
            {
                colId = "A_Disc";
            }
            else
            {
                colId = "col" + cm_paym.Value;
            }
            
            MainSAPForm = Program.SBO_Application.Forms.ActiveForm;
            MainSAPForm.Freeze(true);
            for (int i = 0; i < mx_Disc.RowCount; i++)
            {
                SAPbouiCOM.EditText DiscountCol = (SAPbouiCOM.EditText)mx_Disc.Columns.Item(colId).Cells.Item(i + 1).Specific;
                DiscountCol.Value = tx_ch.Value;
            }
            mx_Disc.FlushToDataSource();
            tx_ch.Value = "0";
            MainSAPForm.Freeze(false);

        }
    }
}