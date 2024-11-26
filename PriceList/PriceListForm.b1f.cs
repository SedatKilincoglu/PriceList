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
    [FormAttribute("PriceList.PriceListForm", "PriceListForm.b1f")]

    class Form1 : UserFormBase
    {
        public string BaseFormUID { get; set; }

        public Form1()
        { 
        }
        private SAPbouiCOM.Matrix mx_Price;
        private SAPbouiCOM.StaticText lb_desc;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText lb_Vf;
        private SAPbouiCOM.StaticText lb_Vu;
        private SAPbouiCOM.EditText tx_VF;
        private SAPbouiCOM.EditText tx_VU;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button bt_Add;
        private SAPbouiCOM.EditText tx_Doc;
        private SAPbouiCOM.StaticText lb_Doc;
        private Button bt_Del;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.mx_Price = ((SAPbouiCOM.Matrix)(this.GetItem("mx_Price").Specific));
            this.lb_desc = ((SAPbouiCOM.StaticText)(this.GetItem("lb_desc").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
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
            this.bt_Del = ((SAPbouiCOM.Button)(this.GetItem("bt_Del").Specific));
            this.bt_Del.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.bt_Del_ClickBefore);
            this.OnCustomInitialize();

        }
        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            //  Program.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private void OnCustomInitialize()
        {

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
            if (!Program.AuthorizedUser)
            {
                BubbleEvent = false;
                return;
            }
            BubbleEvent = false;
            SAPbouiCOM.Form MainSAPForm;
            MainSAPForm = Program.SBO_Application.Forms.ActiveForm;
            if (MainSAPForm.Mode == BoFormMode.fm_OK_MODE)
            {
                BubbleEvent = true;
                return;
            }

            MainSAPForm.Freeze(true);
            try
            {
                if (tx_VF.Value == "" || tx_VU.Value == "")
                {
                    setStatusBarText("Lütfen Geçerlilik Tarihi alanlarını doldurun", "error");
                    MainSAPForm.Freeze(false);
                    return;
                }
                if (int.Parse(tx_VF.Value) > int.Parse(tx_VU.Value))
                {
                    setStatusBarText("Geçerlilik Başlangıcı, Geçerlilik Bitişinden küçük olmalıdır", "error");
                    MainSAPForm.Freeze(false);
                    return;
                }
                if (mx_Price.RowCount == 0)
                {
                    setStatusBarText("Lütfen Kalemleri girin", "error");
                    MainSAPForm.Freeze(false);
                    return;
                }
                for (var i = 0; i < mx_Price.RowCount; i++)
                {
                    SAPbouiCOM.EditText priceCell = (SAPbouiCOM.EditText)mx_Price.Columns.Item("Price").Cells.Item(i + 1).Specific;
                    SAPbouiCOM.ComboBox currCell = (SAPbouiCOM.ComboBox)mx_Price.Columns.Item("Currency").Cells.Item(i + 1).Specific;
                    if (float.Parse(priceCell.Value) <= 0 || priceCell.Value == null || currCell.Value == "")
                    {
                        setStatusBarText("Lütfen Fiyatları ve para birimlerini doldurun", "error");
                        MainSAPForm.Freeze(false);
                        return;
                    }
                }

            }
            catch(Exception e)
            { 
                Program.SBO_Application.StatusBar.SetText("Hata: " + e.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                MainSAPForm.Freeze(false);
                return;
            }
            MainSAPForm.Freeze(false);
            BubbleEvent = true;

        }

        SAPbouiCOM.Form MainSAPForm;
        private void bt_Add_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            if (!Program.AuthorizedUser)
            {
                BubbleEvent = false;
                return;
            }
            MainSAPForm = Program.SBO_Application.Forms.ActiveForm;
            var frmFilter = new FilterForm(Program.SBO_Application, PassFilterData);
            frmFilter.Show();
            BubbleEvent = true;
        } 

        private bool isExistonMatrix(SAPbouiCOM.Matrix matrix, string ColName, string FindingValue) {

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
            for (var i = 0; i < mx_Price.RowCount; i++)
            {
                Models.ItemSet matrixItem = new Models.ItemSet();
                matrixItem.itemCode = ((SAPbouiCOM.EditText)mx_Price.Columns.Item("ItemCode").Cells.Item(i + 1).Specific).Value;
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
                mx_Price.AddRow();
                mx_Price.ClearRowData(mx_Price.RowCount);
                SAPbouiCOM.EditText itemCodeCell = (SAPbouiCOM.EditText)mx_Price.Columns.Item("ItemCode").Cells.Item(mx_Price.VisualRowCount).Specific;
                SAPbouiCOM.EditText itemNameCell = (SAPbouiCOM.EditText)mx_Price.Columns.Item("ItemName").Cells.Item(mx_Price.VisualRowCount).Specific;
                SAPbouiCOM.EditText LineIdCell = (SAPbouiCOM.EditText)mx_Price.Columns.Item("#").Cells.Item(mx_Price.VisualRowCount).Specific;
                itemCodeCell.Value = selectedMaterials[i].itemCode;
                itemNameCell.Value = selectedMaterials[i].itemName;
                LineIdCell.Value = null;
                isNewRowAdded = true;
            }

            if (MainSAPForm.Mode != BoFormMode.fm_ADD_MODE && isNewRowAdded)
            {
                MainSAPForm.Mode = BoFormMode.fm_UPDATE_MODE;
            }
            mx_Price.FlushToDataSource();
            MainSAPForm.Freeze(false);
            mx_Price.AutoResizeColumns();

        }


        private void Form_DataLoadAfter(ref BusinessObjectInfo pVal)
        {
            mx_Price.AutoResizeColumns();

        }

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
             mx_Price.AutoResizeColumns();
        }

        

        private void bt_Del_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            if (!Program.AuthorizedUser)
            {
                BubbleEvent = false;
                return;
            }
            BubbleEvent = true;
            // Seçili satırı al ve sil
            int selectedRow = mx_Price.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

            if (selectedRow > 0)
            {
                DeleteMatrixRow(mx_Price, selectedRow);
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
                // Matrixten satırı kaldır
                matrix.DeleteRow(rowIndex);

                matrix.FlushToDataSource();
                // Formu güncelle
                SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

                if (form == null)
                {
                    setStatusBarText("Aktif form alınamadı.", "error");
                    return;
                }
                if (form.Mode != BoFormMode.fm_ADD_MODE)
                {
                    form.Mode = BoFormMode.fm_UPDATE_MODE;
                }

            }
            catch (Exception ex)
            {
                setStatusBarText($"Satır silinirken bir hata oluştu: {ex.Message}", "error");
            }
        }


        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            Program.SBO_Application.StatusBar.SetText(pVal.EventType.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
        }

        private void Form_LoadAfter(SBOItemEventArg pVal)
        {

            
        }
    }
}