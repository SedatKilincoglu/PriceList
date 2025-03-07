using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using SAPbouiCOM;
using EROPASAPLib;
using System.Linq;
using System.Globalization;

namespace PriceList
{
    [FormAttribute("PriceList.PriceListForm", "PriceListForm.b1f")]

    class Form1 : UserFormBase
    {
        public string BaseFormUID { get; set; }
        public string linkDocEntry { get; set; }

        private Matrix mx_Price;
        private StaticText lb_desc;
        private EditText EditText0;
        private StaticText lb_Vf;
        private StaticText lb_Vu;
        private EditText tx_VF;
        private EditText tx_VU;
        private Button Button0;
        private Button Button1;
        private Button bt_Add;
        private EditText tx_Doc;
        private StaticText lb_Doc;
        private Button bt_Del;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public Form1(String DocEntry = "")
        {
            linkDocEntry = DocEntry;
        }

        public override void OnInitializeComponent()
        {
            this.mx_Price = ((Matrix)(this.GetItem("mx_Price").Specific));
            this.lb_desc = ((StaticText)(this.GetItem("lb_desc").Specific));
            this.EditText0 = ((EditText)(this.GetItem("Item_1").Specific));
            this.lb_Vf = ((StaticText)(this.GetItem("lb_Vf").Specific));
            this.lb_Vu = ((StaticText)(this.GetItem("lb_Vu").Specific));
            this.lb_Vu.ClickAfter += new _IStaticTextEvents_ClickAfterEventHandler(this.lb_Vu_ClickAfter);
            this.tx_VF = ((EditText)(this.GetItem("tx_VF").Specific));
            this.tx_VU = ((EditText)(this.GetItem("tx_VU").Specific));
            this.Button0 = ((Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new _IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((Button)(this.GetItem("2").Specific));
            this.bt_Add = ((Button)(this.GetItem("bt_Add").Specific));
            this.bt_Add.ClickBefore += new _IButtonEvents_ClickBeforeEventHandler(this.bt_Add_ClickBefore);
            this.tx_Doc = ((EditText)(this.GetItem("tx_Doc").Specific));
            this.lb_Doc = ((StaticText)(this.GetItem("lb_Doc").Specific));
            this.bt_Del = ((Button)(this.GetItem("bt_Del").Specific));
            this.bt_Del.ClickBefore += new _IButtonEvents_ClickBeforeEventHandler(this.bt_Del_ClickBefore);
            this.OnCustomInitialize();

        }
        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            //    Program.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataDeleteBefore += new SAPbouiCOM.Framework.FormBase.DataDeleteBeforeHandler(this.Form_DataDeleteBefore);


        }

        private void OnCustomInitialize()
        {

        }

        private void setStatusBarText(string Message, string Type)
        {
            BoStatusBarMessageType messageType;
            messageType = BoStatusBarMessageType.smt_None;
            if (Type == "error")
            {
                messageType = BoStatusBarMessageType.smt_Error;
            }
            if (Type == "warning")
            {
                messageType = BoStatusBarMessageType.smt_Warning;
            }
            if (Type == "success")
            {
                messageType = BoStatusBarMessageType.smt_Success;
            }
            Program.SBO_Application.StatusBar.SetText(Message, BoMessageTime.bmt_Short, messageType);
        }

        private void Button0_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            if (pVal.FormMode == 0)
            {
                BubbleEvent = true;
                return;
            }
            if (!Program.AuthorizedUser)
            {
                Program.SBO_Application.MessageBox("Yetkiniz Yok");
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
                    EditText priceCell = (EditText)mx_Price.Columns.Item("Price").Cells.Item(i + 1).Specific;
                    ComboBox currCell = (ComboBox)mx_Price.Columns.Item("Currency").Cells.Item(i + 1).Specific;
                    if (priceCell.Value == null || currCell.Value == "")
                    {
                        setStatusBarText("Lütfen Para birimlerini doldurun", "error");
                        MainSAPForm.Freeze(false);
                        return;
                    }
                }

            }
            catch (Exception e)
            {
                Program.SBO_Application.StatusBar.SetText("Hata: " + e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                MainSAPForm.Freeze(false);
                return;
            }
            MainSAPForm.Freeze(false);
            BubbleEvent = true;

        }

        SAPbouiCOM.Form MainSAPForm;
        private void bt_Add_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            if (!Program.AuthorizedUser)
            {
                Program.SBO_Application.MessageBox("Yetkiniz Yok");
                BubbleEvent = false;
                return;
            }
            MainSAPForm = Program.SBO_Application.Forms.ActiveForm;
            var frmFilter = new FilterForm(Program.SBO_Application, PassFilterData);
            frmFilter.Show();
            BubbleEvent = true;
        }


        private void PassFilterData(Models.ItemSet[] selectedMaterials)
        {
            bool isNewRowAdded = false;
            MainSAPForm.Freeze(true);
            Models.ItemSet[] matrixItems = new Models.ItemSet[] { };
            for (var i = 0; i < mx_Price.RowCount; i++)
            {
                Models.ItemSet matrixItem = new Models.ItemSet
                {
                    itemCode = ((EditText)mx_Price.Columns.Item("ItemCode").Cells.Item(i + 1).Specific).Value
                };
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
                EditText itemCodeCell = (EditText)mx_Price.Columns.Item("ItemCode").Cells.Item(mx_Price.VisualRowCount).Specific;
                EditText itemNameCell = (EditText)mx_Price.Columns.Item("ItemName").Cells.Item(mx_Price.VisualRowCount).Specific;
                EditText LineIdCell = (EditText)mx_Price.Columns.Item("#").Cells.Item(mx_Price.VisualRowCount).Specific;
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
            try
            {
                DateTime today = DateTime.Today;
                int rowCount = mx_Price.RowCount;
                int rowcolor;
                string validUntilStr = tx_VU.Value;
                DateTime.TryParseExact(tx_VU.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime validUntil);
                int diffDays = (validUntil - today).Days;

                if (diffDays >= 2)
                {
                    rowcolor = Program.colors["green"];
                }
                else if (diffDays >= 0)
                {
                    rowcolor = Program.colors["yellow"];
                }
                else
                {
                    rowcolor = Program.colors["red"];
                }

                for (int i = 1; i <= rowCount; i++)
                {
                    mx_Price.CommonSetting.SetRowBackColor(i, rowcolor);

                }
            }
            catch
            {

            }

        }

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            mx_Price.AutoResizeColumns();
            if (linkDocEntry != "")
            {
                SAPbouiCOM.Form prcForm;
                prcForm = Program.SBO_Application.Forms.Item(pVal.FormUID);
                prcForm.Mode = BoFormMode.fm_FIND_MODE;
                tx_Doc.Item.Enabled = true;
                tx_Doc.Value = linkDocEntry;
                Button0.Item.Click();
                linkDocEntry = "";
            }
        }


        private void bt_Del_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            if (!Program.AuthorizedUser)
            {
                Program.SBO_Application.MessageBox("Yetkiniz Yok");
                BubbleEvent = false;
                return;
            }
            BubbleEvent = true;
            // Seçili satırı al ve sil
            int selectedRow = mx_Price.GetNextSelectedRow(0, BoOrderType.ot_RowOrder);

            if (selectedRow > 0)
            {
                DeleteMatrixRow(mx_Price, selectedRow);
            }
            else
            {
                setStatusBarText("Silinecek satır seçilmedi.", "warning");
            }

        }

        private void DeleteMatrixRow(Matrix matrix, int rowIndex)
        {
            try
            {
                // Matrixten satırı kaldır
                matrix.DeleteRow(rowIndex);

                matrix.FlushToDataSource();
                // Formu güncelle
                Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

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


        private void Form_LoadAfter(SBOItemEventArg pVal)
        {

        }

        private void Form_DataDeleteBefore(ref BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (!Program.AuthorizedUser)
            {
                Program.SBO_Application.MessageBox("Silme Yetkiniz Yok");
                BubbleEvent = false;
                return;
            }


        }

        private void lb_Vu_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {



        }

    }
}