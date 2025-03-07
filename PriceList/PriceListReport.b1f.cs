using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PriceList
{
    [FormAttribute("PriceList.PriceListReport", "PriceListReport.b1f")]
    class PriceListReport : UserFormBase
    {
        private SAPbouiCOM.Button bt_List;
        private SAPbouiCOM.CheckBox chk_Valid;
        private SAPbouiCOM.EditText tx_VDate;
        private SAPbouiCOM.StaticText lb_VDate;
        private SAPbouiCOM.StaticText lb_Type;
        private SAPbouiCOM.ComboBox cb_Type;
        private SAPbouiCOM.Grid gr_List;
        public PriceListReport()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.gr_List = ((SAPbouiCOM.Grid)(this.GetItem("gr_List").Specific));
            this.gr_List.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.gr_List_DoubleClickAfter);
            this.bt_List = ((SAPbouiCOM.Button)(this.GetItem("bt_List").Specific));
            this.bt_List.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.bt_List_ClickBefore);
            this.chk_Valid = ((SAPbouiCOM.CheckBox)(this.GetItem("chk_Valid").Specific));
            this.chk_Valid.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.chk_Valid_PressedAfter);
            this.chk_Valid.ClickAfter += new SAPbouiCOM._ICheckBoxEvents_ClickAfterEventHandler(this.chk_Valid_ClickAfter);
            this.tx_VDate = ((SAPbouiCOM.EditText)(this.GetItem("tx_VDate").Specific));
            this.lb_VDate = ((SAPbouiCOM.StaticText)(this.GetItem("lb_VDate").Specific));
            this.lb_Type = ((SAPbouiCOM.StaticText)(this.GetItem("lb_Type").Specific));
            this.cb_Type = ((SAPbouiCOM.ComboBox)(this.GetItem("cb_Type").Specific));
            this.OnCustomInitialize();

        }


        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        /// 
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

        }

        private void OnCustomInitialize()
        {
            tx_VDate.Value = DateTime.Today.ToString("yyyyMMdd");
            chk_Valid.Checked = true;
            cb_Type.Select("");
        }

        private void bt_List_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                string tableName = "[@SML_PRCHEAD]";
                string type = "Fiyat Listesi";
                if (cb_Type.Value != "")
                {
                    tableName = "[@SML_DSCHEAD]";
                    type = "Dönemsel İndirimler";
                }

                string sqlQuery = $"SELECT DocNum as 'Kayıt No',U_ValidFrom as 'Geçerlilik Başlangıç',U_ValidUntil as 'Geçerlilik Bitiş', U_Description as 'Açıklama','{type}' as 'Tip' FROM {tableName}";
                if (chk_Valid.Checked)
                {
                    sqlQuery += $" WHERE U_ValidFrom <= '{tx_VDate.Value}' AND U_ValidUntil >= '{tx_VDate.Value}'";
                }

                gr_List.DataTable.ExecuteQuery(sqlQuery);
                SAPbobsCOM.Recordset tmpRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                tmpRecordSet.DoQuery(sqlQuery);
                EditTextColumn oColumns = (EditTextColumn)gr_List.Columns.Item("Kayıt No");
                oColumns.LinkedObjectType = "UDOSML_PRCHEAD";

                // Satırların renklendirilmesi
                DateTime today = DateTime.Today;
                int rowCount = gr_List.DataTable.Rows.Count;
                if (tmpRecordSet.RecordCount > 0)
                {
                    for (int i = 0; i < rowCount; i++)
                    {
                        string validUntilStr = gr_List.DataTable.GetValue("Geçerlilik Bitiş", i).ToString();
                        if (DateTime.TryParse(validUntilStr, out DateTime validUntil))
                        {
                            int diffDays = (validUntil - today).Days;

                            if (diffDays >= 2)
                            {
                                gr_List.CommonSetting.SetRowBackColor(i + 1, Program.colors["green"]); // Yeşil
                            }
                            else if (diffDays >= 0)
                            {
                                gr_List.CommonSetting.SetRowBackColor(i + 1, Program.colors["yellow"]); // Sarı
                            }
                            else
                            {
                                gr_List.CommonSetting.SetRowBackColor(i + 1, Program.colors["red"]); // Kırmızı
                            }
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                Program.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
            }
        }


        private void chk_Valid_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

        }

        private void chk_Valid_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            tx_VDate.Item.Visible = chk_Valid.Checked;
            lb_VDate.Item.Visible = chk_Valid.Checked;

        }

        private void gr_List_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
        }

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            bt_List.Item.Click();

        }
    }
}
