using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace PriceList
{
    [FormAttribute("PriceList.FilterForm", "FilterForm.b1f")]
    class FilterForm : UserFormBase
    {
        private SAPbouiCOM.Application SBO_Application;
        
        private Action<Models.ItemSet[]> Callback;
        public FilterForm(SAPbouiCOM.Application app, Action<Models.ItemSet[]> callback)
        {
            SBO_Application = app;
            Callback = callback;
        }

        private SAPbouiCOM.EditText tx_IcUntil;
        private SAPbouiCOM.EditText tx_IcFrom;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText tx_InFrom;
        private SAPbouiCOM.EditText tx_InUntil;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.ComboBox cm_Grp;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.ComboBox cm_sGrp;
        private SAPbouiCOM.Button bt_Find;
        private SAPbouiCOM.Grid gr_List;
        private SAPbouiCOM.Button bt_Check;
        private SAPbouiCOM.Button bt_UnCheck;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.DataTable dt_OITM;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.ComboBox cb_IcTyp;
        private SAPbouiCOM.ComboBox cb_InTyp;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.ComboBox cm_Fc;
        private SAPbobsCOM.Recordset ItemGroupRecordSet;
        private SAPbobsCOM.Recordset ItemSubGroupRecordSet;
        private SAPbobsCOM.Recordset FirmCodeRecordSet;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.tx_IcUntil = ((SAPbouiCOM.EditText)(this.GetItem("tx_IcUntil").Specific));
            this.tx_IcFrom = ((SAPbouiCOM.EditText)(this.GetItem("tx_IcFrom").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.tx_InFrom = ((SAPbouiCOM.EditText)(this.GetItem("tx_InFrom").Specific));
            this.tx_InUntil = ((SAPbouiCOM.EditText)(this.GetItem("tx_InUntil").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.cm_Grp = ((SAPbouiCOM.ComboBox)(this.GetItem("cm_Grp").Specific));
            this.cm_Grp.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.cm_Grp_ComboSelectAfter);
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.cm_sGrp = ((SAPbouiCOM.ComboBox)(this.GetItem("cm_sGrp").Specific));
            this.bt_Find = ((SAPbouiCOM.Button)(this.GetItem("bt_Find").Specific));
            this.bt_Find.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.bt_Find_ClickBefore);
            this.gr_List = ((SAPbouiCOM.Grid)(this.GetItem("gr_List").Specific));
            this.bt_Check = ((SAPbouiCOM.Button)(this.GetItem("bt_Check").Specific));
            this.bt_Check.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.bt_Check_ClickBefore);
            this.bt_UnCheck = ((SAPbouiCOM.Button)(this.GetItem("bt_UnCheck").Specific));
            this.bt_UnCheck.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.bt_UnCheck_ClickBefore);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.cb_IcTyp = ((SAPbouiCOM.ComboBox)(this.GetItem("cb_IcTyp").Specific));
            this.cb_IcTyp.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.cb_IcTyp_ComboSelectAfter);
            this.cb_InTyp = ((SAPbouiCOM.ComboBox)(this.GetItem("cb_InTyp").Specific));
            this.cb_InTyp.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.cb_InTyp_ComboSelectAfter);
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.cm_Fc = ((SAPbouiCOM.ComboBox)(this.GetItem("cm_Fc").Specific));

            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {
            ItemGroupRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            ItemSubGroupRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            FirmCodeRecordSet = (SAPbobsCOM.Recordset)Program.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            cm_Grp.ValidValues.Add("", "Seçiniz");
            cm_sGrp.ValidValues.Add("", "Seçiniz");
            cm_Fc.ValidValues.Add("", "Seçiniz");

            string strQuery = "SELECT ItmsGrpCod AS Code,ItmsGrpNam AS Name FROM OITB";
            ItemGroupRecordSet.DoQuery(strQuery);
            ItemGroupRecordSet.MoveFirst();
            while (!ItemGroupRecordSet.EoF)
            {
                cm_Grp.ValidValues.Add(ItemGroupRecordSet.Fields.Item("Code").Value.ToString(), ItemGroupRecordSet.Fields.Item("Name").Value.ToString());
                ItemGroupRecordSet.MoveNext();
            }
            strQuery = "SELECT * FROM [@ALT_KALEM_GRUBU]";
            ItemSubGroupRecordSet.DoQuery(strQuery);

            strQuery = "SELECT FirmCode AS Code,FirmName AS Name FROM OMRC";
            FirmCodeRecordSet.DoQuery(strQuery);
            FirmCodeRecordSet.MoveFirst();
            while (!FirmCodeRecordSet.EoF)
            {
                cm_Fc.ValidValues.Add(FirmCodeRecordSet.Fields.Item("Code").Value.ToString(), FirmCodeRecordSet.Fields.Item("Name").Value.ToString());
                FirmCodeRecordSet.MoveNext();
            }
            tx_IcUntil.Item.Enabled = false;
            tx_InUntil.Item.Enabled = false;
        }



        private void OrgVisibles(int comp)
        {
            SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.ActiveForm;
            if (comp == 0)
            {
                if (cb_IcTyp.Value == "3")
                {
                    tx_IcUntil.Item.Enabled = true;
                }
                else
                {
                    tx_IcUntil.Value = "";
                    oForm.Items.Item("tx_IcFrom").Click();
                    tx_IcUntil.Item.Enabled = false;
                }
            }
            if (comp == 1)
            {
                if (cb_InTyp.Value == "3")
                {
                    tx_InUntil.Item.Enabled = true;
                }
                else
                {
                    tx_InUntil.Value = "";
                    oForm.Items.Item("tx_InFrom").Click();
                    tx_InUntil.Item.Enabled = false;
                }
            }


        }


        int validvaluecounter = 0;
        private string GetOperator()
        {
            string opr;
            validvaluecounter ++;
            if (validvaluecounter == 1)
                opr = " WHERE ";
            else
                opr = " AND ";
            return opr;
        }

        private string GetLogic(string colName, string value, string opr)
        {
            string retval = colName + " ";
            if (opr == "")
            {
                retval = retval + "LIKE '%" + value + "%'";
            }
            if (opr == "1")
            {
                retval = retval + "LIKE '" + value + "%'";
            }
            if (opr == "2")
            {
                retval = retval + "LIKE '%" + value + "'";
            }
            if (opr == "3")
            {
                retval = retval + ">= '" + value + "'";
            }
            return retval;
        }

        private string prepareSQL()
        {
            validvaluecounter = 0;
            string SqlString;
            SqlString = "SELECT 'N' as Chk, ItemCode,ItemName FROM OITM ";
            if (tx_IcFrom.Value != "")
            {
                SqlString = SqlString + GetOperator() + GetLogic("ItemCode", tx_IcFrom.Value, cb_IcTyp.Value);
            }
            if (tx_IcUntil.Value != "" && cb_IcTyp.Value == "3")
            {
                SqlString = SqlString + GetOperator() + $"ItemCode <= '{tx_IcUntil.Value}'";
            }
            if (tx_InFrom.Value != "")
            {
                SqlString = SqlString + GetOperator() + GetLogic("ItemName", tx_InFrom.Value, cb_InTyp.Value);
            }
            if (tx_InUntil.Value != "" && cb_InTyp.Value == "3")
            {
                SqlString = SqlString + GetOperator() + $"ItemName <= '{tx_InUntil.Value}'";
            }
            if (cm_Grp.Value != "")
            {
                SqlString = SqlString + GetOperator() + $"ItmsGrpCod = '{cm_Grp.Value}'";
            }
            if (cm_sGrp.Value != "")
            {
                SqlString = SqlString + GetOperator() + $"U_SUBTYPE = '{cm_Grp.Value}'";
            }
            if (cm_Fc.Value != "")
            {
                SqlString = SqlString + GetOperator() + $"FirmCode = '{cm_Fc.Value}'";
            }
            SqlString = SqlString + " ORDER BY ItemCode";
            return SqlString;
        }


        private void bt_Find_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            dt_OITM = gr_List.DataTable;
            dt_OITM.ExecuteQuery(prepareSQL());
            gr_List.Columns.Item("Chk").TitleObject.Caption = "Seç";
            gr_List.Columns.Item("Chk").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            gr_List.Columns.Item("ItemCode").TitleObject.Caption = "Kalem";
            gr_List.Columns.Item("ItemName").TitleObject.Caption = "Kalem Tanıtıcı";
            gr_List.AutoResizeColumns();
        }

        private void allCheck(string isChecked)
        {
            SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.ActiveForm;
            oForm.Freeze(true);
            for (var i = 0; i < gr_List.DataTable.Rows.Count; i++)
            {
                gr_List.DataTable.Columns.Item("Chk").Cells.Item(i).Value = isChecked;
            }
            oForm.Freeze(false);
        }


        private void bt_Check_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            allCheck("Y");
        }

        private void bt_UnCheck_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            allCheck("N");

        }

        private void cb_IcTyp_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

            OrgVisibles(0);
        }

        private void cb_InTyp_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            OrgVisibles(1);
        }

        private void cm_Grp_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.ActiveForm;
            try
            {
                int Count = cm_sGrp.ValidValues.Count;
                cm_sGrp.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                while (cm_sGrp.ValidValues.Count > 1)
                {
                    if (cm_sGrp.ValidValues.Item(cm_sGrp.ValidValues.Count - 1).Value != "")
                    {
                        cm_sGrp.ValidValues.Remove(cm_sGrp.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                    } 
                }
                ItemSubGroupRecordSet.MoveFirst();
                while (!ItemSubGroupRecordSet.EoF)
                {
                    if (cm_Grp.Value == ItemSubGroupRecordSet.Fields.Item("U_ParentID").Value.ToString())
                    {
                        cm_sGrp.ValidValues.Add(ItemSubGroupRecordSet.Fields.Item("Code").Value.ToString(), ItemSubGroupRecordSet.Fields.Item("Name").Value.ToString());
                    }
                    ItemSubGroupRecordSet.MoveNext();
                }
            }
            catch
            {

            }

        }

        private void Button1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            Models.ItemSet[] selectedMaterials = new Models.ItemSet[] { };
            for (var i = 0; i < gr_List.DataTable.Rows.Count; i++)
            {
                if (gr_List.DataTable.Columns.Item("Chk").Cells.Item(i).Value.ToString() == "Y")
                {
                    Models.ItemSet selectedMaterial = new Models.ItemSet
                    {
                        itemCode = gr_List.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value.ToString(),
                        itemName = gr_List.DataTable.Columns.Item("ItemName").Cells.Item(i).Value.ToString()
                    };
                    selectedMaterials = selectedMaterials.Append(selectedMaterial).ToArray();
                }
                
            }
            Callback?.Invoke(selectedMaterials);
        }
    }
}
 