    using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using SAPbouiCOM.Framework;

namespace PriceList
{
    class Menu
    {
        public void AddMenuItems()
        {
            SAPbouiCOM.MenuItem oMenuItem = null;

            SAPbouiCOM.Menus oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "PriceList";
            oCreationPackage.String = "Fiyat Listesi";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;
            string iconPath = AppDomain.CurrentDomain.BaseDirectory + "price_icon.png";
            if (!File.Exists(iconPath))
                Properties.Resources.price_icon.Save(iconPath, System.Drawing.Imaging.ImageFormat.Png);
            oCreationPackage.Image = iconPath;
            oMenus = oMenuItem.SubMenus; 

            try
            {
                oMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("PriceList");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "PriceList.PriceListForm";
                oCreationPackage.String = "Fiyat Listesi Yönetimi";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "PriceList.DiscountForm";
                oCreationPackage.String = "Dönemsel İndirim Yönetimi";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "PriceList.PriceListReport";
                oCreationPackage.String = "Fiyat Listesi Raporları";
                oMenus.AddEx(oCreationPackage);
            }
            catch
            { //  Menu already exists
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "PriceList.PriceListForm" || pVal.MenuUID == "3076")
                {
                    Form1 activeForm = new Form1();
                    activeForm.Show();
                    BubbleEvent = false;
                }
                if (pVal.BeforeAction && pVal.MenuUID == "PriceList.DiscountForm" || pVal.MenuUID == "11781")
                {
                    DiscountForm activeForm = new DiscountForm();
                    activeForm.Show();
                    BubbleEvent = false;
                }
                if (pVal.BeforeAction && pVal.MenuUID == "PriceList.PriceListReport")
                {
                    PriceListReport activeForm = new PriceListReport();
                    activeForm.Show();
                    BubbleEvent = false;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
