using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;

namespace ProjeTablosu
{
    class Program
    {
        public static SAPbouiCOM.Application SBO_Application;
        public static SAPbobsCOM.Company oCompany;
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
                Application.SBO_Application.AppEvent += SBO_Application_AppEvent;
                ConnectToUI();
                OrganizeTables();
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                

                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void ConnectToUI()
        {
            try
            {
                SAPbouiCOM.SboGuiApi SboGuiApi = new SAPbouiCOM.SboGuiApi();
                string sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

                SboGuiApi.Connect(sConnectionString);
                SBO_Application = SboGuiApi.GetApplication();

                ConnectToCompany();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"UI Bağlantı Hatası: {ex.Message}");
            }
        }
        private static void ConnectToCompany()
        {
            try
            {
                oCompany = (Company)SBO_Application.Company.GetDICompany();

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Company Bağlantı Hatası: {ex.Message}");
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
        private static void OrganizeTables()
        {
            try
            {
                // 1) Ana tabloyu oluştur
                CreateTable("PROJECT", "Proje Raporu", BoUTBTableType.bott_Document);

                // 2) Ana tablo alanlarını ekle
                CreateField("PROJECT", "ProjectTitle", "Proje Başlığı", BoFieldTypes.db_Alpha, 100, null);
                CreateField("PROJECT", "NAME", "Talep Eden Kullanıcı", BoFieldTypes.db_Alpha, 100, null);
                CreateField("PROJECT", "RegDate", "Kayıt Tarihi", BoFieldTypes.db_Date, 0, null);
                CreateField("PROJECT", "DelDate", "İstenilen Proje Teslim Tarihi", BoFieldTypes.db_Date, 0, null);
                CreateField("PROJECT", "Branch", "Şube", BoFieldTypes.db_Alpha, 50, null);
                CreateField("PROJECT", "Department", "Departman", BoFieldTypes.db_Alpha, 50, null);
                CreateField("PROJECT", "IsConverted", "Başka Projeye Dönüştürüldü", BoFieldTypes.db_Alpha, 1,
                new Dictionary<string, string>
                {
                    { "Y", "Evet" },
                    { "N", "Reddedildi" },
                    { "P", "Beklemede" }
                });
                CreateField("PROJECT", "Reject", "Reddedilme Nedeni", BoFieldTypes.db_Memo, 0, null);
                CreateField("OPMG", "ProDocEntry", "Proje Tablosu Belge Numarası", BoFieldTypes.db_Numeric, 0, null);


                // 3) Satır tablosunu oluştur
                CreateTable("PROJECTROW", "Proje Kalem Detayları", BoUTBTableType.bott_DocumentLines);

                // 4) Satır tablosu alanlarını ekle
                CreateField("PROJECTROW", "ItemCode", "Kalem Kodu", BoFieldTypes.db_Alpha, 50, null);
                CreateField("PROJECTROW", "ItemName", "Kalem Tanımı", BoFieldTypes.db_Alpha, 100, null);
                CreateField("PROJECTROW", "ReqDate", "İstenilen Tarih", BoFieldTypes.db_Date, 0, null);
                CreateUDO("PROJECT", "PROJECTROW", "Proje Raporu", BoUDOObjType.boud_Document);

                //SBO_Application.StatusBar.SetText("Proje UDO'su başarıyla oluşturuldu.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Hata: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private static void CreateTable(string tableName, string tableDescription, BoUTBTableType tableType)
        {
            UserTablesMD userTableMD = null;

            try
            {
                userTableMD = (UserTablesMD)oCompany.GetBusinessObject(BoObjectTypes.oUserTables);

                if (!userTableMD.GetByKey(tableName))
                {
                    userTableMD.TableName = tableName;
                    userTableMD.TableDescription = tableDescription;
                    userTableMD.TableType = tableType;

                    int retCode = userTableMD.Add();

                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText($"Hata: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (userTableMD != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(userTableMD);
                    userTableMD = null;
                }
                GC.Collect();
            }
        }

        private static void CreateField(string tableName, string fieldName, string fieldDescription, BoFieldTypes fieldType, int fieldSize, Dictionary<string, string> validValues)
        {
            UserFieldsMD userFieldMD = null;

            try
            {
                userFieldMD = (UserFieldsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserFields);
                userFieldMD.TableName = tableName;
                userFieldMD.Name = fieldName;
                userFieldMD.Description = fieldDescription;
                userFieldMD.Type = fieldType;

                if (fieldSize > 0)
                {
                    userFieldMD.EditSize = fieldSize;
                }

                if (validValues != null)
                {
                    foreach (var entry in validValues)
                    {
                        userFieldMD.ValidValues.Value = entry.Key;
                        userFieldMD.ValidValues.Description = entry.Value;
                        userFieldMD.ValidValues.Add();
                    }
                }

                int retCode = userFieldMD.Add();

            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText($"Hata: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (userFieldMD != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(userFieldMD);
                    userFieldMD = null;
                }
                GC.Collect();
            }
        }
        private static void CreateUDO(String MainTable, String ChildTable, String MenuCaption, SAPbobsCOM.BoUDOObjType ObjectType)
        {
            String UdoName = "UDO" + MainTable;
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            SAPbobsCOM.UserObjectMD_FindColumns oUDOFind = null;
            SAPbobsCOM.UserObjectMD_FormColumns oUDOForm = null;
            SAPbobsCOM.UserObjectMD_EnhancedFormColumns oUDOEnhancedForm = null;
            GC.Collect();
            oUserObjectMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
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
                try
                {
                    int rv = oUserObjectMD.Add();
                }
                catch (Exception ex)
                {
                    SBO_Application.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = null;
            GC.Collect();
        }


    }
}