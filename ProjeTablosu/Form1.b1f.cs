using SAPbouiCOM.Framework;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Xml;
using SAPbobsCOM;
using System.Globalization;
using System.Diagnostics;
using Application = SAPbouiCOM.Framework.Application;

namespace ProjeTablosu
{
    /// <summary>
    /// Project Management Form for SAP Business One  
    /// Handles project creation and management with a user-friendly interface
    /// </summary>
    [FormAttribute("ProjeTablosu.Form1", "Form1.b1f")]
    public class Form1 : UserFormBase
    {
        #region Form Controls
        private SAPbouiCOM.EditText txtProjectName;
        private SAPbouiCOM.EditText txt_DocNum;
        private SAPbouiCOM.EditText txtUser;
        private SAPbouiCOM.ComboBox cmbBranch;
        private SAPbouiCOM.ComboBox cmbDepartment;
        private SAPbouiCOM.EditText txtDeliveryDate;
        private SAPbouiCOM.EditText txtRegistrationDate;
        private SAPbouiCOM.CheckBox chkProject;
        private SAPbouiCOM.Matrix matrixItems;
        private Button btnOk;
        private Button btnCancel;
        private Button btnAddRow;
        private Button btnDellRow;
        #endregion

        #region Constants & Formats
        private const string MATRIX_COLUMN_ITEMCODE = "KalemKodu";
        private const string MATRIX_COLUMN_ITEMNAME = "Kalem_Tan";
        private const string DATE_FORMAT = "yyyyMMdd";          // MSSQL için
        private const string UDO_PROJECT_TABLE = "@PROJECT";
        private const string UDO_PROJECT_ROWS_TABLE = "@PROJECTROW";

        // Kabul edilecek formatlar
        private readonly string[] ALLOWED_DATE_FORMATS = new string[]
        {
            "dd/MM/yyyy",
            "dd.MM.yyyy",
            "yyyyMMdd",
            "MM/dd/yyyy",
            "yyyy-MM-dd",        // HANA formatı
            "yyyy/MM/dd",        // HANA formatı
            "dd.MM.yyyy HH:mm:ss",
            "dd/MM/yyyy HH:mm:ss",
            "dd-MMM-yy hh:mm:ss tt",    // örn: 23-Apr-25 12:00:00 AM
            "dd-MMM-yyyy hh:mm:ss tt"   // örn: 23-Apr-2025 12:00:00 AM
        };
        #endregion

        /// <summary>
        /// Default constructor
        /// </summary>
        public Form1()
        {
            // Constructor intentionally left empty
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            try
            {
                // Initialize form controls
                this.txtProjectName = ((SAPbouiCOM.EditText)(this.GetItem("txt_probas").Specific));
                this.txtUser = ((SAPbouiCOM.EditText)(this.GetItem("txt_kul").Specific));
                this.txtUser.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.OnUserChooseFromListAfter);
                this.txt_DocNum = ((SAPbouiCOM.EditText)(this.GetItem("txt_DocNum").Specific));
                this.cmbBranch = ((SAPbouiCOM.ComboBox)(this.GetItem("cmb_sube").Specific));
                this.cmbDepartment = ((SAPbouiCOM.ComboBox)(this.GetItem("cmb_dep").Specific));
                this.txtDeliveryDate = ((SAPbouiCOM.EditText)(this.GetItem("txt_teslim").Specific));
                this.txtRegistrationDate = ((SAPbouiCOM.EditText)(this.GetItem("txt_kayıt").Specific));
                this.chkProject = ((SAPbouiCOM.CheckBox)(this.GetItem("cb_proje").Specific));
                this.matrixItems = ((SAPbouiCOM.Matrix)(this.GetItem("mtx").Specific));
                this.matrixItems.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.OnMatrixChooseFromListAfter);
                this.btnOk = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
                this.btnOk.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.OnOkButtonPressedBefore);
                this.btnCancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
                this.btnAddRow = ((SAPbouiCOM.Button)(this.GetItem("btn_newrow").Specific));
                this.btnAddRow.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.OnAddRowButtonPressedAfter);
                this.btnDellRow = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
                this.btnDellRow.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.OnDellRowButtonPressedAfter);

            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Error initializing components: {ex.Message}\n{ex.StackTrace}\n");
            }

            this.OnCustomInitialize();
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                // İhtiyaç duyulursa burada ek item event işlemleri yapılabilir.
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Hata oluştu (ItemEvent): {ex.Message}\n{ex.StackTrace}\n");
            }
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            Application.SBO_Application.MenuEvent += SBO_Application_MenuEvent;
        }

        private void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.MenuUID == "1282" && pVal.BeforeAction == false)
            {
                var form = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
                var txtUser = (SAPbouiCOM.EditText)form.Items.Item("txt_kul").Specific;
                var txtRegDate = (SAPbouiCOM.EditText)form.Items.Item("txt_kayıt").Specific;

                // Kullanıcı adı
                if (string.IsNullOrEmpty(txtUser.Value))
                    txtUser.Value = Program.oCompany.UserName;

                // Kayıt tarihi
                if (string.IsNullOrEmpty(txtRegDate.Value))
                    txtRegDate.Value = DateTime.Today.ToString(DATE_FORMAT);
            }
        }

        /// <summary>
        /// Custom initialization logic for the form
        /// </summary>
        private void OnCustomInitialize()
        {
            try
            {
                // Initialize combo boxes
                InitializeBranchComboBox();
                InitializeDepartmentComboBox();
                SAPbobsCOM.Company company = GetCompany();

                // Set user name
                txtUser.Value = company.UserName;

                // Set registration date (today)
                this.txtRegistrationDate.Value = DateTime.Today.ToString("yyyyMMdd");


            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Error during custom initialization: {ex.Message}\n{ex.StackTrace}\n");
            }
        }

        #region Document Loading

        public void SetDocNum(string DocNum)
        {
            try
            {
                LoadData(DocNum);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Form2 değer setlenirken hata: " + ex.Message);
            }
        }

        private void LoadData(string docNumber)
        {
            try
            {
                SAPbobsCOM.Company company = GetCompany();
                SAPbobsCOM.Recordset recordset = null;
                try
                {
                    // SQL sorgusunu dosyadan okuyoruz (SelectProject.sql içeriği: 
                    // "SELECT * FROM [@PROJECT] WHERE DocNum = @DocNum")
                    var replacements = new Dictionary<string, string>
                    {
                        { "@DocNum", docNumber }
                    };
                    if (company.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        recordset = Helper.executeSQLFromFile("SelectProject", replacements, string.Empty, string.Empty);
                    }
                    else
                    {
                        recordset = Helper.executeSQLFromFile("SelectProjectHana", replacements, string.Empty, string.Empty);
                    }

                    if (recordset.RecordCount > 0)
                    {
                        // Fill header fields
                        txtProjectName.Value = recordset.Fields.Item("U_ProjectTitle").Value.ToString();
                        txtUser.Value = recordset.Fields.Item("U_NAME").Value.ToString();

                        // Select Branch and Department in ComboBoxes
                        string branch = recordset.Fields.Item("U_Branch").Value.ToString();
                        string department = recordset.Fields.Item("U_Department").Value.ToString();
                        for (int i = 0; i < cmbBranch.ValidValues.Count; i++)
                        {
                            if (cmbBranch.ValidValues.Item(i).Value == branch)
                            {
                                cmbBranch.Select(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                break;
                            }
                        }
                        for (int i = 0; i < cmbDepartment.ValidValues.Count; i++)
                        {
                            if (cmbDepartment.ValidValues.Item(i).Value == department)
                            {
                                cmbDepartment.Select(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                break;
                            }
                        }

                        // Set dates if available
                        if (recordset.Fields.Item("U_DelDate").Value != null)
                        {
                            DateTime deliveryDate = (DateTime)recordset.Fields.Item("U_DelDate").Value;
                            txtDeliveryDate.Value = deliveryDate.ToString(DATE_FORMAT);
                        }
                        if (recordset.Fields.Item("U_RegDate").Value != null)
                        {
                            DateTime registrationDate = (DateTime)recordset.Fields.Item("U_RegDate").Value;
                            txtRegistrationDate.Value = registrationDate.ToString(DATE_FORMAT);
                        }

                        // Set project checkbox
                        chkProject.Checked = recordset.Fields.Item("U_IsConverted").Value.ToString() == "Y";

                        // Load document lines using SQL file for project rows
                        LoadDocumentLines(docNumber);
                    }
                    else
                    {
                        ShowMessage($"Belirtilen doküman numarasına ({docNumber}) ait kayıt bulunamadı.");
                    }
                }
                finally
                {
                    if (recordset != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);
                        recordset = null;
                        GC.Collect();
                    }
                }
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Doküman ({docNumber}) yüklenirken hata oluştu: {ex.Message}\n{ex.StackTrace}\n");
                ShowMessage($"Doküman yüklenirken bir hata oluştu: {ex.Message}");
            }
        }

        private void LoadDocumentLines(string docNumber)
        {
            SAPbobsCOM.Company company = GetCompany();
            SAPbobsCOM.Recordset recordset = null;
            try
            {
                this.UIAPIRawForm.Freeze(true);
                // SQL dosyası "SelectProjectRows.sql" içeriği şöyle olmalı:
                // "SELECT * FROM [@PROJECTROW] WHERE DocEntry = @DocNum ORDER BY LineId"
                var replacements = new Dictionary<string, string>
                {
                    { "@DocNum", docNumber }
                };
                if (company.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    recordset = Helper.executeSQLFromFile("SelectProjectRows", replacements, string.Empty, string.Empty);
                }
                else
                {
                    recordset = Helper.executeSQLFromFile("SelectProjectRowsHana", replacements, string.Empty, string.Empty);
                }

                SAPbouiCOM.DBDataSource dbDataSourceLines = UIAPIRawForm.DataSources.DBDataSources.Item(UDO_PROJECT_ROWS_TABLE);
                dbDataSourceLines.Clear();

                int rowIndex = 0;
                while (!recordset.EoF)
                {
                    dbDataSourceLines.InsertRecord(rowIndex);
                    dbDataSourceLines.SetValue("U_ItemCode", rowIndex, recordset.Fields.Item("U_ItemCode").Value.ToString());
                    dbDataSourceLines.SetValue("U_ItemName", rowIndex, recordset.Fields.Item("U_ItemName").Value.ToString());
                    if (recordset.Fields.Item("U_ReqDate").Value != null)
                    {
                        DateTime reqDate = (DateTime)recordset.Fields.Item("U_ReqDate").Value;
                        dbDataSourceLines.SetValue("U_ReqDate", rowIndex, reqDate.ToString(DATE_FORMAT));
                    }
                    rowIndex++;
                    recordset.MoveNext();
                }
                matrixItems.LoadFromDataSource();
                UpdateMatrixRowNumbers();
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Doküman satırları ({docNumber}) yüklenirken hata oluştu: {ex.Message}\n{ex.StackTrace}\n");
                ShowMessage($"Doküman satırları yüklenirken bir hata oluştu: {ex.Message}");
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                if (recordset != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);
                    recordset = null;
                    GC.Collect();
                }
            }
        }
        #endregion

        #region ComboBox Initialization

        private void InitializeBranchComboBox()
        {
            SAPbobsCOM.Company company = GetCompany();
            SAPbobsCOM.Recordset recordset = null;
            try
            {
                // SQL dosyası "SelectBranch.sql" içeriği: "SELECT DISTINCT Name FROM OUBR"
                if (company.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    recordset = Helper.executeSQLFromFile("SelectBranch", new Dictionary<string, string>(), string.Empty, string.Empty);
                }
                else
                {
                    recordset = Helper.executeSQLFromFile("SelectBranchHana", new Dictionary<string, string>(), string.Empty, string.Empty);
                }
                ClearComboBoxValues(cmbBranch);
                if (recordset.RecordCount > 0)
                {
                    while (!recordset.EoF)
                    {
                        string branchName = recordset.Fields.Item("Name").Value.ToString();
                        cmbBranch.ValidValues.Add(branchName, branchName);
                        recordset.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Error initializing branch combo box: {ex.Message}\n{ex.StackTrace}\n");
            }
            finally
            {
                if (recordset != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);
                    recordset = null;
                    GC.Collect();
                }
            }
        }

        private void InitializeDepartmentComboBox()
        {
            SAPbobsCOM.Company company = GetCompany();
            SAPbobsCOM.Recordset recordset = null;
            try
            {
                // SQL dosyası "SelectDepartment.sql" içeriği: "SELECT DISTINCT Name FROM OUDP"
                if (company.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    recordset = Helper.executeSQLFromFile("SelectDepartment", new Dictionary<string, string>(), string.Empty, string.Empty);
                }
                else
                {
                    recordset = Helper.executeSQLFromFile("SelectDepartmentHana", new Dictionary<string, string>(), string.Empty, string.Empty);
                }
                ClearComboBoxValues(cmbDepartment);
                if (recordset.RecordCount > 0)
                {
                    while (!recordset.EoF)
                    {
                        string departmentName = recordset.Fields.Item("Name").Value.ToString();
                        cmbDepartment.ValidValues.Add(departmentName, departmentName);
                        recordset.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Error initializing department combo box: {ex.Message}\n{ex.StackTrace}\n");
            }
            finally
            {
                if (recordset != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);
                    recordset = null;
                    GC.Collect();
                }
            }
        }

        private void ClearComboBoxValues(SAPbouiCOM.ComboBox comboBox)
        {
            if (comboBox.ValidValues.Count > 0)
            {
                for (int i = comboBox.ValidValues.Count - 1; i >= 0; i--)
                {
                    comboBox.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            }
        }
        #endregion

        #region Choose From List Event Handlers

        private void HandleChooseFromListAfter(string getValue, string setValue, string tableName, SBOItemEventArg pVal)
        {
            try
            {
                SBOChooseFromListEventArg cfl = (SBOChooseFromListEventArg)pVal;
                if (cfl.SelectedObjects == null)
                {
                    return;
                }
                string value = cfl.SelectedObjects.GetValue(getValue, 0).ToString();
                UIAPIRawForm.DataSources.DBDataSources.Item(tableName).SetValue(setValue, 0, value);
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Error in HandleChooseFromListAfter: {ex.Message}\n{ex.StackTrace}\n");
            }
        }

        private void HandleChooseFromListAfterMatrixItems(string getValue, string setValue, string tableName, SBOItemEventArg pVal)
        {
            try
            {
                SBOChooseFromListEventArg cfl = (SBOChooseFromListEventArg)pVal;
                if (cfl.SelectedObjects == null)
                {
                    return;
                }
                string value = cfl.SelectedObjects.GetValue(getValue, 0).ToString();
                var dataSource = UIAPIRawForm.DataSources.DBDataSources.Item(tableName);
                dataSource.SetValue(setValue, pVal.Row - 1, value);
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Error in HandleChooseFromListAfterMatrixItems: {ex.Message}\n{ex.StackTrace}\n");
            }
        }

        private void OnUserChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            HandleChooseFromListAfter("U_NAME", "U_NAME", UDO_PROJECT_TABLE, pVal);
        }

        private void OnMatrixChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID == MATRIX_COLUMN_ITEMCODE || pVal.ColUID == MATRIX_COLUMN_ITEMNAME)
                {
                    HandleChooseFromListAfterMatrixItems("ItemCode", "U_ItemCode", UDO_PROJECT_ROWS_TABLE, pVal);
                    HandleChooseFromListAfterMatrixItems("ItemName", "U_ItemName", UDO_PROJECT_ROWS_TABLE, pVal);
                }
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Error in OnMatrixChooseFromListAfter: {ex.Message}\n{ex.StackTrace}\n");
            }
        }
        #endregion

        #region Button Event Handlers

        private void OnOkButtonPressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (!ValidateRequiredFields())
                {
                    BubbleEvent = false;
                    return;
                }
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Error in OnOkButtonPressedBefore: {ex.Message}\n{ex.StackTrace}\n");
                BubbleEvent = false;
            }
        }



        private void OnAddRowButtonPressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ActionSuccess)
            {
                try
                {
                    AddNewMatrixRow();
                }
                catch (Exception ex)
                {
                    Helper.LogToFile($"Error in OnAddRowButtonPressedAfter: {ex.Message}\n{ex.StackTrace}\n");
                }
            }
        }
        private void OnDellRowButtonPressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ActionSuccess)
            {
                try
                {
                    DellMatrixRow();
                }
                catch (Exception ex)
                {
                    Helper.LogToFile($"Error in OnAddRowButtonPressedAfter: {ex.Message}\n{ex.StackTrace}\n");
                }
            }
        }
        #endregion

        #region Date Parsing Helper
        /// <summary>
        /// Verilen string'i ALLOWED_DATE_FORMATS üzerinden deniyor.
        /// Başarılı olursa out parametrelerine set ediyor.
        /// </summary>
        private bool TryDetectDateFormat(string dateString, out DateTime parsedDate, out string usedFormat)
        {
            parsedDate = DateTime.MinValue;
            usedFormat = null;

            foreach (var fmt in ALLOWED_DATE_FORMATS)
            {
                if (DateTime.TryParseExact(
                        dateString,
                        fmt,
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out parsedDate))
                {
                    usedFormat = fmt;
                    return true;
                }
            }
            return false;
        }
        #endregion

        #region Matrix Operations
        private void AddNewMatrixRow()
        {
            UIAPIRawForm.Freeze(true);
            try
            {
                // Matrix’teki mevcut verileri DataSource’a bas
                try { matrixItems.FlushToDataSource(); } catch { /* boş tarih satırı HANA’da hata fırlatabilir → yok say */ }

                var lines = UIAPIRawForm.DataSources.DBDataSources.Item(UDO_PROJECT_ROWS_TABLE);

                // Gereksiz boş ilk satırı sil
                if (lines.Size == 1 && string.IsNullOrWhiteSpace(lines.GetValue("U_ItemName", 0)))
                    lines.RemoveRecord(0);

                // Yeni satırı ekle
                int row = lines.Size;
                lines.InsertRecord(row);
                lines.SetValue("U_ItemCode", row, "");
                lines.SetValue("U_ItemName", row, "");

                // Teslim tarihini algıla
                if (!TryDetectDateFormat(txtDeliveryDate.Value, out DateTime del, out _))
                {
                    ShowMessage($"Tarih formatı anlaşılamadı: {txtDeliveryDate.Value}");
                    return;
                }

                // *** Kritik kısım: UI için daima YYYYMMDD ***
                lines.SetValue("U_ReqDate", row, del.ToString("yyyyMMdd"));

                // Matrix’i yeniden yükle
                matrixItems.LoadFromDataSource();
                matrixItems.SetLineData(matrixItems.RowCount);
                UpdateMatrixRowNumbers();
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Add row err: {ex}");
                ShowMessage("Satır eklenirken hata: " + ex.Message);
            }
            finally { UIAPIRawForm.Freeze(false); }
        }
        private void DellMatrixRow()
        {
            UIAPIRawForm.Freeze(true);
            try
            {
                // 1) Matrix üzerindeki değişiklikleri DataSource'a aktar
                matrixItems.FlushToDataSource();

                // 2) UDO child satırlarının DBDataSource'unu al
                var lines = UIAPIRawForm.DataSources
                              .DBDataSources
                              .Item(UDO_PROJECT_ROWS_TABLE);

                // 3) Seçili satırı bul (1-based)
                int selectedRow = matrixItems.GetNextSelectedRow(
                                      0,
                                      SAPbouiCOM.BoOrderType.ot_RowOrder
                                  );
                if (selectedRow <= 0)
                {
                    ShowMessage("Lütfen silmek istediğiniz satırı seçin!");
                    return;
                }

                // 4) 1-based → 0-based index dönüşümü
                int index = selectedRow - 1;

                // 5) Geçerli index mi? Sil ve UI’ı yenile
                if (index >= 0 && index < lines.Size)
                {
                    lines.RemoveRecord(index);
                    matrixItems.LoadFromDataSource();
                    UpdateMatrixRowNumbers();
                }
                else
                {
                    ShowMessage($"Silinecek satır bulunamadı: {selectedRow}");
                }
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Delete row err: {ex.Message}\n{ex.StackTrace}");
                ShowMessage("Satır silinirken hata oluştu: " + ex.Message);
            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        private void UpdateMatrixRowNumbers()
        {
            for (int i = 1; i <= matrixItems.RowCount; i++)
            {
                ((SAPbouiCOM.EditText)matrixItems.Columns.Item("#").Cells.Item(i).Specific).Value = i.ToString();
            }
        }
        #endregion

        #region Validation

        private bool ValidateRequiredFields()
        {
            var form = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            if (form.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                return true;

            // Başlık alanları
            if (string.IsNullOrWhiteSpace(txtProjectName.Value) ||
                string.IsNullOrWhiteSpace(txtUser.Value) ||
                string.IsNullOrWhiteSpace(txtDeliveryDate.Value) ||
                string.IsNullOrWhiteSpace(txtRegistrationDate.Value))
            {
                ShowMessage("Lütfen tüm zorunlu alanları doldurunuz!");
                return false;
            }

            if (cmbBranch.Selected == null || cmbDepartment.Selected == null)
            {
                ShowMessage("Lütfen şube ve departman seçiniz!");
                return false;
            }

            // Başlık tarihleri: farklı formatları kabul et
            if (!TryDetectDateFormat(txtRegistrationDate.Value, out DateTime regDate, out _) ||
                !TryDetectDateFormat(txtDeliveryDate.Value, out DateTime delDate, out _))
            {
                ShowMessage($"Tarih formatı hatalı! Geçerli formatlar: {string.Join(", ", ALLOWED_DATE_FORMATS)}");
                return false;
            }
            if (regDate > delDate)
            {
                ShowMessage("Kayıt tarihi, teslim tarihinden büyük olamaz.");
                return false;
            }

            // Matrix satırları – önce UI’dan DataSource’a yaz
            try { matrixItems.FlushToDataSource(); } catch { /* ignore */ }
            var lines = UIAPIRawForm.DataSources.DBDataSources.Item(UDO_PROJECT_ROWS_TABLE);

            if (lines.Size == 0)
            {
                ShowMessage("Lütfen en az bir satır ekleyiniz!");
                return false;
            }

            // Her satırı kontrol et
            for (int row = 0; row < lines.Size; row++)
            {
                string itemName = lines.GetValue("U_ItemName", row)?.Trim() ?? "";
                string reqDateStr = lines.GetValue("U_ReqDate", row)?.Trim() ?? "";

                if (string.IsNullOrEmpty(itemName) || string.IsNullOrEmpty(reqDateStr))
                {
                    ShowMessage($"Satır {row + 1}: Kalem ve Tarih alanları boş bırakılamaz!");
                    return false;
                }
                if (!TryDetectDateFormat(reqDateStr, out _, out _))
                {
                    ShowMessage($"Satır {row + 1}: Tarih formatı hatalı ({reqDateStr})!");
                    return false;
                }
            }

            return true;
        }

        #endregion

        #region Utility Methods

        private SAPbobsCOM.Company GetCompany()
        {
            return (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
        }

        private void ShowMessage(string message)
        {
            Application.SBO_Application.MessageBox(message);
        }
        #endregion
    }
}
