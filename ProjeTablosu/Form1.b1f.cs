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
        #endregion

        #region Constants
        private const string MATRIX_COLUMN_ITEMCODE = "KalemKodu";
        private const string MATRIX_COLUMN_ITEMNAME = "Kalem_Tan";
        private const string DATE_FORMAT = "yyyyMMdd";
        private const string UDO_PROJECT_TABLE = "@PROJECT";
        private const string UDO_PROJECT_ROWS_TABLE = "@PROJECTROW";
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
                this.btnOk.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.OnOkButtonPressedAfter);
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
                SAPbouiCOM.Form form = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
                var Usern = (SAPbouiCOM.EditText)form.Items.Item("txt_kul").Specific;
                if (string.IsNullOrEmpty(Usern.Value))
                {
                    Usern.Value = Program.oCompany.UserName;
                }
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
                txtRegistrationDate.Value = DateTime.Today.ToString(DATE_FORMAT);
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
                    recordset = Helper.executeSQLFromFile("SelectProject", replacements, string.Empty, string.Empty);

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
                recordset = Helper.executeSQLFromFile("SelectProjectRows", replacements, string.Empty, string.Empty);

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
                recordset = Helper.executeSQLFromFile("SelectBranch", new Dictionary<string, string>(), string.Empty, string.Empty);
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
                recordset = Helper.executeSQLFromFile("SelectDepartment", new Dictionary<string, string>(), string.Empty, string.Empty);
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

        private void OnOkButtonPressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ActionSuccess)
            {
                ShowMessage("Proje başarıyla kaydedildi.");
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
        #endregion

        #region Matrix Operations

        private void AddNewMatrixRow()
        {
            this.UIAPIRawForm.Freeze(true);
            try
            {
                matrixItems.FlushToDataSource();
                SAPbouiCOM.DBDataSource dbDataSourceLines = UIAPIRawForm.DataSources.DBDataSources.Item(UDO_PROJECT_ROWS_TABLE);
                if (dbDataSourceLines.Size == 1 && string.IsNullOrEmpty(dbDataSourceLines.GetValue("U_ItemName", 0).Trim()))
                {
                    dbDataSourceLines.RemoveRecord(0);
                }
                int rowIndex = dbDataSourceLines.Size;
                dbDataSourceLines.InsertRecord(rowIndex);
                dbDataSourceLines.Offset = rowIndex;
                dbDataSourceLines.SetValue("U_ItemCode", rowIndex, "");
                dbDataSourceLines.SetValue("U_ItemName", rowIndex, "");
                dbDataSourceLines.SetValue("U_ReqDate", rowIndex, txtDeliveryDate.Value);
                matrixItems.LoadFromDataSource();
                matrixItems.SetLineData(matrixItems.RowCount);
                UpdateMatrixRowNumbers();
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Error adding new matrix row: {ex.Message}\n{ex.StackTrace}\n");
                ShowMessage("Error adding new row: " + ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
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
            if (string.IsNullOrEmpty(txtProjectName.Value.Trim()) ||
                string.IsNullOrEmpty(txtUser.Value.Trim()) ||
                string.IsNullOrEmpty(txtDeliveryDate.Value.Trim()) ||
                string.IsNullOrEmpty(txtRegistrationDate.Value.Trim()))
            {
                ShowMessage("Please fill in all required text fields!");
                return false;
            }
            if (cmbBranch.Selected == null || string.IsNullOrEmpty(cmbBranch.Selected.Value) ||
                cmbDepartment.Selected == null || string.IsNullOrEmpty(cmbDepartment.Selected.Value))
            {
                ShowMessage("Please select both branch and department!");
                return false;
            }
            SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)this.GetItem("mtx").Specific;
            int rowCount = matrix.RowCount;
            if (rowCount == 0)
            {
                ShowMessage("Lütfen en az bir satır ekleyiniz!");
                return false;
            }
            if (rowCount > 0)
            {
                for (int i = 1; i <= rowCount; i++)
                {
                    string kalemTan = ((SAPbouiCOM.EditText)matrix.Columns.Item("Kalem_Tan").Cells.Item(i).Specific).Value.Trim();
                    string tarih = ((SAPbouiCOM.EditText)matrix.Columns.Item("Tarih").Cells.Item(i).Specific).Value.Trim();
                    if (string.IsNullOrEmpty(kalemTan) || string.IsNullOrEmpty(tarih))
                    {
                        ShowMessage($"Row {i}: Item definition and date fields cannot be empty!");
                        return false;
                    }
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
