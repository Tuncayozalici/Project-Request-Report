using SAPbouiCOM.Framework;
using SAPbouiCOM;
using System;
using System.Text;
using SAPbobsCOM; // DI API için kullanılıyor
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;
using ProjeTablosu;  // Helper sınıfını içeren namespace (aynı projede)
using Application = SAPbouiCOM.Framework.Application;

namespace ProjeTablosu
{
    [FormAttribute("ProjeTablosu.Form2", "Form2.b1f")]
    class Form2 : UserFormBase
    {
        #region Form Kontrolleri
        private Button btn_prj;
        private Grid grd_liste;
        private EditText txt_reg;      // Başlangıç tarihi
        private EditText txt_regbit;   // Bitiş tarihi
        private Button btn_fltr;
        private CheckBox check_prj;    // CheckBox: Filtre için kullanılıyor
        #endregion

        #region Sabitler
        /// <summary>
        /// Tarih için kabul edilebilir formatlar.
        /// Örneğin "15.04.2025", "15.04.2025 00:00:00", "20250415" vb.
        /// </summary>
        private readonly string[] ALLOWED_DATE_FORMATS = new string[]
        {
            "dd/MM/yyyy",
            "dd.MM.yyyy",
            "yyyyMMdd",
            "MM/dd/yyyy",
            "dd.MM.yyyy HH:mm:ss",
            "dd/MM/yyyy HH:mm:ss"
        };

        private const string UDO_PROJECT_TABLE = "@PROJECT";
        private const string UDO_PROJECT_ROWS_TABLE = "@PROJECTROW";
        #endregion

        #region Yapıcı ve Başlatma Metodları

        public Form2() { }

        /// <summary>
        /// Form üzerindeki kontrollerin referanslarının alınması.
        /// </summary>
        public override void OnInitializeComponent()
        {
            try
            {
                this.btn_prj = ((SAPbouiCOM.Button)(this.GetItem("btn_prj").Specific));
                this.grd_liste = ((SAPbouiCOM.Grid)(this.GetItem("grd_liste").Specific));
                this.grd_liste.DoubleClickBefore += new SAPbouiCOM._IGridEvents_DoubleClickBeforeEventHandler(this.grd_liste_DoubleClickBefore);
                this.txt_reg = ((SAPbouiCOM.EditText)(this.GetItem("txt_reg").Specific));
                this.txt_regbit = ((SAPbouiCOM.EditText)(this.GetItem("txt_regbit").Specific));
                this.btn_fltr = ((SAPbouiCOM.Button)(this.GetItem("btn_fltr").Specific));
                this.check_prj = ((SAPbouiCOM.CheckBox)(this.GetItem("check_prj").Specific));
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Form2 OnInitializeComponent hatası: {ex.Message}\n{ex.StackTrace}\n");
            }
            this.OnCustomInitialize();
        }

        public override void OnInitializeFormEvents()
        {
            // Form eventleri eklenecekse buraya
        }

        /// <summary>
        /// Özel başlangıç işlemleri: DataTable oluşturma, sorgu çalıştırma ve grid’e atama.
        /// Form açılışında tarih aralığı uygulanır fakat checkbox filtresi (U_IsConverted) uygulanmaz.
        /// Ayrıca sorgu sonucu BuildQuery metoduyla oluşturulan sorgu Helper.LogToFile ile txt dosyasına yazılır.
        /// </summary>
        private void OnCustomInitialize()
        {
            try
            {
                this.btn_prj.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.OnProjectButtonClickBefore);
                this.btn_fltr.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnFilter_ClickBefore);

                // EditText'lere varsayılan değer ataması ("yyyyMMdd" formatında)
                this.txt_reg.Value = DateTime.Today.ToString("yyyyMMdd");
                this.txt_regbit.Value = DateTime.Today.ToString("yyyyMMdd");

                // Form açılışında, checkbox filtresi uygulanmadan sorgu çalışsın (onlyConverted = false)
                SAPbouiCOM.DataTable oDataTable = this.UIAPIRawForm.DataSources.DataTables.Add("MyTable");
                string initialQuery = BuildQuery(this.txt_reg.Value, this.txt_regbit.Value, false);
                oDataTable.ExecuteQuery(initialQuery);
                this.grd_liste.DataTable = oDataTable;

                CheckUserDepartmentPermissions();
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Form2 OnCustomInitialize hatası: {ex.Message}\n{ex.StackTrace}\n");
            }
        }

        #endregion

        #region SQL Sorgusu Oluşturma ve Loglama

        /// <summary>
        /// SQL sorgusunu oluşturur.
        /// SQL dosyası, veritabanı türüne göre (MSSQL/HANA) farklı scriptleri kullanacak şekilde yüklenir.
        /// Oluşturulan sorgu Helper.LogToFile metodu ile txt dosyasına kaydedilir.
        /// </summary>
        /// <param name="startDateFilter">Başlangıç tarihi (yyyyMMdd formatı)</param>
        /// <param name="endDateFilter">Bitiş tarihi (yyyyMMdd formatı)</param>
        /// <param name="onlyConverted">Sadece U_IsConverted = 'Y' kayıtları</param>
        /// <returns>Oluşturulmuş sorgu stringi</returns>
        private string BuildQuery(string startDateFilter, string endDateFilter, bool onlyConverted)
        {
            SAPbobsCOM.Company company = GetCompany();
            string baseQuery = string.Empty;

            // Veritabanı türüne göre uygun SQL script dosyasını yüklüyoruz.
            if (company.DbServerType != BoDataServerTypes.dst_HANADB)
            {
                baseQuery = Helper.LoadSqlScript("SelectProjectList.sql", BoDataServerTypes.dst_MSSQL);
            }
            else
            {
                baseQuery = Helper.LoadSqlScript("SelectProjectListHana.sql", BoDataServerTypes.dst_HANADB);
            }

            string dateCondition = "";
            if (!string.IsNullOrEmpty(startDateFilter) && !string.IsNullOrEmpty(endDateFilter))
            {
                if (DateTime.TryParseExact(startDateFilter, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dtStart) &&
                    DateTime.TryParseExact(endDateFilter, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dtEnd))
                {
                    string formattedStart = dtStart.ToString("dd.MM.yyyy");
                    string formattedEnd = dtEnd.ToString("dd.MM.yyyy");
                    if (company.DbServerType != BoDataServerTypes.dst_HANADB)
                    {
                        // MSSQL için CONVERT kullanıyoruz.
                        dateCondition = $"AND U_RegDate BETWEEN CONVERT(DATETIME, '{formattedStart}', 104) AND CONVERT(DATETIME, '{formattedEnd}', 104) ";
                    }
                    else
                    {
                        // HANA için TO_DATE kullanıyoruz.
                        dateCondition = $"AND \"U_RegDate\" BETWEEN TO_DATE('{formattedStart}', 'DD.MM.YYYY') AND TO_DATE('{formattedEnd}', 'DD.MM.YYYY') ";
                    }
                }
            }
            else if (!string.IsNullOrEmpty(startDateFilter))
            {
                if (DateTime.TryParseExact(startDateFilter, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dtStart))
                {
                    string formattedStart = dtStart.ToString("dd.MM.yyyy");
                    if (company.DbServerType != BoDataServerTypes.dst_HANADB)
                    {
                        dateCondition = $"AND U_RegDate >= CONVERT(DATETIME, '{formattedStart}', 104) ";
                    }
                    else
                    {
                        dateCondition = $"AND \"U_RegDate\" >= TO_DATE('{formattedStart}', 'DD.MM.YYYY') ";
                    }
                }
            }
            else if (!string.IsNullOrEmpty(endDateFilter))
            {
                if (DateTime.TryParseExact(endDateFilter, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dtEnd))
                {
                    string formattedEnd = dtEnd.ToString("dd.MM.yyyy");
                    if (company.DbServerType != BoDataServerTypes.dst_HANADB)
                    {
                        dateCondition = $"AND U_RegDate <= CONVERT(DATETIME, '{formattedEnd}', 104) ";
                    }
                    else
                    {
                        dateCondition = $"AND \"U_RegDate\" <= TO_DATE('{formattedEnd}', 'DD.MM.YYYY') ";
                    }
                }
            }

            // U_IsConverted filtrelemesi
            string convertedCondition = onlyConverted
                ? "AND U_IsConverted = 'Y' "
                : "AND (\"U_IsConverted\" <> 'Y' OR \"U_IsConverted\" IS NULL) ";

            // SQL scriptindeki placeholder'ları oluşturduğumuz koşullarla değiştiriyoruz.
            string finalQuery = baseQuery.Replace("--DATEFILTER--", dateCondition)
                                         .Replace("--CONVERTEDFILTER--", convertedCondition);

            // Oluşturulan sorguyu txt dosyasına logluyoruz.
            Helper.LogToFile($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Executing Query:\n{finalQuery}\n");

            return finalQuery;
        }

        #endregion

        #region Button Event Handlers

        private void BtnFilter_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                string startFilter = txt_reg.Value.Trim();
                string endFilter = txt_regbit.Value.Trim();
                bool onlyConverted = check_prj.Checked;

                SAPbouiCOM.DataTable oDataTable = this.UIAPIRawForm.DataSources.DataTables.Item("MyTable");
                // BuildQuery çağrısı içinde sorgu oluşturulup loglanacaktır.
                oDataTable.ExecuteQuery(BuildQuery(startFilter, endFilter, onlyConverted));
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Filtreleme hatası: {ex.Message}\n{ex.StackTrace}\n");
                BubbleEvent = false;
            }
        }

        private void OnProjectButtonClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Form form = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Item btnPrjItem = form.Items.Item("btn_prj");
                if (!btnPrjItem.Enabled)
                {
                    BubbleEvent = false;
                    return;
                }
                CreateNewProject();
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Projeye geçiş hatası: {ex.Message}\n{ex.StackTrace}\n");
                BubbleEvent = false;
            }
        }

        #endregion

        #region Project Management

        /// <summary>
        /// Seçili satırdaki proje isteği üzerinden yeni proje oluşturur.
        /// Hem MSSQL hem de HANA ortamında çalışacak şekilde güncellendi.
        /// </summary>
        private void CreateNewProject()
        {
            if (this.grd_liste.Rows.SelectedRows.Count == 0)
            {
                ShowMessage("Lütfen yeni proje oluşturmak için bir satır seçiniz.");
                return;
            }

            int selectedRow = this.grd_liste.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder);
            SAPbouiCOM.DataTable dt = this.grd_liste.DataTable;
            this.UIAPIRawForm.Freeze(true);

            string docEntry = dt.GetValue("Döküman Numarası", selectedRow)?.ToString().Trim() ?? "";
            string projectTitle = dt.GetValue("Proje Talebi Tanımı", selectedRow)?.ToString().Trim() ?? "";
            string regDateStr = dt.GetValue("Kayıt Tarihi", selectedRow)?.ToString().Trim() ?? "";
            string delDateStr = dt.GetValue("İstenilen Tarih", selectedRow)?.ToString().Trim() ?? "";

            Trace.WriteLine("DEBUG - Kayıt Tarihi: " + regDateStr);
            Trace.WriteLine("DEBUG - İstenilen Tarih: " + delDateStr);

            SAPbobsCOM.Company company = GetCompany();
            SAPbobsCOM.CompanyService companyService = null;
            SAPbobsCOM.ProjectManagementService projectService = null;
            SAPbobsCOM.Recordset recordset = null;

            try
            {
                companyService = company.GetCompanyService();
                projectService = (SAPbobsCOM.ProjectManagementService)companyService.GetBusinessService(ServiceTypes.ProjectManagementService);

                SAPbobsCOM.PM_ProjectDocumentData newProject =
                    (SAPbobsCOM.PM_ProjectDocumentData)projectService.GetDataInterface(ProjectManagementServiceDataInterfaces.pmsPM_ProjectDocumentData);

                string projectName = $"{projectTitle} {docEntry}";
                newProject.ProjectName = projectName;

                // Kullanıcı alanı U_ProDocEntry'nin setlenmesi
                for (int i = 0; i < newProject.UserFields.Count; i++)
                {
                    if (newProject.UserFields.Item(i).Name == "U_ProDocEntry")
                    {
                        newProject.UserFields.Item(i).Value = docEntry;
                        break;
                    }
                }

                // Kayıt Tarihi (StartDate) ayarlaması
                if (!string.IsNullOrEmpty(regDateStr))
                {
                    if (DateTime.TryParseExact(regDateStr, ALLOWED_DATE_FORMATS, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate))
                    {
                        newProject.StartDate = startDate;
                    }
                    else
                    {
                        ShowMessage("Kayıt Tarihi geçerli bir formatta değil: " + regDateStr);
                        return;
                    }
                }

                // İstenilen Tarih (DueDate) ayarlaması
                if (!string.IsNullOrEmpty(delDateStr))
                {
                    if (DateTime.TryParseExact(delDateStr, ALLOWED_DATE_FORMATS, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dueDate))
                    {
                        newProject.DueDate = dueDate;
                    }
                    else
                    {
                        ShowMessage("İstenilen Tarih geçerli bir formatta değil: " + delDateStr);
                        return;
                    }
                }

                SAPbobsCOM.PM_ProjectDocumentParams newProjectParams = projectService.AddProject(newProject);
                int newProjectAbsEntry = newProjectParams.AbsEntry;

                // Update sorgusu: U_ProDocEntry alanını güncellemek için kullanılır.
                // HANA ve MSSQL arasında söz dizimi farklılıkları olabileceğinden kontroller eklenmiştir.
                recordset = (SAPbobsCOM.Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string updateQuery = "";
                if (company.DbServerType != BoDataServerTypes.dst_HANADB)
                {
                    updateQuery = $"UPDATE OPMG SET U_ProDocEntry = '{docEntry}' WHERE AbsEntry = {newProjectAbsEntry}";
                }
                else
                {
                    // HANA ortamında tablo/sütun isimleri büyük harf ile tanımlanmışsa/veya tırnak gerekliyse:
                    updateQuery = $"UPDATE \"OPMG\" SET \"U_ProDocEntry\" = '{docEntry}' WHERE \"AbsEntry\" = {newProjectAbsEntry}";
                }
                // Güncelleme sorgusunu çalıştırmak isterseniz:
                // recordset.DoQuery(updateQuery);

                dt.SetValue("Projeye Dönüştürüldü", selectedRow, "Evet");
                this.UIAPIRawForm.Freeze(false);

                ShowStatusMessage($"Proje başarıyla oluşturuldu. Proje Kodu: {newProjectAbsEntry}", BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Yeni proje oluşturma hatası: {ex.Message}\n{ex.StackTrace}\n");
                ShowMessage("Yeni proje oluşturulurken hata oluştu: " + ex.Message);
            }
            finally
            {
                if (recordset != null)
                {
                    Marshal.ReleaseComObject(recordset);
                    recordset = null;
                }
                if (projectService != null)
                {
                    Marshal.ReleaseComObject(projectService);
                    projectService = null;
                }
                if (companyService != null)
                {
                    Marshal.ReleaseComObject(companyService);
                    companyService = null;
                }
                GC.Collect();
            }
        }

        #endregion

        #region Grid Event Handler

        private void grd_liste_DoubleClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                Grid clickedGrid = sboObject as Grid;
                if (clickedGrid == null)
                    return;
                int rowIndex = pVal.Row;
                if (rowIndex < 0)
                    return;

                string DocNum = clickedGrid.DataTable.GetValue("Döküman Numarası", rowIndex)?.ToString();
                if (string.IsNullOrEmpty(DocNum))
                {
                    throw new Exception("Döküman Numarası boş olamaz.");
                }

                // Form1 üzerinden ilgili kaydı açmak için Form1'i çağırıyoruz.
                Form1 form1 = new Form1();
                form1.SetDocNum(DocNum);
                form1.Show();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Çift tıklama işlemi sırasında hata: " + ex.Message);
            }
        }

        #endregion

        #region Permission Management

        /// <summary>
        /// Kullanıcının bağlı olduğu departmana göre proje oluşturma butonunu devre dışı bırakır.
        /// </summary>
        private void CheckUserDepartmentPermissions()
        {
            try
            {
                SAPbobsCOM.Company company = GetCompany();
                SAPbobsCOM.Recordset recordset = null;
                try
                {
                    recordset = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string currentUser = Application.SBO_Application.Company.UserName;
                    string query = string.Empty;

                    // MSSQL için sorgu
                    if (company.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        query = $"SELECT Department FROM OUSR WHERE USER_CODE = '{currentUser}'";
                    }
                    else
                    {
                        // HANA için sorgu; tablo ve alan isimleri büyük harfle tanımlanmışsa 
                        // çift tırnak kullanılması gerekebilir.
                        query = $"SELECT \"Department\" FROM \"OUSR\" WHERE \"USER_CODE\" = '{currentUser}'";
                    }

                    recordset.DoQuery(query);

                    if (recordset.RecordCount > 0)
                    {
                        int userDeptCode = Convert.ToInt32(recordset.Fields.Item("Department").Value);
                        if (userDeptCode != 3)
                        {
                            DisableProjectButton();
                        }
                    }
                    else
                    {
                        DisableProjectButton();
                    }
                }
                finally
                {
                    if (recordset != null)
                    {
                        Marshal.ReleaseComObject(recordset);
                        recordset = null;
                        GC.Collect();
                    }
                }
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Departman kontrolü hatası: {ex.Message}\n{ex.StackTrace}\n");
            }
        }

        /// <summary>
        /// Proje oluşturma butonunu devre dışı bırakır.
        /// </summary>
        private void DisableProjectButton()
        {
            try
            {
                SAPbouiCOM.Form form = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
                SAPbouiCOM.Item btnPrjItem = form.Items.Item("btn_prj");
                btnPrjItem.Enabled = false;
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Projeye geçiş butonu devre dışı bırakma hatası: {ex.Message}\n{ex.StackTrace}\n");
            }
        }

        #endregion

        #region Yardımcı Metodlar

        private SAPbobsCOM.Company GetCompany()
        {
            return (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
        }

        private void ShowMessage(string message)
        {
            Application.SBO_Application.MessageBox(message);
        }

        private void ShowStatusMessage(string message, BoStatusBarMessageType messageType)
        {
            Application.SBO_Application.StatusBar.SetText(message, BoMessageTime.bmt_Short, messageType);
        }

        #endregion
    }
}
