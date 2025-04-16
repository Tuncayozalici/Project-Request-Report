using SAPbouiCOM.Framework;
using SAPbouiCOM;
using System;
using System.Text;
using SAPbobsCOM; // DI API için kullanılıyor
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;
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
        private CheckBox check_prj;    // CheckBox: Butona basınca durumuna göre filtre uygulanacak

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
            this.btn_prj = ((SAPbouiCOM.Button)(this.GetItem("btn_prj").Specific));
            this.grd_liste = ((SAPbouiCOM.Grid)(this.GetItem("grd_liste").Specific));
            this.grd_liste.DoubleClickBefore += new SAPbouiCOM._IGridEvents_DoubleClickBeforeEventHandler(this.grd_liste_DoubleClickBefore);
            this.txt_reg = ((SAPbouiCOM.EditText)(this.GetItem("txt_reg").Specific));
            this.txt_regbit = ((SAPbouiCOM.EditText)(this.GetItem("txt_regbit").Specific));
            this.btn_fltr = ((SAPbouiCOM.Button)(this.GetItem("btn_fltr").Specific));
            this.check_prj = ((SAPbouiCOM.CheckBox)(this.GetItem("check_prj").Specific));
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents() { }

        /// <summary>
        /// Özel başlangıç işlemleri: DataTable oluşturma, sorgu çalıştırma ve grid’e atama.
        /// Form açılışında tarih aralığı uygulanır fakat check box filtresi (U_IsConverted şartı)
        /// devreye alınmaz.
        /// </summary>
        private void OnCustomInitialize()
        {
            this.btn_prj.ClickBefore +=
                new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.OnProjectButtonClickBefore);
            this.btn_fltr.ClickBefore +=
                new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnFilter_ClickBefore);

            // EditText'lere varsayılan değer ataması ("yyyyMMdd" formatında)
            this.txt_reg.Value = DateTime.Today.ToString("yyyyMMdd");
            this.txt_regbit.Value = DateTime.Today.ToString("yyyyMMdd");

            // Form açılışında checkbox filtresi uygulanmadan sorgu çalışsın (onlyConverted = false)
            SAPbouiCOM.DataTable oDataTable = this.UIAPIRawForm.DataSources.DataTables.Add("MyTable");
            oDataTable.ExecuteQuery(BuildQuery(this.txt_reg.Value, this.txt_regbit.Value, false));
            this.grd_liste.DataTable = oDataTable;
        }

        #endregion

        #region SQL Sorgusu Oluşturma ve Loglama

        /// <summary>
        /// SQL sorgusunu oluşturur.
        /// EditText'lerden gelen tarih değerleri "yyyyMMdd" formatındadır; 
        /// bunlar önce DateTime'a çevrilip sonra "dd.MM.yyyy" formatına dönüştürülür.
        /// Sonrasında, U_RegDate alanı için BETWEEN koşulu oluşturulur.
        /// Checkbox değerine göre ek olarak U_IsConverted filtresi eklenir.
        /// </summary>
        private string BuildQuery(string startDateFilter, string endDateFilter, bool onlyConverted)
        {
            string formattedStart = "";
            string formattedEnd = "";

            // Başlangıç tarihi formatlama
            if (!string.IsNullOrEmpty(startDateFilter))
            {
                if (DateTime.TryParseExact(startDateFilter, "yyyyMMdd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dt))
                {
                    formattedStart = dt.ToString("dd.MM.yyyy");
                }
                else
                {
                    formattedStart = startDateFilter;
                }
            }

            // Bitiş tarihi formatlama
            if (!string.IsNullOrEmpty(endDateFilter))
            {
                if (DateTime.TryParseExact(endDateFilter, "yyyyMMdd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dt))
                {
                    formattedEnd = dt.ToString("dd.MM.yyyy");
                }
                else
                {
                    formattedEnd = endDateFilter;
                }
            }

            // Sorgu metnini oluşturma
            StringBuilder query = new StringBuilder();
            query.Append(@"
        SELECT 
            DocNum AS [Döküman Numarası],
            U_ProjectTitle AS [Proje Talebi Tanımı],
            U_NAME AS [Talep Eden Kullanıcı],
            U_Branch AS [Şube],
            U_Department AS [Departman],
            U_RegDate AS [Kayıt Tarihi],
            U_DelDate AS [İstenilen Tarih],
            CASE 
                WHEN U_IsConverted = 'Y' THEN 'Evet'
                ELSE 'Hayır'
            END AS [Projeye Dönüştürüldü]
        FROM [@PROJECT]");

            bool whereAdded = false;

            // Tarih aralığı filtresi
            if (!string.IsNullOrEmpty(formattedStart) && !string.IsNullOrEmpty(formattedEnd))
            {
                query.Append($" WHERE U_RegDate BETWEEN CONVERT(DATETIME, '{formattedStart}', 104) " +
                             $"AND CONVERT(DATETIME, '{formattedEnd}', 104) ");
                whereAdded = true;
            }
            else if (!string.IsNullOrEmpty(formattedStart))
            {
                query.Append($" WHERE U_RegDate >= CONVERT(DATETIME, '{formattedStart}', 104) ");
                whereAdded = true;
            }
            else if (!string.IsNullOrEmpty(formattedEnd))
            {
                query.Append($" WHERE U_RegDate <= CONVERT(DATETIME, '{formattedEnd}', 104) ");
                whereAdded = true;
            }

            // Checkbox'a göre U_IsConverted filtresi
            if (onlyConverted)
            {
                if (!whereAdded)
                {
                    query.Append(" WHERE ");
                    whereAdded = true;
                }
                else
                {
                    query.Append(" AND ");
                }
                query.Append(" U_IsConverted = 'Y' ");
            }
            else
            {
                if (!whereAdded)
                {
                    query.Append(" WHERE ");
                    whereAdded = true;
                }
                else
                {
                    query.Append(" AND ");
                }
                query.Append(" (U_IsConverted <> 'Y' OR U_IsConverted IS NULL) ");
            }

            // Sorguyu logla
            LogQuery(query.ToString());
            return query.ToString();
        }

        /// <summary>
        /// SQL sorgusunu loglar: Trace çıktısı ve Logs klasöründeki günlük dosyasına.
        /// </summary>
        private void LogQuery(string query)
        {
            Trace.WriteLine("SQL Sorgusu: " + query);
            try
            {
                string basePath = AppDomain.CurrentDomain.BaseDirectory;
                string logPath = System.IO.Path.Combine(basePath, "Logs");
                if (!System.IO.Directory.Exists(logPath))
                {
                    System.IO.Directory.CreateDirectory(logPath);
                }
                string fileName = System.IO.Path.Combine(logPath,
                    "SQLQueryLog_" + DateTime.Now.ToString("yyyyMMdd") + ".txt");
                string logRecord = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                                  + " - " + query + Environment.NewLine;
                System.IO.File.AppendAllText(fileName, logRecord);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Log kaydı yazılırken hata oluştu: " + ex.Message);
            }
        }

        #endregion

        #region Button Event Handlers

        /// <summary>
        /// Filtre butonuna tıklandığında, girilen kriterlere göre grid verilerini yeniler.
        /// </summary>
        private void BtnFilter_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                string startFilter = txt_reg.Value.Trim();
                string endFilter = txt_regbit.Value.Trim();
                bool onlyConverted = check_prj.Checked;

                SAPbouiCOM.DataTable oDataTable =
                    this.UIAPIRawForm.DataSources.DataTables.Item("MyTable");
                oDataTable.ExecuteQuery(BuildQuery(startFilter, endFilter, onlyConverted));
            }
            catch (Exception ex)
            {
                LogError("Error filtering records", ex);
                BubbleEvent = false;
            }
        }

        /// <summary>
        /// Proje oluştur butonuna tıklandığında, seçili satırla yeni proje oluşturur.
        /// </summary>
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
                LogError("Error in OnProjectButtonClickBefore", ex);
                BubbleEvent = false;
            }
        }

        #endregion

        #region Project Management

        /// <summary>
        /// Grid’de seçili satırdaki verilerle SAP Project Management üzerinde yeni proje oluşturur.
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
                projectService = (SAPbobsCOM.ProjectManagementService)companyService
                                 .GetBusinessService(ServiceTypes.ProjectManagementService);

                // Yeni proje için veri objesi
                SAPbobsCOM.PM_ProjectDocumentData newProject =
                    (SAPbobsCOM.PM_ProjectDocumentData)projectService
                    .GetDataInterface(ProjectManagementServiceDataInterfaces.pmsPM_ProjectDocumentData);

                // Proje adı
                string projectName = $"{projectTitle} {docEntry}";
                newProject.ProjectName = projectName;

                // U_ProDocEntry alanına doküman numarası
                for (int i = 0; i < newProject.UserFields.Count; i++)
                {
                    if (newProject.UserFields.Item(i).Name == "U_ProDocEntry")
                    {
                        newProject.UserFields.Item(i).Value = docEntry;
                        break;
                    }
                }

                // Kayıt tarihi
                if (!string.IsNullOrEmpty(regDateStr))
                {
                    if (DateTime.TryParseExact(regDateStr, ALLOWED_DATE_FORMATS,
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate))
                    {
                        newProject.StartDate = startDate;
                    }
                    else
                    {
                        ShowMessage("Kayıt Tarihi geçerli bir formatta değil: " + regDateStr);
                        return;
                    }
                }

                // İstenilen tarih
                if (!string.IsNullOrEmpty(delDateStr))
                {
                    if (DateTime.TryParseExact(delDateStr, ALLOWED_DATE_FORMATS,
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dueDate))
                    {
                        newProject.DueDate = dueDate;
                    }
                    else
                    {
                        ShowMessage("İstenilen Tarih geçerli bir formatta değil: " + delDateStr);
                        return;
                    }
                }

                // Proje ekleme
                SAPbobsCOM.PM_ProjectDocumentParams newProjectParams = projectService.AddProject(newProject);
                int newProjectAbsEntry = newProjectParams.AbsEntry;

                // İsteğe bağlı SQL güncellemesi (örnek)
                recordset = (SAPbobsCOM.Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string updateQuery = $"UPDATE OPMG SET U_ProDocEntry = '{docEntry}' WHERE AbsEntry = {newProjectAbsEntry}";
                // recordset.DoQuery(updateQuery); // Gerekirse

                // Grid üzerinde ilgili satırın "Projeye Dönüştürüldü" alanını 'Evet' yap
                dt.SetValue("Projeye Dönüştürüldü", selectedRow, "Evet");
                this.UIAPIRawForm.Freeze(false);

                ShowStatusMessage($"Proje başarıyla oluşturuldu. Proje Kodu: {newProjectAbsEntry}",
                                  BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                LogError("Error creating new project", ex);
                ShowMessage("Yeni proje oluşturulurken hata oluştu: " + ex.Message);
            }
            finally
            {
                // COM nesnelerini serbest bırakma
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

        private void LogError(string message, Exception ex)
        {
            string errorMessage = $"{message}: {ex.Message}";
            Trace.WriteLine(errorMessage);
            Trace.WriteLine(ex.StackTrace);
        }

        private void grd_liste_DoubleClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Grid clickedGrid = sboObject as SAPbouiCOM.Grid;
                if (clickedGrid == null) return;
                int rowIndex = pVal.Row;
                if (rowIndex < 0) return;

                string DocNum = clickedGrid.DataTable.GetValue("Döküman Numarası", rowIndex)?.ToString();
                if (string.IsNullOrEmpty(DocNum))
                {
                    throw new Exception("Döküman Numarası boş olamaz.");
                }

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
    }
}
