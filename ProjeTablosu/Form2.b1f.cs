using SAPbouiCOM.Framework;
using SAPbouiCOM;
using System;
using System.Text;
using SAPbobsCOM; // DI API için kullanılıyor
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization; // CultureInfo için
using Application = SAPbouiCOM.Framework.Application;

namespace ProjeTablosu
{
    [FormAttribute("ProjeTablosu.Form2", "Form2.b1f")]
    class Form2 : UserFormBase
    {
        #region Form Kontrolleri

        private Button btn_prj;
        private Grid grd_liste;

        #endregion

        #region Sabitler

        /// <summary>
        /// Tarih için kabul edilebilir formatlar. Gelen veri 
        /// örneğin "15.04.2025", "15.04.2025 00:00:00", "20250415" vb. formatlarda olabilir.
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
            // Form üzerindeki kontrol referanslarını al
            this.btn_prj = ((SAPbouiCOM.Button)(this.GetItem("btn_prj").Specific));
            this.grd_liste = ((SAPbouiCOM.Grid)(this.GetItem("grd_liste").Specific));

            OnCustomInitialize();
        }

        public override void OnInitializeFormEvents() { }

        /// <summary>
        /// Özel başlangıç işlemleri: DataTable oluşturma, sorgu çalıştırma ve grid’e atama.
        /// </summary>
        private void OnCustomInitialize()
        {
            this.btn_prj.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.OnProjectButtonClickBefore);

            SAPbouiCOM.DataTable oDataTable = this.UIAPIRawForm.DataSources.DataTables.Add("MyTable");
            oDataTable.ExecuteQuery(BuildQuery());
            this.grd_liste.DataTable = oDataTable;
        }

        #endregion

        #region SQL Sorgusu Oluşturma ve Loglama

        /// <summary>
        /// SQL sorgusunu oluşturur.
        /// </summary>
        /// <returns>Oluşturulan SQL sorgusu.</returns>
        private string BuildQuery()
        {
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

            LogQuery(query.ToString());
            return query.ToString();
        }

        /// <summary>
        /// SQL sorgusunu loglar.
        /// </summary>
        /// <param name="query">Loglanacak sorgu metni.</param>
        private void LogQuery(string query)
        {
            Trace.WriteLine("SQL Sorgusu: " + query);
        }

        #endregion

        #region Button Event Handlers

        /// <summary>
        /// Butona tıklanma olayında, seçili satır verilerini alıp yeni proje oluşturma methodunu çağırır.
        /// </summary>
        private void OnProjectButtonClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.Form form = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.Item btnPrjItem = form.Items.Item("btn_prj");

                // Buton devre dışıysa, işlem iptal edilir.
                if (!btnPrjItem.Enabled)
                {
                    BubbleEvent = false;
                    return;
                }

                // Seçili satır verileri ile yeni proje oluştur.
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
        /// Grid’de seçili satırdaki verileri kullanarak SAP Project Management üzerinde yeni proje oluşturur.
        /// Tarih alanlarının formatlarını da kontrol eder.
        /// </summary>
        private void CreateNewProject()
        {
            // Grid’de seçili en az bir satır var mı kontrol et
            if (this.grd_liste.Rows.SelectedRows.Count == 0)
            {
                ShowMessage("Lütfen yeni proje oluşturmak için bir satır seçiniz.");
                return;
            }

            // Seçili satır index bilgisini al
            int selectedRow = this.grd_liste.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder);
            SAPbouiCOM.DataTable dt = this.grd_liste.DataTable;
            // İşlem başlamadan önce ekranı donduruyoruz
            this.UIAPIRawForm.Freeze(true);

            // Grid’den verileri çek; alias’ların (örn. "Kayıt Tarihi") sorgudaki isimle birebir aynı olduğundan emin olun.
            string docEntry = dt.GetValue("Döküman Numarası", selectedRow)?.ToString().Trim() ?? "";
            string projectTitle = dt.GetValue("Proje Talebi Tanımı", selectedRow)?.ToString().Trim() ?? "";
            string regDateStr = dt.GetValue("Kayıt Tarihi", selectedRow)?.ToString().Trim() ?? "";
            string delDateStr = dt.GetValue("İstenilen Tarih", selectedRow)?.ToString().Trim() ?? "";

            // Debug amaçlı loglama
            Trace.WriteLine("DEBUG - Kayıt Tarihi: " + regDateStr);
            Trace.WriteLine("DEBUG - İstenilen Tarih: " + delDateStr);

            SAPbobsCOM.Company company = GetCompany();
            SAPbobsCOM.CompanyService companyService = null;
            SAPbobsCOM.ProjectManagementService projectService = null;
            SAPbobsCOM.Recordset recordset = null;

            try
            {
                // Şirket ve proje yönetim servislerini al
                companyService = company.GetCompanyService();
                projectService = (SAPbobsCOM.ProjectManagementService)companyService.GetBusinessService(ServiceTypes.ProjectManagementService);

                // Yeni proje dokümanı oluştur
                SAPbobsCOM.PM_ProjectDocumentData newProject =
                    (SAPbobsCOM.PM_ProjectDocumentData)projectService.GetDataInterface(ProjectManagementServiceDataInterfaces.pmsPM_ProjectDocumentData);

                // Proje adını, seçili satırdaki proje talebi tanımı ve döküman numarası ile birleştir.
                string projectName = $"{projectTitle} {docEntry}";
                newProject.ProjectName = projectName;

                // UserField (örn. U_ProDocEntry) değerini döküman numarası ile doldur.
                for (int i = 0; i < newProject.UserFields.Count; i++)
                {
                    if (newProject.UserFields.Item(i).Name == "U_ProDocEntry")
                    {
                        newProject.UserFields.Item(i).Value = docEntry;
                        break;
                    }
                }

                // Kayıt Tarihi verisini, tanımlı formatlardan biriyle parse et
                if (!string.IsNullOrEmpty(regDateStr))
                {
                    DateTime startDate;
                    bool parsed = DateTime.TryParseExact(
                        regDateStr,
                        ALLOWED_DATE_FORMATS,
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out startDate);

                    if (parsed)
                        newProject.StartDate = startDate;
                    else
                    {
                        // Eğer parse edilemezse, kullanıcıya mesaj ver ve işlemi sonlandır.
                        ShowMessage("Kayıt Tarihi geçerli bir formatta değil: " + regDateStr);
                        return;
                    }
                }

                // İstenilen Tarih verisini, tanımlı formatlardan biriyle parse et
                if (!string.IsNullOrEmpty(delDateStr))
                {
                    DateTime dueDate;
                    bool parsed = DateTime.TryParseExact(
                        delDateStr,
                        ALLOWED_DATE_FORMATS,
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out dueDate);

                    if (parsed)
                        newProject.DueDate = dueDate;
                    else
                    {
                        ShowMessage("İstenilen Tarih geçerli bir formatta değil: " + delDateStr);
                        return;
                    }
                }

                // Projeyi SAP'ye ekle
                SAPbobsCOM.PM_ProjectDocumentParams newProjectParams = projectService.AddProject(newProject);
                int newProjectAbsEntry = newProjectParams.AbsEntry;

                // İstendiğinde, ek güncelleme sorgusu çalıştırılabilir
                recordset = (SAPbobsCOM.Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string updateQuery = $"UPDATE OPMG SET U_ProDocEntry = '{docEntry}' WHERE AbsEntry = {newProjectAbsEntry}";
                // recordset.DoQuery(updateQuery); // Güncelleme yapılacaksa yorum satırından çıkarın
                dt.SetValue("Projeye Dönüştürüldü", selectedRow, "Evet");
                this.UIAPIRawForm.Freeze(false);

                ShowStatusMessage($"Proje başarıyla oluşturuldu. Proje Kodu: {newProjectAbsEntry}", BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                LogError("Error creating new project", ex);
                ShowMessage("Yeni proje oluşturulurken hata oluştu: " + ex.Message);
            }
            finally
            {
                // COM nesnelerini serbest bırak, bellek yönetimi için
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

        /// <summary>
        /// SAP B1 DI API üzerinden Company nesnesini döner.
        /// </summary>
        /// <returns>Company nesnesi.</returns>
        private SAPbobsCOM.Company GetCompany()
        {
            return (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
        }

        /// <summary>
        /// Kullanıcıya mesaj kutusu ile mesaj gösterir.
        /// </summary>
        private void ShowMessage(string message)
        {
            Application.SBO_Application.MessageBox(message);
        }

        /// <summary>
        /// Status bar üzerinde mesaj gösterir.
        /// </summary>
        private void ShowStatusMessage(string message, BoStatusBarMessageType messageType)
        {
            Application.SBO_Application.StatusBar.SetText(message, BoMessageTime.bmt_Short, messageType);
        }

        /// <summary>
        /// Hata mesajlarını loglar.
        /// </summary>
        private void LogError(string message, Exception ex)
        {
            string errorMessage = $"{message}: {ex.Message}";
            Trace.WriteLine(errorMessage);
            Trace.WriteLine(ex.StackTrace);
        }

        #endregion
    }
}
