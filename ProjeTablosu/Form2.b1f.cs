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
        private EditText txt_reg;
        private Button btn_fltr;
        private EditText txt_regbit;
        private CheckBox check_prj;

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
            //  Form üzerindeki kontrol referanslarını alıyoruz.
            this.btn_prj = ((SAPbouiCOM.Button)(this.GetItem("btn_prj").Specific));
            this.grd_liste = ((SAPbouiCOM.Grid)(this.GetItem("grd_liste").Specific));
            this.txt_reg = ((SAPbouiCOM.EditText)(this.GetItem("txt_reg").Specific));
            this.btn_fltr = ((SAPbouiCOM.Button)(this.GetItem("btn_fltr").Specific));
            this.txt_regbit = ((SAPbouiCOM.EditText)(this.GetItem("txt_regbit").Specific));
            this.check_prj = ((SAPbouiCOM.CheckBox)(this.GetItem("check_prj").Specific));
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents() { }

        /// <summary>
        /// Özel başlangıç işlemleri: DataTable oluşturma, sorgu çalıştırma ve grid’e atama.
        /// </summary>
        private void OnCustomInitialize()
        {
            this.btn_prj.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.OnProjectButtonClickBefore);
            this.btn_fltr.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnFilter_ClickBefore);

            // SAP B1'in EditText kontrolü, tarih alanlarında sistemin beklediği formatı (örneğin "yyyyMMdd") kullanır.
            // Bu nedenle EditText'e atama "yyyyMMdd" formatında yapılıyor.
            this.txt_reg.Value = DateTime.Today.ToString("yyyyMMdd");

            // DataTable oluşturularak grid'e bağlanıyor.
            SAPbouiCOM.DataTable oDataTable = this.UIAPIRawForm.DataSources.DataTables.Add("MyTable");
            // BuildQuery metoduna EditText'ten gelen değeri gönderiyoruz.
            oDataTable.ExecuteQuery(BuildQuery(this.txt_reg.Value));
            this.grd_liste.DataTable = oDataTable;
        }

        #endregion

        #region SQL Sorgusu Oluşturma ve Loglama

        /// <summary>
        /// SQL sorgusunu oluşturur.
        /// Not: EditText'ten gelen tarih değeri "yyyyMMdd" formatında olduğundan,
        /// bunu DateTime üzerinden "dd.MM.yyyy" formatına çeviriyoruz. Çünkü SQL'de
        /// CONVERT(VARCHAR(10), U_RegDate, 104) ifadesi "dd.MM.yyyy" döndürür.
        /// </summary>
        /// <param name="regDateFilter">EditText'ten alınan kayıt tarihi değeri (yyyyMMdd formatında).</param>
        /// <returns>Oluşturulan SQL sorgusu.</returns>
        private string BuildQuery(string regDateFilter = "")
        {
            string formattedDate = "";
            if (!string.IsNullOrEmpty(regDateFilter))
            {
                DateTime dt;
                if (DateTime.TryParseExact(regDateFilter, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                {
                    // SQL sorgusunun beklediği format
                    formattedDate = dt.ToString("dd.MM.yyyy");
                }
                else
                {
                    formattedDate = regDateFilter;
                }
            }

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

            if (!string.IsNullOrEmpty(formattedDate))
            {
                query.Append(" WHERE CONVERT(VARCHAR(10), U_RegDate, 104) = '" + formattedDate + "'");
            }

            // SQL sorgusu log dosyasına kaydediliyor.
            LogQuery(query.ToString());
            return query.ToString();
        }

        /// <summary>
        /// SQL sorgusunu loglar: Hem Trace çıktısına hem de Logs klasöründeki dosyaya.
        /// </summary>
        /// <param name="query">Loglanacak sorgu metni.</param>
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
                // Log dosyası adı: SQLQueryLog_YYYYMMDD.txt
                string fileName = System.IO.Path.Combine(logPath, "SQLQueryLog_" + DateTime.Now.ToString("yyyyMMdd") + ".txt");
                string logRecord = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + query + Environment.NewLine;
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
        /// txt_reg alanına girilen tarih değeriyle filtre uygulanması için btn_fltr butonuna tıklanıldığında çalışır.
        /// </summary>
        private void BtnFilter_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                string filterDate = txt_reg.Value.Trim();
                SAPbouiCOM.DataTable oDataTable = this.UIAPIRawForm.DataSources.DataTables.Item("MyTable");
                oDataTable.ExecuteQuery(BuildQuery(filterDate));
            }
            catch (Exception ex)
            {
                LogError("Error filtering records", ex);
                BubbleEvent = false;
            }
        }

        /// <summary>
        /// btn_prj butonuna tıklanıldığında, seçili satır verileri ile yeni proje oluşturur.
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

            // Grid’den değerleri alıyoruz.
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

                // Kullanıcı tanımlı alan U_ProDocEntry'nin değeri atanıyor.
                for (int i = 0; i < newProject.UserFields.Count; i++)
                {
                    if (newProject.UserFields.Item(i).Name == "U_ProDocEntry")
                    {
                        newProject.UserFields.Item(i).Value = docEntry;
                        break;
                    }
                }

                // Kayıt tarihi alanı kontrol edilip, uygun formatta parse ediliyor.
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
                        ShowMessage("Kayıt Tarihi geçerli bir formatta değil: " + regDateStr);
                        return;
                    }
                }

                // İstenilen tarih alanı kontrol edilip, parse ediliyor.
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

                // Yeni proje SAP’ye ekleniyor.
                SAPbobsCOM.PM_ProjectDocumentParams newProjectParams = projectService.AddProject(newProject);
                int newProjectAbsEntry = newProjectParams.AbsEntry;

                recordset = (SAPbobsCOM.Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string updateQuery = $"UPDATE OPMG SET U_ProDocEntry = '{docEntry}' WHERE AbsEntry = {newProjectAbsEntry}";
                // Gerekirse update sorgusu çalıştırılabilir: recordset.DoQuery(updateQuery);
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
