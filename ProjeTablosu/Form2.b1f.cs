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
        private Button btn_red;
        private Grid grd_liste;
        private EditText txt_reg;      // Başlangıç tarihi
        private EditText txt_regbit;   // Bitiş tarihi
        private Button btn_fltr;
        private ComboBox cmb_filter;   // ComboBox filtresi
        private EditText txt_reject;
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
            "dd/MM/yyyy HH:mm:ss",
            "dd-MMM-yy hh:mm:ss tt",    // örn: 23-Apr-25 12:00:00 AM
            "dd-MMM-yyyy hh:mm:ss tt"   // örn: 23-Apr-2025 12:00:00 AM
        };

        private const string UDO_PROJECT_TABLE = "@PROJECT";
        private const string UDO_PROJECT_ROWS_TABLE = "@PROJECTROW";

        // ComboBox değerleri için sabitler
        private const int FILTER_CONVERTED = 0;      // Projeye Dönüştürüldü
        private const int FILTER_NOT_CONVERTED = 1;  // Projeye Dönüştürülmedi
        private const int FILTER_PENDING = 2;
        private const int FILTER_ALL = 3;            // Hepsi
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
                this.cmb_filter = ((SAPbouiCOM.ComboBox)(this.GetItem("cmb_filter").Specific));
                this.btn_red = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
                this.txt_reject = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));

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
        /// Özel başlangıç işlemleri: DataTable oluşturma, sorgu çalıştırma ve grid'e atama.
        /// Form açılışında tarih aralığı uygulanır ve tüm kayıtlar gösterilir (varsayılan filtre: Hepsi).
        /// </summary>
        private void OnCustomInitialize()
        {
            try
            {
                this.grd_liste.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
                this.btn_prj.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.OnProjectButtonClickBefore);
                this.btn_fltr.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnFilter_ClickBefore);
                this.btn_red.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnRed_ClickBefore);

                // EditText'lere varsayılan değer ataması ("yyyyMMdd" formatında)
                this.txt_reg.Value = DateTime.Today.ToString("yyyyMMdd");
                this.txt_regbit.Value = DateTime.Today.ToString("yyyyMMdd");

                // ComboBox'a değerleri ekle (zaten UI'da tanımlanmış olabilir, ama güvenlik için)
                InitializeComboBox();

                // Varsayılan olarak "Hepsi" seçeneğini seç
                this.cmb_filter.Select(FILTER_ALL, BoSearchKey.psk_Index);

                // Form açılışında, ComboBox filtresi uygulanmadan sorgu çalışsın (tüm kayıtlar)
                SAPbouiCOM.DataTable oDataTable = this.UIAPIRawForm.DataSources.DataTables.Add("MyTable");
                string initialQuery = BuildQuery(this.txt_reg.Value, this.txt_regbit.Value, FILTER_ALL);
                oDataTable.ExecuteQuery(initialQuery);
                this.grd_liste.DataTable = oDataTable;

                //CheckUserDepartmentPermissions();
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Form2 OnCustomInitialize hatası: {ex.Message}\n{ex.StackTrace}\n");
            }
        }

        /// <summary>
        /// ComboBox'a filtreleme seçeneklerini ekler
        /// </summary>
        private void InitializeComboBox()
        {
            try
            {
                // ComboBox boş ise (runtime'da oluşturulmuşsa) değerleri ekleyelim
                if (cmb_filter.ValidValues.Count == 0)
                {
                    cmb_filter.ValidValues.Add(FILTER_CONVERTED.ToString(), "Projeye Dönüştürüldü");
                    cmb_filter.ValidValues.Add(FILTER_NOT_CONVERTED.ToString(), "Reddedildi");
                    cmb_filter.ValidValues.Add(FILTER_PENDING.ToString(), "Beklemede");
                    cmb_filter.ValidValues.Add(FILTER_ALL.ToString(), "Hepsi");
                }
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"ComboBox başlatma hatası: {ex.Message}\n{ex.StackTrace}\n");
            }
        }

        #endregion

        #region SQL Sorgusu Oluşturma ve Loglama

        /// <summary>
        /// SQL sorgusunu oluşturur.
        /// </summary>
        /// <param name="startDateFilter">Başlangıç tarihi (yyyyMMdd formatı)</param>
        /// <param name="endDateFilter">Bitiş tarihi (yyyyMMdd formatı)</param>
        /// <param name="filterOption">Filtreleme seçeneği (0: Dönüştürüldü, 1: Dönüştürülmedi, 2: Hepsi)</param>
        private string BuildQuery(string startDateFilter, string endDateFilter, int filterOption)
        {
            // 1) Şirket bağlantısı
            SAPbobsCOM.Company company = GetCompany();

            // 2) Base query'yi yükle
            string baseQuery = company.DbServerType != BoDataServerTypes.dst_HANADB
                ? Helper.LoadSqlScript("SelectProjectList.sql", BoDataServerTypes.dst_MSSQL)
                : Helper.LoadSqlScript("SelectProjectListHana.sql", BoDataServerTypes.dst_HANADB);

            // 3) Tarih filtresini oluştur
            string dateCondition = "";
            DateTime dtStart, dtEnd;
            if (!string.IsNullOrEmpty(startDateFilter) && DateTime.TryParseExact(startDateFilter, "yyyyMMdd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out dtStart)
                && !string.IsNullOrEmpty(endDateFilter) && DateTime.TryParseExact(endDateFilter, "yyyyMMdd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out dtEnd))
            {
                string s = dtStart.ToString("dd.MM.yyyy");
                string e = dtEnd.ToString("dd.MM.yyyy");
                if (company.DbServerType != BoDataServerTypes.dst_HANADB)
                    dateCondition = $"AND U_RegDate BETWEEN CONVERT(DATETIME, '{s}', 104) AND CONVERT(DATETIME, '{e}', 104) ";
                else
                    dateCondition = $"AND \"U_RegDate\" BETWEEN TO_DATE('{s}', 'DD.MM.YYYY') AND TO_DATE('{e}', 'DD.MM.YYYY') ";
            }
            else if (!string.IsNullOrEmpty(startDateFilter) && DateTime.TryParseExact(startDateFilter, "yyyyMMdd",
                         CultureInfo.InvariantCulture, DateTimeStyles.None, out dtStart))
            {
                string s = dtStart.ToString("dd.MM.yyyy");
                if (company.DbServerType != BoDataServerTypes.dst_HANADB)
                    dateCondition = $"AND U_RegDate >= CONVERT(DATETIME, '{s}', 104) ";
                else
                    dateCondition = $"AND \"U_RegDate\" >= TO_DATE('{s}', 'DD.MM.YYYY') ";
            }
            else if (!string.IsNullOrEmpty(endDateFilter) && DateTime.TryParseExact(endDateFilter, "yyyyMMdd",
                         CultureInfo.InvariantCulture, DateTimeStyles.None, out dtEnd))
            {
                string e = dtEnd.ToString("dd.MM.yyyy");
                if (company.DbServerType != BoDataServerTypes.dst_HANADB)
                    dateCondition = $"AND U_RegDate <= CONVERT(DATETIME, '{e}', 104) ";
                else
                    dateCondition = $"AND \"U_RegDate\" <= TO_DATE('{e}', 'DD.MM.YYYY') ";
            }

            // 4) U_IsConverted filtresini oluştur (MSSQL vs HANA ayrımı)
            string convertedCondition;
            if (company.DbServerType != BoDataServerTypes.dst_HANADB)
            {
                // MSSQL
                switch (filterOption)
                {
                    case FILTER_CONVERTED:
                        // Sadece Y olanları göster
                        convertedCondition = "AND U_IsConverted = 'Y' ";
                        break;
                    case FILTER_NOT_CONVERTED:
                        // Sadece Y olmayanları veya NULL olanları göster
                        convertedCondition = "AND U_IsConverted = 'N' ";
                        break;
                    case FILTER_PENDING:
                        convertedCondition = "AND U_IsConverted = 'P' ";
                        break;
                    case FILTER_ALL:
                        // Tümünü göster - filtre yok
                        convertedCondition = "";
                        break;
                    default:
                        // Varsayılan olarak tümünü göster
                        convertedCondition = "";
                        break;
                }
            }
            else
            {
                // HANA
                switch (filterOption)
                {
                    case FILTER_CONVERTED:
                        // Sadece Y olanları göster
                        convertedCondition = "AND \"U_IsConverted\" = 'Y' ";
                        break;
                    case FILTER_NOT_CONVERTED:
                        // Sadece Y olmayanları veya NULL olanları göster
                        convertedCondition = "AND \"U_IsConverted\" = 'N' ";
                        break;
                    case FILTER_PENDING:
                        // Sadece Y olmayanları veya NULL olanları göster
                        convertedCondition = "AND \"U_IsConverted\" = 'P' ";
                        break;
                    case FILTER_ALL:
                        // Tümünü göster - filtre yok
                        convertedCondition = "";
                        break;
                    default:
                        // Varsayılan olarak tümünü göster
                        convertedCondition = "";
                        break;
                }
            }

            // 5) Placeholder'ları değiştir
            string finalQuery = baseQuery
                .Replace("--DATEFILTER--", dateCondition)
                .Replace("--CONVERTEDFILTER--", convertedCondition);

            // 6) Logla ve döndür
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

                // ComboBox'tan seçilen değeri al
                int filterOption = Convert.ToInt32(cmb_filter.Value);

                // Log the values to check
                Helper.LogToFile($"ComboBox değeri - filterOption: {filterOption}");

                SAPbouiCOM.DataTable oDataTable = this.UIAPIRawForm.DataSources.DataTables.Item("MyTable");
                string query = BuildQuery(startFilter, endFilter, filterOption);

                // Sorguyu logla
                Helper.LogToFile($"Executing Filter Query: {query}");

                // Sorguyu çalıştır
                oDataTable.ExecuteQuery(query);

                // Sonuç sayısını logla
                Helper.LogToFile($"Query returned {oDataTable.Rows.Count} rows");
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


        private void BtnRed_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // 1) Form ve Grid referanslarını al
            var form = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            var grid = (SAPbouiCOM.Grid)form.Items.Item("grd_liste").Specific;

            // 2) Satır seçimi kontrolü
            if (grid.Rows.SelectedRows.Count == 0)
            {
                ShowMessage("Lütfen önce reddedeceğiniz kaydı seçiniz.");
                BubbleEvent = false;
                return;
            }
            int row = grid.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder);

            // 3) Doküman numarası ve reddetme sebebi
            string docNum = grid.DataTable.GetValue("Döküman Numarası", row)?.ToString();
            string reason = this.txt_reject.Value?.Trim();
            if (string.IsNullOrEmpty(reason))
            {
                ShowMessage("Lütfen reddetme nedenini giriniz.");
                BubbleEvent = false;
                return;
            }
            string escapedReason = reason.Replace("'", "''");

            var company = GetCompany();

            // 4) UPDATE sorgusu ve COM nesnesinin yönetimi
            SAPbobsCOM.Recordset rsUpdate = null;
            try
            {
                rsUpdate = (SAPbobsCOM.Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string updateQuery = company.DbServerType != BoDataServerTypes.dst_HANADB
                    ? $"UPDATE [@PROJECT] SET U_Reject = '{escapedReason}', U_IsConverted = 'N' WHERE DocEntry = {docNum}"
                    : $"UPDATE \"@PROJECT\" SET \"U_Reject\" = '{escapedReason}', \"U_IsConverted\" = 'N' WHERE \"DocEntry\" = {docNum}";

                Helper.LogToFile($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - BtnRed_Click UPDATE query:\n{updateQuery}");
                rsUpdate.DoQuery(updateQuery);
                Helper.LogToFile($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - BtnRed_Click UPDATE executed for DocEntry {docNum}.");

                // Grid ve UI güncellemesi
                grid.DataTable.SetValue("Durum", row, "Reddedildi");
                form.Items.Item("btn_prj").Enabled = false;
                ShowStatusMessage("Proje reddedildi ve neden kaydedildi.", BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - BtnRed_Click UPDATE error:\n{ex.Message}\n{ex.StackTrace}");
                BubbleEvent = false;
                return;
            }
            finally
            {
                if (rsUpdate != null)
                    Marshal.ReleaseComObject(rsUpdate);
            }

            // 5) Talep sahibinin kullanıcı kodunu alma
            string requesterUserCode = "";
            SAPbobsCOM.Recordset rsUser = null;
            try
            {
                rsUser = (SAPbobsCOM.Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string userQuery = company.DbServerType != BoDataServerTypes.dst_HANADB
                    ? $"SELECT U_NAME FROM [@PROJECT] WHERE DocEntry = {docNum}"
                    : $"SELECT \"U_NAME\" FROM \"@PROJECT\" WHERE \"DocEntry\" = {docNum}";

                Helper.LogToFile($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - BtnRed_Click USER query:\n{userQuery}");
                rsUser.DoQuery(userQuery);

                int count = rsUser.RecordCount;
                string rawCode = (count > 0) ? rsUser.Fields.Item("U_NAME").Value.ToString() : "";

                Helper.LogToFile($"[BtnRed_Click] RecordCount: {count}, RawCode: '{rawCode}'");
                if (count == 0 || string.IsNullOrEmpty(rawCode))
                {
                    ShowStatusMessage(
                        "Reddedilen talep sahibinin kullanıcı kodu alınamadı. Lütfen U_NAME alanını kontrol edin.",
                        BoStatusBarMessageType.smt_Error);
                    return;
                }

                requesterUserCode = rawCode;
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - BtnRed_Click USER error:\n{ex.Message}\n{ex.StackTrace}");
                return;
            }
            finally
            {
                if (rsUser != null)
                    Marshal.ReleaseComObject(rsUser);
            }

            // 6) Bildirimi gönder
            try
            {
                string udoType = GetUdoObjectType(company, "UDOPROJECT");
                string subject = $"Proje Talebi #{docNum} Reddedildi";
                string body = $"{DateTime.Now:yyyy-MM-dd HH:mm} tarihinde proje talebiniz “Reddedildi” olarak güncellendi. Sebep: {reason}";

                MessagesHelper.SendMessage(
                    company,
                    requesterUserCode,
                    subject,
                    body,
                    linkTable: udoType,
                    linkKey: docNum);
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - BtnRed_Click SENDMSG error:\n{ex.Message}\n{ex.StackTrace}");
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
            string requesterUserCode = "";
            string udoType = "";
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

                dt.SetValue("Durum", selectedRow, "Onaylandı");
                var form = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
                form.Items.Item("btn_prj").Enabled = false;

                this.UIAPIRawForm.Freeze(false);

                ShowStatusMessage($"Proje başarıyla oluşturuldu. Proje Kodu: {newProjectAbsEntry}", BoStatusBarMessageType.smt_Success);

                string userSql = company.DbServerType != BoDataServerTypes.dst_HANADB
                         ? $"SELECT U_NAME FROM [@PROJECT] WHERE DocEntry = {docEntry}"
                         : $"SELECT \"U_NAME\" FROM \"@PROJECT\" WHERE \"DocEntry\" = {docEntry}";

                var rsUser = (SAPbobsCOM.Recordset)
                    company.GetBusinessObject(BoObjectTypes.BoRecordset);

                try
                {
                    rsUser.DoQuery(userSql);
                    int count = rsUser.RecordCount;
                    string rawCode = (count > 0)
                        ? rsUser.Fields.Item("U_NAME").Value.ToString()
                        : "";

                    if (count == 0 || string.IsNullOrEmpty(rawCode))
                    {
                        ShowStatusMessage(
                            "Reddedilen talep sahibinin kullanıcı kodu alınamadı. Lütfen U_NAME alanını kontrol edin.",
                            BoStatusBarMessageType.smt_Error);
                        return;
                    }

                    requesterUserCode = rawCode;
                }
                finally
                {
                    Marshal.ReleaseComObject(rsUser);
                }

                // 7) UDO ObjectType kodunu al
                udoType = GetUdoObjectType(company, "UDOPROJECT");

                // 8) Mesajı gönder
                string subject = $"Proje Talebi #{docEntry} Onaylandı";
                string body = $"{DateTime.Now:yyyy-MM-dd HH:mm} tarihinde proje talebiniz “Onaylandı” olarak işleme alındı. Yeni Proje Kodu: {newProjectAbsEntry}";

                MessagesHelper.SendMessage(
                    company,
                    requesterUserCode,
                    subject,
                    body,
                    linkTable: udoType,
                    linkKey: docEntry);
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Yeni proje oluşturma hatası: {ex.Message}\n{ex.StackTrace}\n");
                ShowMessage("Yeni proje oluşturulurken hata oluştu: " + ex.Message);
            }
            finally
            {
                // 9) DI API nesnelerini serbest bırak ve form kilidini aç
                if (projectService != null) Marshal.ReleaseComObject(projectService);
                if (companyService != null) Marshal.ReleaseComObject(companyService);
                this.UIAPIRawForm.Freeze(false);
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
        private void Grid0_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                var form = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
                var btnPrjItem = form.Items.Item("btn_prj");
                var grid = (SAPbouiCOM.Grid)sboObject;
                int row = pVal.Row;

                // Geçerli satır yoksa butonu kapat ve txt_reject'i temizle
                if (row < 0)
                {
                    btnPrjItem.Enabled = false;
                    this.txt_reject.Value = string.Empty;
                    return;
                }

                // “Projeye Dönüştürüldü” durumuna göre butonu ayarla
                string converted = grid.DataTable.GetValue("Durum", row).ToString();
                btnPrjItem.Enabled = !converted.Equals("Onaylandı", StringComparison.InvariantCultureIgnoreCase);

                // Satırın DocEntry değerini al
                string docNum = grid.DataTable.GetValue("Döküman Numarası", row).ToString();

                // DI API ile Recordset oluştur
                var company = GetCompany();
                var rs = (SAPbobsCOM.Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

                // MSSQL vs HANA ayrımıyla SELECT sorgusu
                string selectQuery;
                if (company.DbServerType != BoDataServerTypes.dst_HANADB)
                {
                    selectQuery = $"SELECT U_Reject FROM [@PROJECT] WHERE DocEntry = {docNum}";
                }
                else
                {
                    selectQuery = $"SELECT \"U_Reject\" FROM \"@PROJECT\" WHERE \"DocEntry\" = {docNum}";
                }

                // (İstersen loglamak için)
                Helper.LogToFile($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Grid0_ClickAfter SELECT:\n{selectQuery}\n");

                rs.DoQuery(selectQuery);

                // Sonucu txt_reject'e ata
                string reason = rs.RecordCount > 0
                    ? rs.Fields.Item(0).Value.ToString()
                    : string.Empty;
                this.txt_reject.Value = reason;

                // Temizlik
                Marshal.ReleaseComObject(rs);
            }
            catch (Exception ex)
            {
                Helper.LogToFile($"Grid0_ClickAfter hatası: {ex.Message}\n{ex.StackTrace}\n");
            }
        }
        /// <summary>
        /// @PROJECT gibi bir UDO tablosunun DI API ObjectType (Type) kodunu getirir.
        /// </summary>
        /// <param name="company">GetDICompany() ile alınan Company nesnesi</param>
        /// <param name="udoTableName">UDO tablosunun adı, örn. "PROJECT"</param>
        /// <returns>UDO ObjectType kodu, örn. "1390000084"</returns>
        private string GetUdoObjectType(SAPbobsCOM.Company company, string udoTableName)
        {
            // 1) UserObjectsMD business object’ını al
            var md = (UserObjectsMD)company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

            // 2) GetByKey parametresi UDO adını alır (tablodan “@” olmadan)
            bool found = md.GetByKey(udoTableName);
            if (!found)
                throw new Exception($"UDO metadata bulunamadı: {udoTableName}");

            // 3) Type property’si, sayısal ObjectType kodunu içerir
            string objectType = md.ObjectType.ToString();

            // 4) COM objesini serbest bırak
            System.Runtime.InteropServices.Marshal.ReleaseComObject(md);
            return objectType;
        }


    }
}