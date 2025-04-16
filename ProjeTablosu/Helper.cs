using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace ProjeTablosu
{
    public static class Helper
    {
        public static void OpenActiveLog()
        {
            string logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ProjeTablosu", "logs");
            string logFileName = DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            string logFilePath = Path.Combine(logDirectory, logFileName);
            if (!File.Exists(logFilePath))
            {
                ProjeTablosu.Program.SBO_Application.StatusBar.SetText($"Log dosyası bulunamadı", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }
            Process.Start(new ProcessStartInfo(logFilePath) { UseShellExecute = true });
        }

        public static void LogToFile(string logMessage)
        {
            string logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ProjeTablosu", "logs");
            try
            {
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }
                string logFileName = DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
                string logFilePath = Path.Combine(logDirectory, logFileName);
                File.AppendAllText(logFilePath, logMessage);
            }
            catch { }
        }

        public static string LoadSqlScript(string fileName, SAPbobsCOM.BoDataServerTypes dbServerType)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string resourceName = dbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB
                ? $"ProjeTablosu.Scripts.HANA.{fileName}"
                : $"ProjeTablosu.Scripts.MSSQL.{fileName}";
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    throw new FileNotFoundException($"SQL script file not found: {resourceName}");
                using (StreamReader reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public static string ReplaceVariables(string sqlScript, Dictionary<string, string> tobeReplace)
        {
            foreach (var item in tobeReplace)
            {
                sqlScript = sqlScript.Replace(item.Key, item.Value);
            }
            return sqlScript;
        }

        public static string AppendConditions(string sqlScript, string additionalCondition, string orderScript, SAPbobsCOM.BoDataServerTypes dbServerType)
        {
            if (!string.IsNullOrEmpty(additionalCondition))
            {
                sqlScript += " AND " + additionalCondition;
            }
            if (!string.IsNullOrEmpty(orderScript))
            {
                orderScript = orderScript != "" ? orderScript + ".txt" : "";
                string orderScriptContent = LoadSqlScript(orderScript, dbServerType);
                sqlScript += " " + orderScriptContent;
            }
            return sqlScript;
        }

        public static SAPbobsCOM.Recordset executeSQLFromFile(string fileName, Dictionary<string, string> tobeReplace, string AdditionalCondition, string orderScript)
        {
            SAPbobsCOM.Company bobsCompany = (SAPbobsCOM.Company)ProjeTablosu.Program.oCompany;
            SAPbobsCOM.Recordset tmpRecordSet = (SAPbobsCOM.Recordset)ProjeTablosu.Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string sqlScript = LoadSqlScript(fileName + ".sql", bobsCompany.DbServerType);
                sqlScript = ReplaceVariables(sqlScript, tobeReplace);
                sqlScript = AppendConditions(sqlScript, AdditionalCondition, orderScript, bobsCompany.DbServerType);
                LogToFile($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Executing Query:\n{sqlScript}\n");
                tmpRecordSet.DoQuery(sqlScript);
                LogToFile($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Query executed successfully. Row Count: {tmpRecordSet.RecordCount}\n");
            }
            catch (Exception e)
            {
                LogToFile($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Error Occurred:\n{e.Message}\n{e.StackTrace}\n");
                ProjeTablosu.Program.SBO_Application.StatusBar.SetText($"SQL Sorgusu Çalıştırılamadı. Hata:{e.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            return tmpRecordSet;
        }
    }
}
