using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjeTablosu
{
    class MessagesHelper
    {
       public static void SendMessage(
       SAPbobsCOM.Company company,
       string recipientUserCode,
       string subject,
       string body,
       string linkTable = null,
       string linkKey = null)
        {
            var cmpSrv = company.GetCompanyService();
            var messageSrv = (SAPbobsCOM.MessagesService)
                cmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService);

            var message = (SAPbobsCOM.Message)
                messageSrv.GetDataInterface(
                    SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage);

            message.Subject = subject;
            message.Text = body;

 
            var rc = message.RecipientCollection;
            rc.Add();
            rc.Item(0).UserCode = recipientUserCode;
            rc.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
            rc.Item(0).SendEmail = SAPbobsCOM.BoYesNoEnum.tNO;

            if (!string.IsNullOrEmpty(linkTable) && !string.IsNullOrEmpty(linkKey))
            {
                var cols = message.MessageDataColumns;
                var col = cols.Add();
                col.ColumnName = "Doküman";        
                col.Link = SAPbobsCOM.BoYesNoEnum.tYES;
                var line = col.MessageDataLines.Add();
                line.Value = linkKey;      
                line.Object = linkTable;        
                line.ObjectKey = linkKey;
            }

            // 5) Gönder
            messageSrv.SendMessage(message);
        }
    }
}
