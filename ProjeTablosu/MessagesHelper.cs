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
            // 1) CompanyService ve MessagesService’e erişim
            var cmpSrv = company.GetCompanyService();
            var messageSrv = (SAPbobsCOM.MessagesService)
                cmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService);

            // 2) Message nesnesi oluşturma
            var message = (SAPbobsCOM.Message)
                messageSrv.GetDataInterface(
                    SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage);

            message.Subject = subject;
            message.Text = body;

            // 3) Alıcı tanımlama
            var rc = message.RecipientCollection;
            rc.Add();
            rc.Item(0).UserCode = recipientUserCode;
            rc.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
            rc.Item(0).SendEmail = SAPbobsCOM.BoYesNoEnum.tYES;

            // 4) (İsteğe bağlı) Mesajın içinde tıklanabilir link eklemek
            if (!string.IsNullOrEmpty(linkTable) && !string.IsNullOrEmpty(linkKey))
            {
                var cols = message.MessageDataColumns;
                var col = cols.Add();
                col.ColumnName = "Doküman";        // istediğiniz başlık
                col.Link = SAPbobsCOM.BoYesNoEnum.tYES;
                var line = col.MessageDataLines.Add();
                line.Value = linkKey;          // örn. DocEntry veya U_DocNum
                line.Object = linkTable;        // SAP obje tipi (ör. "17" = SalesOrder, UDO için 13900000xx)
                line.ObjectKey = linkKey;
            }

            // 5) Gönder
            messageSrv.SendMessage(message);
        }
    }
}
