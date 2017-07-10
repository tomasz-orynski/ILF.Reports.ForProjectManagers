using BlueBit.ILF.Reports.ForProjectManagers.Diagnostics;
using MoreLinq;
using System.Collections.Generic;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.Reports.ForProjectManagers
{
    partial class Program
    {
        class SendData
        {
            public string ID { get; set; }
            public string Title { get; set; }
            public string MsgBody { get; set; }
            public string AddressTo { get; set; }
            public string AttachmentPath { get; set; }
        }

        private static void SendReportData(string pathSend, IEnumerable<SendData> items)
            => _logger.OnEntryCall(() =>
            {
                var app = new Outlook.Application();
                items.ForEach(item =>
                {
                    var tmpFile = Path.Combine(pathSend, Path.Combine(item.Title + "." + Path.GetExtension(item.AttachmentPath)));
                    File.Copy(item.AttachmentPath, tmpFile);

                    _logger.Info($"SEND BEG: #[{item.ID}] to [{item.AddressTo}].");
                    var mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                    mailItem.Subject = item.Title;
                    mailItem.To = item.AddressTo;
                    mailItem.Body = item.MsgBody;
                    var attachment = mailItem.Attachments.Add(tmpFile, DisplayName: item.Title);
                    attachment.DisplayName = item.Title;
                    mailItem.Save();

                    File.Delete(tmpFile);
                    _logger.Info($"SEND END: #[{item.ID}].");
                });
            });
    }
}
