using BlueBit.ILF.Reports.ForProjectManagers.Diagnostics;
using MoreLinq;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.Reports.ForProjectManagers
{
    partial class Program
    {
        private static void SendReportData(IEnumerable<(string id, string path, string title, string info, string addr)> data)
            => _logger.OnEntryCall(() =>
            {
                var app = new Outlook.Application();
                data.ForEach(_ =>
                {
                    _logger.Info($"SEND BEG: #[{_.id}] to [{_.addr}].");
                    var mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                    mailItem.Subject = _.title;
                    mailItem.To = _.addr;
                    mailItem.Body = _.info;
                    var attachment = mailItem.Attachments.Add(_.path);
                    attachment.DisplayName = _.info;
                    //mailItem.Send();
                    mailItem.Save();
                    _logger.Info($"SEND END: #[{_.id}].");
                });
            });
    }
}
