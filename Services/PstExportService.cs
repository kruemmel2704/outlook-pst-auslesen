using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using MimeKit;
using XstReader;

namespace Outlook_PST_auslesen.Services;

public static class PstExportService
{
    public static async Task ExportMessageAsync(ViewModels.MessageViewModel msgVm, string targetDir, bool exportAsEml)
    {
        var message = msgVm.Message;
        
        // Dateiname sicher machen
        string safeSubject = MakeSafeFilename(msgVm.Subject);
        string dateStr = msgVm.Date?.ToLocalTime().ToString("yyyy-MM-dd_HHmm") ?? "unbekannt";
        string baseName = $"{dateStr}_{safeSubject}";
        
        if (exportAsEml)
        {
            var mimeMessage = new MimeMessage();
            
            // Absender setzen
            if (!string.IsNullOrEmpty(msgVm.From))
            {
                if (MailboxAddress.TryParse(msgVm.From, out var fromAddr))
                {
                    mimeMessage.From.Add(fromAddr);
                }
                else
                {
                    mimeMessage.From.Add(new MailboxAddress(msgVm.From, "unknown@sender.com"));
                }
            }
            
            // Empfänger setzen
            if (!string.IsNullOrEmpty(msgVm.To))
            {
                var toList = msgVm.To.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (var to in toList)
                {
                    if (MailboxAddress.TryParse(to, out var toAddr))
                    {
                        mimeMessage.To.Add(toAddr);
                    }
                    else
                    {
                        mimeMessage.To.Add(new MailboxAddress(to, "unknown@recipient.com"));
                    }
                }
            }

            if (!string.IsNullOrEmpty(msgVm.Cc))
            {
                var ccList = msgVm.Cc.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (var cc in ccList)
                {
                    if (MailboxAddress.TryParse(cc, out var ccAddr))
                    {
                        mimeMessage.Cc.Add(ccAddr);
                    }
                    else
                    {
                        mimeMessage.Cc.Add(new MailboxAddress(cc, "unknown@cc.com"));
                    }
                }
            }

            mimeMessage.Subject = msgVm.Subject;
            mimeMessage.Date = msgVm.Date ?? DateTimeOffset.Now;

            var bodyBuilder = new BodyBuilder();
            string html = msgVm.BodyHtml;
            string text = msgVm.BodyText;

            if (!string.IsNullOrEmpty(html))
            {
                bodyBuilder.HtmlBody = html;
            }
            else
            {
                bodyBuilder.TextBody = text;
            }

            // Anhänge verarbeiten
            foreach (var attachment in message.Attachments)
            {
                if (attachment.IsFile)
                {
                    string tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                    try
                    {
                        attachment.SaveToFile(tempFile, attachment.LastModificationTime);
                        bodyBuilder.Attachments.Add(attachment.FileName, await File.ReadAllBytesAsync(tempFile));
                    }
                    finally
                    {
                        if (File.Exists(tempFile))
                        {
                            File.Delete(tempFile);
                        }
                    }
                }
            }

            mimeMessage.Body = bodyBuilder.ToMessageBody();
            string emlFilePath = Path.Combine(targetDir, $"{baseName}.eml");
            
            // Doppelte Dateinamen verhindern
            int counter = 1;
            while (File.Exists(emlFilePath))
            {
                emlFilePath = Path.Combine(targetDir, $"{baseName}_{counter++}.eml");
            }

            await using var stream = File.Create(emlFilePath);
            await mimeMessage.WriteToAsync(stream);
        }
        else
        {
            // Als Ordner-Struktur exportieren (HTML + Anhänge)
            string mailFolder = Path.Combine(targetDir, baseName);
            
            // Doppelte Ordnernamen verhindern
            int counter = 1;
            while (Directory.Exists(mailFolder))
            {
                mailFolder = Path.Combine(targetDir, $"{baseName}_{counter++}");
            }
            
            Directory.CreateDirectory(mailFolder);

            // Metadaten und Textkörper in Textdatei schreiben
            string metaFile = Path.Combine(mailFolder, "nachricht.txt");
            using (var writer = new StreamWriter(metaFile, false, System.Text.Encoding.UTF8))
            {
                await writer.WriteLineAsync($"Von: {msgVm.From}");
                await writer.WriteLineAsync($"An: {msgVm.To}");
                await writer.WriteLineAsync($"Cc: {msgVm.Cc}");
                await writer.WriteLineAsync($"Datum: {msgVm.DisplayDate}");
                await writer.WriteLineAsync($"Betreff: {msgVm.Subject}");
                await writer.WriteLineAsync(new string('-', 50));
                await writer.WriteLineAsync(msgVm.BodyText);
            }

            // HTML-Version schreiben, falls vorhanden
            if (!string.IsNullOrEmpty(msgVm.BodyHtml))
            {
                string htmlFile = Path.Combine(mailFolder, "nachricht.html");
                await File.WriteAllTextAsync(htmlFile, msgVm.BodyHtml, System.Text.Encoding.UTF8);
            }

            // Anhänge exportieren
            if (message.Attachments != null && message.Attachments.Any())
            {
                string attachFolder = Path.Combine(mailFolder, "Anhänge");
                Directory.CreateDirectory(attachFolder);

                foreach (var attachment in message.Attachments)
                {
                    if (attachment.IsFile)
                    {
                        string safeAttachName = MakeSafeFilename(attachment.FileName);
                        string attachPath = Path.Combine(attachFolder, safeAttachName);
                        
                        // Doppelte Dateinamen in den Anhängen verhindern
                        int aCounter = 1;
                        string attachNameOnly = Path.GetFileNameWithoutExtension(safeAttachName);
                        string attachExt = Path.GetExtension(safeAttachName);
                        while (File.Exists(attachPath))
                        {
                            attachPath = Path.Combine(attachFolder, $"{attachNameOnly}_{aCounter++}{attachExt}");
                        }

                        attachment.SaveToFile(attachPath, attachment.LastModificationTime);
                    }
                }
            }
        }
    }

    private static string MakeSafeFilename(string filename)
    {
        if (string.IsNullOrEmpty(filename))
        {
            return "unbenannt";
        }
        
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            filename = filename.Replace(c, '_');
        }
        
        filename = filename.Replace(" ", "_");
        
        if (filename.Length > 100)
        {
            filename = filename.Substring(0, 100);
        }
            
        return filename;
    }
}
