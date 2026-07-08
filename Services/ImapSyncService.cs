using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using MailKit;
using MailKit.Net.Imap;
using MimeKit;
using XstReader;

namespace Outlook_PST_auslesen.Services;

public static class ImapSyncService
{
    public static async Task SyncEmailsAsync(
        List<ViewModels.MessageViewModel> messages, 
        string host, 
        int port, 
        string username, 
        string password, 
        bool usePstFolderStructure,
        Action<int, int, string> reportProgress)
    {
        using var client = new ImapClient();
        
        // SSL/TLS erzwingen für Port 993, sonst STARTTLS oder kein SSL
        bool useSsl = port == 993;
        
        // Verbindung herstellen
        await client.ConnectAsync(host, port, useSsl);
        
        // Authentifizieren
        await client.AuthenticateAsync(username, password);

        var rootFolder = client.GetFolder(client.PersonalNamespaces[0]);
        var folderCache = new Dictionary<string, IMailFolder>();

        int total = messages.Count;
        int current = 0;

        foreach (var msgVm in messages)
        {
            // Kontakte überspringen (werden nicht in E-Mail-Postfächer gesynct)
            if (msgVm.IsContact)
            {
                current++;
                continue;
            }

            IMailFolder targetFolder;

            if (usePstFolderStructure)
            {
                // Pfad ermitteln
                var path = GetFolderPath(msgVm.Folder);
                string pathKey = string.Join("/", path);

                if (folderCache.TryGetValue(pathKey, out var cachedFolder))
                {
                    targetFolder = cachedFolder;
                }
                else
                {
                    targetFolder = rootFolder;
                    foreach (var folderName in path)
                    {
                        IMailFolder subFolder;
                        try
                        {
                            subFolder = await targetFolder.GetSubfolderAsync(folderName);
                        }
                        catch (FolderNotFoundException)
                        {
                            subFolder = await targetFolder.CreateAsync(folderName, true);
                        }
                        targetFolder = subFolder;
                    }
                    folderCache[pathKey] = targetFolder;
                }
            }
            else
            {
                // Alle in einen "PST-Import" Ordner
                if (folderCache.TryGetValue("PST-Import", out var cachedFolder))
                {
                    targetFolder = cachedFolder;
                }
                else
                {
                    try
                    {
                        targetFolder = await rootFolder.GetSubfolderAsync("PST-Import");
                    }
                    catch (FolderNotFoundException)
                    {
                        targetFolder = await rootFolder.CreateAsync("PST-Import", true);
                    }
                    folderCache["PST-Import"] = targetFolder;
                }
            }

            // MimeMessage erstellen und hochladen
            var mimeMessage = await PstExportService.CreateMimeMessageAsync(msgVm);
            
            // Markieren als gelesen (Seen)
            await targetFolder.AppendAsync(mimeMessage, MessageFlags.Seen);

            current++;
            reportProgress(current, total, $"Synchronisiere E-Mail {current} von {total}: {msgVm.Subject}");
        }

        await client.DisconnectAsync(true);
    }

    private static List<string> GetFolderPath(XstFolder? folder)
    {
        var path = new List<string>();
        var current = folder;
        
        // Den gesamten Pfad nach oben traversieren, aber den Root-Ordner (PST-Dateiname) ignorieren
        while (current != null && current.ParentFolder != null)
        {
            if (!string.IsNullOrEmpty(current.DisplayName))
            {
                path.Insert(0, current.DisplayName);
            }
            current = current.ParentFolder;
        }
        
        return path;
    }
}
