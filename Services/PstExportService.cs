using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using MimeKit;
using XstReader;

namespace Outlook_PST_auslesen.Services;

public static class PstExportService
{
    public static async Task ExportMessageAsync(ViewModels.MessageViewModel msgVm, string targetDir, bool exportAsEml, bool keepFolderStructure = false)
    {
        // Falls es ein Kontakt ist, als vCard (.vcf) exportieren
        if (msgVm.IsContact)
        {
            await ExportContactAsVcardAsync(msgVm, targetDir);
            return;
        }

        // Zielpfad anpassen, falls Ordnerstruktur beibehalten werden soll
        if (keepFolderStructure)
        {
            var folderPathList = GetFolderPath(msgVm.Folder);
            if (folderPathList.Any())
            {
                string subPath = string.Join(Path.DirectorySeparatorChar.ToString(), folderPathList);
                targetDir = Path.Combine(targetDir, subPath);
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }
            }
        }

        var message = msgVm.Message;
        
        // Dateiname sicher machen
        string safeSubject = MakeSafeFilename(msgVm.Subject);
        string dateStr = msgVm.Date?.ToLocalTime().ToString("yyyy-MM-dd_HHmm") ?? "unbekannt";
        string baseName = $"{dateStr}_{safeSubject}";
        
        if (exportAsEml)
        {
            var mimeMessage = await CreateMimeMessageAsync(msgVm);
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

    public static async Task<MimeMessage> CreateMimeMessageAsync(ViewModels.MessageViewModel msgVm)
    {
        var mimeMessage = new MimeMessage();
        var message = msgVm.Message;
        
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
        return mimeMessage;
    }

    public static async Task ExportContactAsVcardAsync(ViewModels.MessageViewModel msgVm, string targetDir)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("BEGIN:VCARD");
        sb.AppendLine("VERSION:3.0");
        
        // Voller Name
        string fullName = $"{msgVm.GivenName} {msgVm.Surname}".Trim();
        if (string.IsNullOrEmpty(fullName)) fullName = msgVm.Subject;
        sb.AppendLine($"FN:{fullName}");
        
        // Strukturierter Name: Nachname;Vorname;Zweitname;Präfix;Suffix
        sb.AppendLine($"N:{msgVm.Surname};{msgVm.GivenName};{msgVm.MiddleName};{msgVm.DisplayNamePrefix};");
        
        // Firma & Abteilung
        if (!string.IsNullOrEmpty(msgVm.CompanyName))
        {
            if (!string.IsNullOrEmpty(msgVm.DepartmentName))
                sb.AppendLine($"ORG:{msgVm.CompanyName};{msgVm.DepartmentName}");
            else
                sb.AppendLine($"ORG:{msgVm.CompanyName}");
        }
        
        // Job-Titel
        if (!string.IsNullOrEmpty(msgVm.Title))
        {
            sb.AppendLine($"TITLE:{msgVm.Title}");
        }
        
        // E-Mails
        if (!string.IsNullOrEmpty(msgVm.Email1)) sb.AppendLine($"EMAIL;TYPE=PREF,INTERNET:{msgVm.Email1}");
        if (!string.IsNullOrEmpty(msgVm.Email2)) sb.AppendLine($"EMAIL;TYPE=INTERNET:{msgVm.Email2}");
        if (!string.IsNullOrEmpty(msgVm.Email3)) sb.AppendLine($"EMAIL;TYPE=INTERNET:{msgVm.Email3}");
        
        // Telefonnummern
        if (!string.IsNullOrEmpty(msgVm.MobilePhone)) sb.AppendLine($"TEL;TYPE=CELL:{msgVm.MobilePhone}");
        if (!string.IsNullOrEmpty(msgVm.BusinessPhone)) sb.AppendLine($"TEL;TYPE=WORK,VOICE:{msgVm.BusinessPhone}");
        if (!string.IsNullOrEmpty(msgVm.HomePhone)) sb.AppendLine($"TEL;TYPE=HOME,VOICE:{msgVm.HomePhone}");
        
        // Adresse (Geschäftlich): Postfach;Zusatz;Straße;Ort;Region;PLZ;Land
        if (!string.IsNullOrEmpty(msgVm.WorkStreet) || !string.IsNullOrEmpty(msgVm.WorkCity) || !string.IsNullOrEmpty(msgVm.WorkPostalCode))
        {
            sb.AppendLine($"ADR;TYPE=WORK:;;{msgVm.WorkStreet};{msgVm.WorkCity};{msgVm.WorkState};{msgVm.WorkPostalCode};{msgVm.WorkCountry}");
        }
        
        // Notizen (Body Text)
        var notes = msgVm.Message.Body?.Text;
        if (!string.IsNullOrEmpty(notes))
        {
            string escapedNotes = notes.Replace("\r\n", "\\n").Replace("\n", "\\n").Replace("\r", "\\n");
            sb.AppendLine($"NOTE:{escapedNotes}");
        }
        
        sb.AppendLine("END:VCARD");
        
        string safeName = MakeSafeFilename(fullName);
        string vcfPath = Path.Combine(targetDir, $"{safeName}.vcf");
        
        int counter = 1;
        while (File.Exists(vcfPath))
        {
            vcfPath = Path.Combine(targetDir, $"{safeName}_{counter++}.vcf");
        }
        
        await File.WriteAllTextAsync(vcfPath, sb.ToString(), System.Text.Encoding.UTF8);
    }

    public static async Task ExportContactsToCsvAsync(IEnumerable<ViewModels.MessageViewModel> contacts, string targetDir)
    {
        var sb = new System.Text.StringBuilder();
        
        // Header-Zeile schreiben (wir benutzen Semikolon als Standardtrenner für deutsche Excel-Versionen)
        var headers = new[]
        {
            "Vorname", "Nachname", "Zweiter Vorname", "Anrede", "Firma", "Position", "Abteilung",
            "E-Mail 1", "E-Mail 2", "E-Mail 3", "Mobiltelefon", "Telefon geschäftlich", "Telefon privat",
            "Straße (geschäftlich)", "Ort (geschäftlich)", "Bundesland (geschäftlich)", "PLZ (geschäftlich)", "Land (geschäftlich)",
            "Notizen"
        };
        
        sb.AppendLine(string.Join(";", headers.Select(EscapeCsvField)));
        
        foreach (var contact in contacts)
        {
            var notes = contact.Message.Body?.Text ?? "";
            var row = new[]
            {
                contact.GivenName,
                contact.Surname,
                contact.MiddleName,
                contact.DisplayNamePrefix,
                contact.CompanyName,
                contact.Title,
                contact.DepartmentName,
                contact.Email1,
                contact.Email2,
                contact.Email3,
                contact.MobilePhone,
                contact.BusinessPhone,
                contact.HomePhone,
                contact.WorkStreet,
                contact.WorkCity,
                contact.WorkState,
                contact.WorkPostalCode,
                contact.WorkCountry,
                notes
            };
            
            sb.AppendLine(string.Join(";", row.Select(EscapeCsvField)));
        }
        
        string csvPath = Path.Combine(targetDir, "kontakte.csv");
        int counter = 1;
        while (File.Exists(csvPath))
        {
            csvPath = Path.Combine(targetDir, $"kontakte_{counter++}.csv");
        }
        
        // UTF-8 mit BOM schreiben, damit Excel die Umlaute korrekt erkennt
        await File.WriteAllTextAsync(csvPath, sb.ToString(), new System.Text.UTF8Encoding(true));
    }
    
    private static string EscapeCsvField(string? field)
    {
        if (string.IsNullOrEmpty(field))
        {
            return "";
        }
        
        // Wenn das Feld Semikolons, Anführungszeichen oder Zeilenumbrüche enthält, muss es maskiert werden
        if (field.Contains(";") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
        {
            return $"\"{field.Replace("\"", "\"\"")}\"";
        }
        
        return field;
    }

    public static List<List<string>> ParseCsv(string csvContent, char separator = ';')
    {
        var records = new List<List<string>>();
        var currentRecord = new List<string>();
        var currentField = new System.Text.StringBuilder();
        bool inQuotes = false;
        
        for (int i = 0; i < csvContent.Length; i++)
        {
            char c = csvContent[i];
            
            if (inQuotes)
            {
                if (c == '"')
                {
                    // Wenn es ein maskiertes Anführungszeichen ist ("")
                    if (i + 1 < csvContent.Length && csvContent[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i++; // Das nächste Anführungszeichen überspringen
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    currentField.Append(c);
                }
            }
            else
            {
                if (c == '"')
                {
                    inQuotes = true;
                }
                else if (c == separator)
                {
                    currentRecord.Add(currentField.ToString());
                    currentField.Clear();
                }
                else if (c == '\r')
                {
                    // '\r' ignorieren, wenn darauf '\n' folgt (wird beim nächsten Schritt verarbeitet)
                    if (i + 1 < csvContent.Length && csvContent[i + 1] == '\n')
                    {
                        i++;
                    }
                    currentRecord.Add(currentField.ToString());
                    records.Add(currentRecord);
                    currentField.Clear();
                    currentRecord = new List<string>();
                }
                else if (c == '\n')
                {
                    currentRecord.Add(currentField.ToString());
                    records.Add(currentRecord);
                    currentField.Clear();
                    currentRecord = new List<string>();
                }
                else
                {
                    currentField.Append(c);
                }
            }
        }
        
        if (currentField.Length > 0 || currentRecord.Count > 0)
        {
            currentRecord.Add(currentField.ToString());
            records.Add(currentRecord);
        }
        
        return records;
    }

    public static async Task ConvertCsvToVcfAsync(string csvPath, string targetDir)
    {
        if (!File.Exists(csvPath))
        {
            throw new FileNotFoundException("CSV-Datei wurde nicht gefunden.", csvPath);
        }
        
        if (!Directory.Exists(targetDir))
        {
            Directory.CreateDirectory(targetDir);
        }

        string csvContent = await File.ReadAllTextAsync(csvPath, System.Text.Encoding.UTF8);
        var records = ParseCsv(csvContent);
        
        if (records.Count < 2)
        {
            throw new InvalidOperationException("Die CSV-Datei enthält keine ausreichenden Daten (mindestens eine Header-Zeile und eine Daten-Zeile erforderlich).");
        }

        var headers = records[0].Select(h => h.Trim().ToLower()).ToList();
        
        // Spalten-Indizes ermitteln
        int idxVorname = headers.IndexOf("vorname");
        int idxNachname = headers.IndexOf("nachname");
        int idxZweiterVorname = headers.IndexOf("zweiter vorname");
        int idxAnrede = headers.IndexOf("anrede");
        int idxFirma = headers.IndexOf("firma");
        int idxPosition = headers.IndexOf("position");
        int idxAbteilung = headers.IndexOf("abteilung");
        int idxEmail1 = headers.IndexOf("e-mail 1");
        int idxEmail2 = headers.IndexOf("e-mail 2");
        int idxEmail3 = headers.IndexOf("e-mail 3");
        int idxMobil = headers.IndexOf("mobiltelefon");
        int idxTelGeschäft = headers.IndexOf("telefon geschäftlich");
        int idxTelPrivat = headers.IndexOf("telefon privat");
        int idxStrasse = headers.IndexOf("straße (geschäftlich)");
        int idxOrt = headers.IndexOf("ort (geschäftlich)");
        int idxBundesland = headers.IndexOf("bundesland (geschäftlich)");
        int idxPlz = headers.IndexOf("plz (geschäftlich)");
        int idxLand = headers.IndexOf("land (geschäftlich)");
        int idxNotizen = headers.IndexOf("notizen");

        if (idxVorname == -1 && idxNachname == -1 && idxEmail1 == -1)
        {
            throw new InvalidOperationException("Das CSV-Format wird nicht erkannt. Es müssen mindestens die Spalten 'Vorname', 'Nachname' oder 'E-Mail 1' vorhanden sein.");
        }

        for (int i = 1; i < records.Count; i++)
        {
            var row = records[i];
            if (row.Count == 0 || (row.Count == 1 && string.IsNullOrEmpty(row[0])))
            {
                continue; // Leere Zeile überspringen
            }

            string GetValue(int index) => (index >= 0 && index < row.Count) ? row[index] : "";

            string vorname = GetValue(idxVorname);
            string nachname = GetValue(idxNachname);
            string zweiterVorname = GetValue(idxZweiterVorname);
            string anrede = GetValue(idxAnrede);
            string firma = GetValue(idxFirma);
            string position = GetValue(idxPosition);
            string abteilung = GetValue(idxAbteilung);
            string email1 = GetValue(idxEmail1);
            string email2 = GetValue(idxEmail2);
            string email3 = GetValue(idxEmail3);
            string mobil = GetValue(idxMobil);
            string telGeschäft = GetValue(idxTelGeschäft);
            string telPrivat = GetValue(idxTelPrivat);
            string strasse = GetValue(idxStrasse);
            string ort = GetValue(idxOrt);
            string bundesland = GetValue(idxBundesland);
            string plz = GetValue(idxPlz);
            string land = GetValue(idxLand);
            string notizen = GetValue(idxNotizen);

            if (string.IsNullOrEmpty(vorname) && string.IsNullOrEmpty(nachname) && string.IsNullOrEmpty(email1) && string.IsNullOrEmpty(firma))
            {
                continue; // Leeren Kontakt überspringen
            }

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("BEGIN:VCARD");
            sb.AppendLine("VERSION:3.0");

            string fullName = $"{vorname} {nachname}".Trim();
            if (string.IsNullOrEmpty(fullName))
            {
                fullName = !string.IsNullOrEmpty(email1) ? email1 : (!string.IsNullOrEmpty(firma) ? firma : "unbenannt");
            }
            sb.AppendLine($"FN:{fullName}");

            sb.AppendLine($"N:{nachname};{vorname};{zweiterVorname};{anrede};");

            if (!string.IsNullOrEmpty(firma))
            {
                if (!string.IsNullOrEmpty(abteilung))
                    sb.AppendLine($"ORG:{firma};{abteilung}");
                else
                    sb.AppendLine($"ORG:{firma}");
            }

            if (!string.IsNullOrEmpty(position))
            {
                sb.AppendLine($"TITLE:{position}");
            }

            if (!string.IsNullOrEmpty(email1)) sb.AppendLine($"EMAIL;TYPE=PREF,INTERNET:{email1}");
            if (!string.IsNullOrEmpty(email2)) sb.AppendLine($"EMAIL;TYPE=INTERNET:{email2}");
            if (!string.IsNullOrEmpty(email3)) sb.AppendLine($"EMAIL;TYPE=INTERNET:{email3}");

            if (!string.IsNullOrEmpty(mobil)) sb.AppendLine($"TEL;TYPE=CELL:{mobil}");
            if (!string.IsNullOrEmpty(telGeschäft)) sb.AppendLine($"TEL;TYPE=WORK,VOICE:{telGeschäft}");
            if (!string.IsNullOrEmpty(telPrivat)) sb.AppendLine($"TEL;TYPE=HOME,VOICE:{telPrivat}");

            if (!string.IsNullOrEmpty(strasse) || !string.IsNullOrEmpty(ort) || !string.IsNullOrEmpty(plz))
            {
                sb.AppendLine($"ADR;TYPE=WORK:;;{strasse};{ort};{bundesland};{plz};{land}");
            }

            if (!string.IsNullOrEmpty(notizen))
            {
                string escapedNotes = notizen.Replace("\r\n", "\\n").Replace("\n", "\\n").Replace("\r", "\\n");
                sb.AppendLine($"NOTE:{escapedNotes}");
            }

            sb.AppendLine("END:VCARD");

            string safeName = MakeSafeFilename(fullName);
            string vcfPath = Path.Combine(targetDir, $"{safeName}.vcf");

            int fileCounter = 1;
            while (File.Exists(vcfPath))
            {
                vcfPath = Path.Combine(targetDir, $"{safeName}_{fileCounter++}.vcf");
            }

            await File.WriteAllTextAsync(vcfPath, sb.ToString(), System.Text.Encoding.UTF8);
        }
    }



    private static List<string> GetFolderPath(XstFolder? folder)
    {
        var path = new List<string>();
        var current = folder;
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
