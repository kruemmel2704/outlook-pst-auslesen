using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using CommunityToolkit.Mvvm.ComponentModel;
using XstReader;

namespace Outlook_PST_auslesen.ViewModels;

public partial class MessageViewModel : ViewModelBase
{
    [ObservableProperty]
    private bool _isSelected;

    public XstMessage Message { get; }

    public string Subject => Message.Subject ?? "<Kein Betreff>";
    public string From => Message.From ?? "<Unbekannter Absender>";
    public string To => Message.To ?? "";
    public string Cc => Message.Cc ?? "";
    
    // In der Reflektion haben wir gesehen, dass Nullable`1 Date existiert.
    public DateTime? Date => Message.Date;

    public string DisplayDate => Date?.ToLocalTime().ToString("g") ?? "<Unbekannt>";

    public bool HasAttachments => Message.HasAttachments;

    public List<XstAttachment> Attachments => Message.Attachments?.ToList() ?? new List<XstAttachment>();

    public string BodyText
    {
        get
        {
            try
            {
                var body = Message.Body;
                if (body == null)
                {
                    return "";
                }

                if (body.Format == XstMessageBodyFormat.Html)
                {
                    return StripHtml(body.Text);
                }
                
                return body.Text ?? "";
            }
            catch (Exception ex)
            {
                return $"[Fehler beim Lesen des Textkörpers: {ex.Message}]";
            }
        }
    }

    public string BodyHtml
    {
        get
        {
            try
            {
                var body = Message.Body;
                if (body == null)
                {
                    return "";
                }

                if (body.Format == XstMessageBodyFormat.Html)
                {
                    return body.Text ?? "";
                }
                
                return "";
            }
            catch
            {
                return "";
            }
        }
    }

    public MessageViewModel(XstMessage message)
    {
        Message = message;
    }

    private static string StripHtml(string html)
    {
        if (string.IsNullOrEmpty(html))
        {
            return "";
        }
        
        try
        {
            // Einfache HTML-Zeilenumbrüche in Newlines konvertieren
            string text = html
                .Replace("<br>", "\n")
                .Replace("<br/>", "\n")
                .Replace("<br />", "\n")
                .Replace("<p>", "")
                .Replace("</p>", "\n\n");
                
            // Alle anderen HTML-Tags entfernen
            text = Regex.Replace(text, "<.*?>", string.Empty);
            
            // HTML-Entities dekodieren
            return System.Net.WebUtility.HtmlDecode(text).Trim();
        }
        catch
        {
            return html; // Fallback, falls Regex fehlschlägt
        }
    }
}
