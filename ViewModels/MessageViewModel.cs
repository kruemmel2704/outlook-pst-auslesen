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

    public string Subject => Message.Subject ?? "<Kein Name/Betreff>";
    
    public string From
    {
        get
        {
            if (IsContact)
            {
                if (!string.IsNullOrEmpty(Email1)) return Email1;
                if (!string.IsNullOrEmpty(CompanyName)) return CompanyName;
                return "[Kontakt]";
            }
            return Message.From ?? "<Unbekannter Absender>";
        }
    }

    public string To
    {
        get
        {
            if (IsContact)
            {
                return CompanyName;
            }
            return Message.To ?? "";
        }
    }

    public string Cc => IsContact ? "" : (Message.Cc ?? "");
    
    public DateTime? Date => IsContact ? null : Message.Date;

    public string DisplayDate
    {
        get
        {
            if (IsContact)
            {
                return "[Kontakt]";
            }
            return Date?.ToLocalTime().ToString("g") ?? "<Unbekannt>";
        }
    }

    public bool HasAttachments => Message.HasAttachments;

    public List<XstAttachment> Attachments => Message.Attachments?.ToList() ?? new List<XstAttachment>();

    // MAPI-Properties für Kontakte
    public string MessageClass => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidTagMessageClass]?.DisplayValue ?? "";
    public bool IsContact => MessageClass.StartsWith("IPM.Contact", StringComparison.OrdinalIgnoreCase);

    public string GivenName => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidTagGivenName]?.DisplayValue ?? "";
    public string Surname => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidTagSurname]?.DisplayValue ?? "";
    public string MiddleName => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidTagMiddleName]?.DisplayValue ?? "";
    public string DisplayNamePrefix => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidTagDisplayNamePrefix]?.DisplayValue ?? "";
    
    public string CompanyName => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidTagCompanyName]?.DisplayValue ?? "";
    public string Title => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidTagTitle]?.DisplayValue ?? "";
    public string DepartmentName => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidTagDepartmentName]?.DisplayValue ?? "";
    
    public string BusinessPhone => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidTagBusinessTelephoneNumber]?.DisplayValue ?? "";
    public string HomePhone => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidTagHomeTelephoneNumber]?.DisplayValue ?? "";
    public string MobilePhone => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidTagMobileTelephoneNumber]?.DisplayValue ?? "";
    
    public string Email1 => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidLidEmail1EmailAddress]?.DisplayValue ?? "";
    public string Email2 => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidLidEmail2EmailAddress]?.DisplayValue ?? "";
    public string Email3 => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidLidEmail3EmailAddress]?.DisplayValue ?? "";
    
    public string WorkStreet => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidLidWorkAddressStreet]?.DisplayValue ?? "";
    public string WorkCity => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidLidWorkAddressCity]?.DisplayValue ?? "";
    public string WorkState => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidLidWorkAddressState]?.DisplayValue ?? "";
    public string WorkPostalCode => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidLidWorkAddressPostalCode]?.DisplayValue ?? "";
    public string WorkCountry => Message.Properties[XstReader.ElementProperties.PropertyCanonicalName.PidLidWorkAddressCountry]?.DisplayValue ?? "";

    public string BodyText
    {
        get
        {
            try
            {
                if (IsContact)
                {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine("=== KONTAKT-DETAILS ===");
                    sb.AppendLine($"Name: {DisplayNamePrefix} {GivenName} {MiddleName} {Surname}".Replace("  ", " ").Trim());
                    if (!string.IsNullOrEmpty(Title)) sb.AppendLine($"Position: {Title}");
                    if (!string.IsNullOrEmpty(CompanyName)) sb.AppendLine($"Firma: {CompanyName}");
                    if (!string.IsNullOrEmpty(DepartmentName)) sb.AppendLine($"Abteilung: {DepartmentName}");
                    sb.AppendLine();
                    sb.AppendLine("--- Kontaktinformationen ---");
                    if (!string.IsNullOrEmpty(Email1)) sb.AppendLine($"E-Mail 1: {Email1}");
                    if (!string.IsNullOrEmpty(Email2)) sb.AppendLine($"E-Mail 2: {Email2}");
                    if (!string.IsNullOrEmpty(Email3)) sb.AppendLine($"E-Mail 3: {Email3}");
                    if (!string.IsNullOrEmpty(MobilePhone)) sb.AppendLine($"Mobil: {MobilePhone}");
                    if (!string.IsNullOrEmpty(BusinessPhone)) sb.AppendLine($"Telefon (Geschäftlich): {BusinessPhone}");
                    if (!string.IsNullOrEmpty(HomePhone)) sb.AppendLine($"Telefon (Privat): {HomePhone}");
                    
                    var addr = new List<string> { WorkStreet, WorkPostalCode, WorkCity, WorkState, WorkCountry }
                        .Where(s => !string.IsNullOrEmpty(s)).ToList();
                    if (addr.Any())
                    {
                        sb.AppendLine();
                        sb.AppendLine("--- Adresse (Geschäftlich) ---");
                        if (!string.IsNullOrEmpty(WorkStreet)) sb.AppendLine(WorkStreet);
                        if (!string.IsNullOrEmpty(WorkPostalCode) || !string.IsNullOrEmpty(WorkCity))
                        {
                            sb.AppendLine($"{WorkPostalCode} {WorkCity}".Trim());
                        }
                        if (!string.IsNullOrEmpty(WorkState)) sb.AppendLine(WorkState);
                        if (!string.IsNullOrEmpty(WorkCountry)) sb.AppendLine(WorkCountry);
                    }

                    var notes = Message.Body?.Text;
                    if (!string.IsNullOrEmpty(notes))
                    {
                        sb.AppendLine();
                        sb.AppendLine("--- Notizen ---");
                        sb.AppendLine(notes);
                    }

                    return sb.ToString();
                }

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
                if (IsContact)
                {
                    return "";
                }

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
            string text = html
                .Replace("<br>", "\n")
                .Replace("<br/>", "\n")
                .Replace("<br />", "\n")
                .Replace("<p>", "")
                .Replace("</p>", "\n\n");
                
            text = Regex.Replace(text, "<.*?>", string.Empty);
            return System.Net.WebUtility.HtmlDecode(text).Trim();
        }
        catch
        {
            return html;
        }
    }
}
