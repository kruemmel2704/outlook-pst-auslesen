using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using XstReader;

namespace Outlook_PST_auslesen.ViewModels;

public enum ContactExportFormat
{
    Vcf,
    Csv
}

public partial class MainWindowViewModel : ViewModelBase
{
    private XstFile? _xstFile;

    [ObservableProperty]
    private string _pstFilePath = "";

    [ObservableProperty]
    private string _exportFolderPath = "";

    [ObservableProperty]
    private bool _isLoading;

    [ObservableProperty]
    private string _statusMessage = "Bitte öffnen Sie eine PST-Datei.";

    [ObservableProperty]
    private double _progress;

    [ObservableProperty]
    private bool _exportAsEml = true;

    [ObservableProperty]
    private string _searchText = "";

    [ObservableProperty]
    private bool _showContactExportDialog;

    [ObservableProperty]
    private string _contactExportDialogMessage = "";

    private System.Collections.Generic.List<MessageViewModel> _selectedItemsToExport = new();

    [ObservableProperty]
    private bool _showImapSyncDialog;

    [ObservableProperty]
    private string _imapServer = "";

    [ObservableProperty]
    private int _imapPort = 993;

    [ObservableProperty]
    private string _imapUser = "";

    [ObservableProperty]
    private string _imapPassword = "";

    [ObservableProperty]
    private bool _usePstFolderStructure = true;

    [ObservableProperty]
    private bool _keepFolderStructure = true;

    private FolderNodeViewModel? _selectedFolder;
    public FolderNodeViewModel? SelectedFolder
    {
        get => _selectedFolder;
        set
        {
            if (SetProperty(ref _selectedFolder, value))
            {
                LoadMessagesForSelectedFolder();
            }
        }
    }

    private MessageViewModel? _selectedMessage;
    public MessageViewModel? SelectedMessage
    {
        get => _selectedMessage;
        set
        {
            SetProperty(ref _selectedMessage, value);
        }
    }

    public ObservableCollection<FolderNodeViewModel> Folders { get; } = new();
    
    private readonly ObservableCollection<MessageViewModel> _allMessages = new();
    public ObservableCollection<MessageViewModel> CurrentFolderMessages { get; } = new();

    private bool _isAllSelected;
    public bool IsAllSelected
    {
        get => _isAllSelected;
        set
        {
            if (SetProperty(ref _isAllSelected, value))
            {
                foreach (var msg in CurrentFolderMessages)
                {
                    msg.IsSelected = value;
                }
            }
        }
    }

    partial void OnSearchTextChanged(string value)
    {
        FilterMessages();
    }

    public async Task LoadPstFileAsync(string path)
    {
        if (string.IsNullOrEmpty(path) || !File.Exists(path))
        {
            return;
        }

        IsLoading = true;
        StatusMessage = "PST-Datei wird geladen...";
        Progress = 0;

        try
        {
            // Alte Datei schließen, falls vorhanden
            if (_xstFile != null)
            {
                _xstFile.Dispose();
                _xstFile = null;
            }

            Folders.Clear();
            _allMessages.Clear();
            CurrentFolderMessages.Clear();
            SelectedFolder = null;
            SelectedMessage = null;

            await Task.Run(() =>
            {
                _xstFile = new XstFile(path);
                
                // Root-Ordner laden
                var rootNodes = _xstFile.RootFolder.Folders
                    .Select(f => new FolderNodeViewModel(f))
                    .ToList();

                Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                {
                    PstFilePath = path;
                    foreach (var node in rootNodes)
                    {
                        Folders.Add(node);
                    }
                    StatusMessage = "PST-Datei erfolgreich geladen. Wählen Sie links einen Ordner.";
                });
            });
        }
        catch (Exception ex)
        {
            StatusMessage = $"Fehler beim Laden der PST: {ex.Message}";
        }
        finally
        {
            IsLoading = false;
        }
    }

    private void LoadMessagesForSelectedFolder()
    {
        _allMessages.Clear();
        CurrentFolderMessages.Clear();
        SelectedMessage = null;
        _isAllSelected = false;
        OnPropertyChanged(nameof(IsAllSelected));

        if (SelectedFolder == null)
        {
            return;
        }

        try
        {
            var folder = SelectedFolder.Folder;
            
            // Nachrichten laden
            var messages = folder.Messages
                .Select(m => new MessageViewModel(m, folder))
                .ToList();

            foreach (var msg in messages)
            {
                _allMessages.Add(msg);
            }
            FilterMessages();
            StatusMessage = $"{_allMessages.Count} E-Mails im Ordner '{SelectedFolder.Name}' geladen.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Fehler beim Laden der E-Mails: {ex.Message}";
        }
    }

    private void FilterMessages()
    {
        CurrentFolderMessages.Clear();
        var query = SearchText?.Trim().ToLower() ?? "";
        
        var filtered = _allMessages.AsEnumerable();
        if (!string.IsNullOrEmpty(query))
        {
            filtered = filtered.Where(m => 
                m.Subject.ToLower().Contains(query) || 
                m.From.ToLower().Contains(query) ||
                m.To.ToLower().Contains(query) ||
                m.BodyText.ToLower().Contains(query));
        }

        foreach (var msg in filtered)
        {
            CurrentFolderMessages.Add(msg);
        }
    }

    [RelayCommand]
    public async Task ExportSelectedAsync()
    {
        var selected = CurrentFolderMessages.Where(m => m.IsSelected).ToList();
        
        if (!selected.Any())
        {
            StatusMessage = "Fehler: Keine Einträge für den Export ausgewählt.";
            return;
        }

        if (string.IsNullOrEmpty(ExportFolderPath) || !Directory.Exists(ExportFolderPath))
        {
            StatusMessage = "Fehler: Bitte wählen Sie ein gültiges Export-Zielverzeichnis.";
            return;
        }

        var contacts = selected.Where(m => m.IsContact).ToList();
        if (contacts.Any())
        {
            _selectedItemsToExport = selected;
            ContactExportDialogMessage = $"Es wurden {contacts.Count} Kontakte in Ihrer Auswahl erkannt. Wie möchten Sie diese exportieren?";
            ShowContactExportDialog = true;
        }
        else
        {
            await RunExportAsync(selected, null);
        }
    }

    [RelayCommand]
    public async Task ExportContactsAsVcfAsync()
    {
        ShowContactExportDialog = false;
        var items = _selectedItemsToExport;
        _selectedItemsToExport = new();
        await RunExportAsync(items, ContactExportFormat.Vcf);
    }

    [RelayCommand]
    public async Task ExportContactsAsCsvAsync()
    {
        ShowContactExportDialog = false;
        var items = _selectedItemsToExport;
        _selectedItemsToExport = new();
        await RunExportAsync(items, ContactExportFormat.Csv);
    }

    [RelayCommand]
    public void CancelContactExport()
    {
        ShowContactExportDialog = false;
        _selectedItemsToExport.Clear();
        StatusMessage = "Export abgebrochen.";
    }

    private async Task RunExportAsync(System.Collections.Generic.List<MessageViewModel> items, ContactExportFormat? contactFormat)
    {
        IsLoading = true;
        StatusMessage = "Export wird gestartet...";
        Progress = 0;

        try
        {
            var contacts = items.Where(m => m.IsContact).ToList();
            var emails = items.Where(m => !m.IsContact).ToList();

            int totalSteps = emails.Count;
            if (contacts.Any())
            {
                if (contactFormat == ContactExportFormat.Csv)
                {
                    totalSteps += 1;
                }
                else
                {
                    totalSteps += contacts.Count;
                }
            }

            int currentStep = 0;

            await Task.Run(async () =>
            {
                // 1. Kontakte exportieren
                if (contacts.Any())
                {
                    if (contactFormat == ContactExportFormat.Csv)
                    {
                        try
                        {
                            await Services.PstExportService.ExportContactsToCsvAsync(contacts, ExportFolderPath);
                        }
                        catch (Exception ex)
                        {
                            Avalonia.Threading.Dispatcher.UIThread.Post(() => {
                                StatusMessage = $"Fehler beim CSV-Export: {ex.Message}";
                            });
                        }
                        currentStep++;
                        UpdateProgress(currentStep, totalSteps, "Kontakte in CSV-Datei exportiert...");
                    }
                    else
                    {
                        foreach (var contact in contacts)
                        {
                            try
                            {
                                await Services.PstExportService.ExportContactAsVcardAsync(contact, ExportFolderPath);
                            }
                            catch (Exception ex)
                            {
                                Avalonia.Threading.Dispatcher.UIThread.Post(() => {
                                    StatusMessage = $"Fehler beim VCF-Export von '{contact.Subject}': {ex.Message}";
                                });
                            }
                            currentStep++;
                            UpdateProgress(currentStep, totalSteps, $"Exportiere Kontakt {currentStep} von {contacts.Count}...");
                        }
                    }
                }

                // 2. E-Mails exportieren
                foreach (var email in emails)
                {
                    try
                    {
                        await Services.PstExportService.ExportMessageAsync(email, ExportFolderPath, ExportAsEml, KeepFolderStructure);
                    }
                    catch (Exception ex)
                    {
                        Avalonia.Threading.Dispatcher.UIThread.Post(() => {
                            StatusMessage = $"Fehler beim Exportieren von '{email.Subject}': {ex.Message}";
                        });
                    }

                    currentStep++;
                    int emailIndex = currentStep - (contactFormat == ContactExportFormat.Csv ? 1 : contacts.Count);
                    UpdateProgress(currentStep, totalSteps, $"Exportiere E-Mail {emailIndex} von {emails.Count}...");
                }
            });

            StatusMessage = $"Export erfolgreich abgeschlossen nach: {ExportFolderPath}";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Fehler beim Exportieren: {ex.Message}";
        }
        finally
        {
            IsLoading = false;
        }
    }

    private void UpdateProgress(int current, int total, string message)
    {
        double progressVal = (double)current / total * 100;
        Avalonia.Threading.Dispatcher.UIThread.Post(() => {
            Progress = progressVal;
            StatusMessage = message;
        });
    }

    [RelayCommand]
    public void OpenInBrowser()
    {
        if (SelectedMessage == null)
        {
            return;
        }
        
        string html = SelectedMessage.BodyHtml;
        if (string.IsNullOrEmpty(html))
        {
            html = $"<html><head><meta charset=\"utf-8\"></head><body><pre style=\"white-space: pre-wrap; font-family: sans-serif;\">{SelectedMessage.BodyText}</pre></body></html>";
        }

        try
        {
            string tempPath = Path.Combine(Path.GetTempPath(), $"pst_preview_{Guid.NewGuid()}.html");
            File.WriteAllText(tempPath, html, System.Text.Encoding.UTF8);
            
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = tempPath,
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(psi);
        }
        catch (Exception ex)
        {
            StatusMessage = $"Fehler beim Öffnen im Browser: {ex.Message}";
        }
    }

    [RelayCommand]
    public void OpenImapSyncDialog()
    {
        var selected = CurrentFolderMessages.Where(m => m.IsSelected && !m.IsContact).ToList();
        if (!selected.Any())
        {
            StatusMessage = "Fehler: Keine E-Mails für die Synchronisierung ausgewählt (Kontakte werden ignoriert).";
            return;
        }
        ShowImapSyncDialog = true;
    }

    [RelayCommand]
    public void CancelImapSync()
    {
        ShowImapSyncDialog = false;
        StatusMessage = "Synchronisierung abgebrochen.";
    }

    [RelayCommand]
    public async Task StartImapSyncAsync()
    {
        var selected = CurrentFolderMessages.Where(m => m.IsSelected && !m.IsContact).ToList();
        if (!selected.Any())
        {
            StatusMessage = "Fehler: Keine E-Mails zum Synchronisieren.";
            ShowImapSyncDialog = false;
            return;
        }

        if (string.IsNullOrEmpty(ImapServer) || string.IsNullOrEmpty(ImapUser) || string.IsNullOrEmpty(ImapPassword))
        {
            StatusMessage = "Fehler: Bitte geben Sie Server, E-Mail und Passwort an.";
            return;
        }

        ShowImapSyncDialog = false;
        IsLoading = true;
        StatusMessage = "Synchronisierung wird gestartet...";
        Progress = 0;

        try
        {
            await Services.ImapSyncService.SyncEmailsAsync(
                selected, 
                ImapServer, 
                ImapPort, 
                ImapUser, 
                ImapPassword, 
                UsePstFolderStructure, 
                (current, total, message) =>
                {
                    double progressVal = (double)current / total * 100;
                    Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                    {
                        Progress = progressVal;
                        StatusMessage = message;
                    });
                }
            );

            StatusMessage = $"{selected.Count} E-Mails erfolgreich mit dem Postfach synchronisiert.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Fehler beim Postfach-Sync: {ex.Message}";
        }
        finally
        {
            IsLoading = false;
        }
    }

    public void ClosePst()
    {
        if (_xstFile != null)
        {
            _xstFile.Dispose();
            _xstFile = null;
        }
    }
}
