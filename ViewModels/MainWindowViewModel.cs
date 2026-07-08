using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using XstReader;

namespace Outlook_PST_auslesen.ViewModels;

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
                .Select(m => new MessageViewModel(m))
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
            StatusMessage = "Fehler: Keine E-Mails für den Export ausgewählt.";
            return;
        }

        if (string.IsNullOrEmpty(ExportFolderPath) || !Directory.Exists(ExportFolderPath))
        {
            StatusMessage = "Fehler: Bitte wählen Sie ein gültiges Export-Zielverzeichnis.";
            return;
        }

        IsLoading = true;
        StatusMessage = "Export wird gestartet...";
        Progress = 0;

        try
        {
            int total = selected.Count;
            int current = 0;

            await Task.Run(async () =>
            {
                foreach (var msgVm in selected)
                {
                    try
                    {
                        await Services.PstExportService.ExportMessageAsync(msgVm, ExportFolderPath, ExportAsEml);
                    }
                    catch (Exception ex)
                    {
                        Avalonia.Threading.Dispatcher.UIThread.Post(() => {
                            StatusMessage = $"Fehler beim Exportieren von '{msgVm.Subject}': {ex.Message}";
                        });
                    }

                    current++;
                    double progressVal = (double)current / total * 100;
                    
                    Avalonia.Threading.Dispatcher.UIThread.Post(() => {
                        Progress = progressVal;
                        StatusMessage = $"Exportiere E-Mail {current} von {total}...";
                    });
                }
            });

            StatusMessage = $"{total} E-Mails erfolgreich exportiert nach: {ExportFolderPath}";
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

    public void ClosePst()
    {
        if (_xstFile != null)
        {
            _xstFile.Dispose();
            _xstFile = null;
        }
    }
}
