using System;
using System.IO;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Outlook_PST_auslesen.ViewModels;

namespace Outlook_PST_auslesen.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private async void OpenPst_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null)
        {
            return;
        }

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Outlook PST-Datei öffnen",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Outlook Datendatei (*.pst)")
                {
                    Patterns = new[] { "*.pst" }
                }
            }
        });

        if (files.Count > 0)
        {
            string localPath = files[0].Path.LocalPath;
            if (DataContext is MainWindowViewModel vm)
            {
                await vm.LoadPstFileAsync(localPath);
            }
        }
    }

    private async void ConvertCsvToVcf_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null)
        {
            return;
        }

        // 1. CSV-Datei auswählen
        var csvFiles = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "CSV-Kontaktdatei auswählen",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("CSV-Datei (*.csv)")
                {
                    Patterns = new[] { "*.csv" }
                }
            }
        });

        if (csvFiles.Count == 0)
        {
            return;
        }

        string csvPath = csvFiles[0].Path.LocalPath;

        // 2. Zielverzeichnis für VCFs auswählen
        var targetFolders = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "Zielverzeichnis für vCard-Dateien (.vcf) auswählen",
            AllowMultiple = false
        });

        if (targetFolders.Count == 0)
        {
            return;
        }

        string targetDir = targetFolders[0].Path.LocalPath;

        if (DataContext is MainWindowViewModel vm)
        {
            vm.IsLoading = true;
            vm.StatusMessage = "CSV-Datei wird in VCF umgewandelt...";
            vm.Progress = 0;

            try
            {
                await Task.Run(async () =>
                {
                    await Services.PstExportService.ConvertCsvToVcfAsync(csvPath, targetDir);
                });

                vm.StatusMessage = $"Erfolgreich VCF-Dateien im Ordner erstellt: {targetDir}";
            }
            catch (Exception ex)
            {
                vm.StatusMessage = $"Fehler bei der CSV-Konvertierung: {ex.Message}";
            }
            finally
            {
                vm.IsLoading = false;
            }
        }
    }

    private async void SelectExportFolder_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null)
        {
            return;
        }

        var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "Export-Zielverzeichnis auswählen",
            AllowMultiple = false
        });

        if (folders.Count > 0)
        {
            string localPath = folders[0].Path.LocalPath;
            if (DataContext is MainWindowViewModel vm)
            {
                vm.ExportFolderPath = localPath;
            }
        }
    }

    private async void SaveAttachment_Click(object? sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is XstReader.XstAttachment attachment)
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null)
            {
                return;
            }

            var file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Dateianhang speichern",
                SuggestedFileName = attachment.FileName,
                DefaultExtension = Path.GetExtension(attachment.FileName)
            });

            if (file != null)
            {
                string savePath = file.Path.LocalPath;
                try
                {
                    attachment.SaveToFile(savePath);
                    if (DataContext is MainWindowViewModel vm)
                    {
                        vm.StatusMessage = $"Anhang erfolgreich gespeichert: {Path.GetFileName(savePath)}";
                    }
                }
                catch (Exception ex)
                {
                    if (DataContext is MainWindowViewModel vm)
                    {
                        vm.StatusMessage = $"Fehler beim Speichern des Anhangs: {ex.Message}";
                    }
                }
            }
        }
    }
}