using System.Collections.ObjectModel;
using System.Linq;
using XstReader;

namespace Outlook_PST_auslesen.ViewModels;

public class FolderNodeViewModel : ViewModelBase
{
    public XstFolder Folder { get; }
    public string Name => Folder.DisplayName;
    
    private int? _messagesCount;
    public int MessagesCount
    {
        get
        {
            if (_messagesCount == null)
            {
                try
                {
                    _messagesCount = Folder.Messages?.Count() ?? 0;
                }
                catch
                {
                    _messagesCount = 0;
                }
            }
            return _messagesCount.Value;
        }
    }

    public string DisplayName => MessagesCount > 0 ? $"{Name} ({MessagesCount})" : Name;

    public ObservableCollection<FolderNodeViewModel> Subfolders { get; }

    public FolderNodeViewModel(XstFolder folder)
    {
        Folder = folder;
        
        // Rekursives Laden von Unterordnern
        var subFolders = folder.Folders
            .Select(f => new FolderNodeViewModel(f))
            .ToList();

        Subfolders = new ObservableCollection<FolderNodeViewModel>(subFolders);
    }
}
