using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Media;

namespace WinappManagement.Models;

public sealed class ActivityItem : INotifyPropertyChanged
{
    private bool _isFavorite;
    private bool _isSelected;
    private string _status = "可关闭";

    public event PropertyChangedEventHandler? PropertyChanged;

    public required string Id { get; init; }
    public required ActivityKind Kind { get; init; }
    public required nint WindowHandle { get; init; }
    public required int ProcessId { get; init; }
    public required string ProcessName { get; init; }
    public required string DisplayName { get; init; }
    public required string WindowTitle { get; init; }
    public required string Path { get; set; }
    public string DirectoryPath { get; set; } = string.Empty;
    public required string OpenPath { get; set; }
    public required bool CanFavorite { get; set; }
    public ImageSource? Icon { get; set; }

    public string Status
    {
        get => _status;
        set
        {
            if (_status == value)
            {
                return;
            }

            _status = value;
            OnPropertyChanged(nameof(Status));
        }
    }

    public bool IsFavorite
    {
        get => _isFavorite;
        set
        {
            if (_isFavorite == value)
            {
                return;
            }

            _isFavorite = value;
            OnPropertyChanged(nameof(IsFavorite));
            OnPropertyChanged(nameof(FavoriteGlyph));
        }
    }

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value)
            {
                return;
            }

            _isSelected = value;
            OnPropertyChanged(nameof(IsSelected));
        }
    }

    public string FavoriteGlyph => IsFavorite ? "★" : "☆";

    public string FavoriteKey => FavoriteItem.CreateKey(Kind, OpenPath);

    public bool IsOfficeFile => Kind == ActivityKind.OfficeFile;

    public bool IsApplicationTabItem => Kind != ActivityKind.OfficeFile;

    public bool IsFolder => Kind == ActivityKind.Folder;

    public bool IsSystemWindow => Kind == ActivityKind.Application && IsSystemProcess(ProcessName);

    public string ApplicationGroupName => IsSystemWindow ? "系统窗口" : "常用项目";

    public string SecondaryTitle => ShouldShowSecondaryTitle() ? WindowTitle : string.Empty;

    public Visibility SecondaryTitleVisibility => string.IsNullOrWhiteSpace(SecondaryTitle)
        ? Visibility.Collapsed
        : Visibility.Visible;

    public string DisplayPath => string.IsNullOrWhiteSpace(Path) ? WindowTitle : Path;

    public string OfficeDirectory => !string.IsNullOrWhiteSpace(DirectoryPath)
        ? DirectoryPath
        : TryGetDirectory(Path);

    public string OfficeDirectoryDisplay => string.IsNullOrWhiteSpace(OfficeDirectory)
        ? "未识别目录"
        : OfficeDirectory;

    public bool IsUnknownOfficeDirectory => string.IsNullOrWhiteSpace(OfficeDirectory);

    public Visibility UnknownOfficeDirectoryVisibility => IsUnknownOfficeDirectory
        ? Visibility.Visible
        : Visibility.Collapsed;

    public Visibility KnownOfficeDirectoryVisibility => IsUnknownOfficeDirectory
        ? Visibility.Collapsed
        : Visibility.Visible;

    public string KindLabel => Kind switch
    {
        ActivityKind.Folder => "文件夹",
        ActivityKind.File => "文件",
        ActivityKind.OfficeFile => OfficeTypeLabel,
        _ => "应用"
    };

    public string OfficeTypeLabel
    {
        get
        {
            if (Kind != ActivityKind.OfficeFile)
            {
                return string.Empty;
            }

            if (IsExcelFile() || ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase))
            {
                return "Excel";
            }

            if (IsWordFile() || ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase))
            {
                return "Word";
            }

            if (IsPowerPointFile() || ProcessName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase))
            {
                return "PPT";
            }

            return "Office";
        }
    }

    public int OfficeTypeOrder => OfficeTypeLabel switch
    {
        "Excel" => 0,
        "Word" => 1,
        "PPT" => 2,
        _ => 9
    };

    public int KindOrder => Kind switch
    {
        ActivityKind.Application => 0,
        ActivityKind.Folder => 1,
        ActivityKind.File => 2,
        ActivityKind.OfficeFile => 3,
        _ => 99
    };

    private void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public void ApplyManualOfficePath(string fullPath)
    {
        if (string.IsNullOrWhiteSpace(fullPath))
        {
            return;
        }

        Path = fullPath;
        DirectoryPath = TryGetDirectory(fullPath);
        OpenPath = fullPath;
        CanFavorite = true;

        OnPropertyChanged(nameof(Path));
        OnPropertyChanged(nameof(DirectoryPath));
        OnPropertyChanged(nameof(OpenPath));
        OnPropertyChanged(nameof(CanFavorite));
        OnPropertyChanged(nameof(FavoriteKey));
        OnPropertyChanged(nameof(DisplayPath));
        OnPropertyChanged(nameof(OfficeDirectory));
        OnPropertyChanged(nameof(OfficeDirectoryDisplay));
        OnPropertyChanged(nameof(IsUnknownOfficeDirectory));
        OnPropertyChanged(nameof(UnknownOfficeDirectoryVisibility));
        OnPropertyChanged(nameof(KnownOfficeDirectoryVisibility));
        OnPropertyChanged(nameof(KindLabel));
        OnPropertyChanged(nameof(OfficeTypeLabel));
        OnPropertyChanged(nameof(OfficeTypeOrder));
    }

    private static string TryGetDirectory(string path)
    {
        try
        {
            return string.IsNullOrWhiteSpace(path) ? string.Empty : System.IO.Path.GetDirectoryName(path) ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private bool IsExcelFile()
    {
        var extension = OfficeExtension();
        return extension.Equals(".xls", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase);
    }

    private bool IsWordFile()
    {
        var extension = OfficeExtension();
        return extension.Equals(".doc", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".docx", StringComparison.OrdinalIgnoreCase);
    }

    private bool IsPowerPointFile()
    {
        var extension = OfficeExtension();
        return extension.Equals(".ppt", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase);
    }

    private string OfficeExtension()
    {
        return ExtensionFrom(DisplayName)
            ?? ExtensionFrom(OpenPath)
            ?? ExtensionFrom(Path)
            ?? string.Empty;
    }

    private static string? ExtensionFrom(string value)
    {
        try
        {
            var extension = System.IO.Path.GetExtension(value);
            return string.IsNullOrWhiteSpace(extension) ? null : extension;
        }
        catch
        {
            return null;
        }
    }

    private static bool IsSystemProcess(string processName)
    {
        return processName.Equals("TextInputHost", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("ApplicationFrameHost", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("SystemSettings", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("ShellExperienceHost", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("SearchHost", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("StartMenuExperienceHost", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("RuntimeBroker", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("SecurityHealthSystray", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("LockApp", StringComparison.OrdinalIgnoreCase);
    }

    private bool ShouldShowSecondaryTitle()
    {
        if (Kind == ActivityKind.Folder || Kind == ActivityKind.OfficeFile)
        {
            return false;
        }

        if (ShouldHideApplicationTitle(ProcessName))
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(WindowTitle))
        {
            return false;
        }

        return !WindowTitle.Equals(DisplayName, StringComparison.OrdinalIgnoreCase)
            && !WindowTitle.Equals(Path, StringComparison.OrdinalIgnoreCase);
    }

    private static bool ShouldHideApplicationTitle(string processName)
    {
        return processName.Equals("Code", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("msedge", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("chrome", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("firefox", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("brave", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("brave-browser", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("Obsidian", StringComparison.OrdinalIgnoreCase);
    }
}
