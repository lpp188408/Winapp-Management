using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using WinappManagement.Models;
using AppActivityKind = WinappManagement.Models.ActivityKind;

namespace WinappManagement.Services;

public sealed class WindowInventoryService
{
    private static readonly TimeSpan OfficeDocumentCacheDuration = TimeSpan.FromSeconds(10);
    private static readonly TimeSpan CloseDialogActivationDelay = TimeSpan.FromMilliseconds(450);

    private readonly int _currentProcessId = Environment.ProcessId;
    private readonly string _currentProcessPath = Environment.ProcessPath ?? string.Empty;
    private readonly IconService _iconService = new();
    private readonly OfficeDocumentService _officeDocumentService = new();
    private IReadOnlyList<OfficeDocumentInfo> _officeDocumentsCache = [];
    private DateTime _officeDocumentsCacheTime = DateTime.MinValue;

    public IReadOnlyList<ActivityItem> Snapshot()
    {
        var items = new List<ActivityItem>();
        var folderHandles = new HashSet<nint>();
        var officeDocuments = GetOfficeDocuments();
        var directOfficeFullNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var folder in GetExplorerFolders())
        {
            if (folder.WindowHandle != 0)
            {
                folderHandles.Add(folder.WindowHandle);
            }

            items.Add(folder);
        }

        foreach (var officeItem in CreateDirectOfficeItems(officeDocuments))
        {
            if (!string.IsNullOrWhiteSpace(officeItem.Path))
            {
                directOfficeFullNames.Add(officeItem.Path);
            }

            items.Add(officeItem);
        }

        NativeMethods.EnumWindows((hWnd, _) =>
        {
            if (!ShouldIncludeWindow(hWnd) || folderHandles.Contains(hWnd))
            {
                return true;
            }

            var title = GetWindowTitle(hWnd);
            if (string.IsNullOrWhiteSpace(title))
            {
                return true;
            }

            NativeMethods.GetWindowThreadProcessId(hWnd, out var pidValue);
            var pid = unchecked((int)pidValue);
            if (pid == _currentProcessId)
            {
                return true;
            }

            var (processName, executablePath) = GetProcessInfo(pid);
            if (IsOwnApplication(processName, executablePath))
            {
                return true;
            }

            var (kind, displayName, path, directoryPath, openPath) = WindowTitleClassifier.Classify(title, processName, executablePath);
            if (kind == AppActivityKind.OfficeFile)
            {
                var officeDocument = MatchOfficeDocument(officeDocuments, processName, displayName, title, hWnd);
                if (processName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase)
                    && directOfficeFullNames.Count > 0)
                {
                    return true;
                }

                if (officeDocument is not null)
                {
                    path = officeDocument.FullName;
                    directoryPath = officeDocument.DirectoryPath;
                    openPath = officeDocument.FullName;
                }
            }

            var canFavorite = !string.IsNullOrWhiteSpace(openPath);

            items.Add(new ActivityItem
            {
                Id = $"{kind}:{hWnd}:{pid}:{title}",
                Kind = kind,
                WindowHandle = hWnd,
                ProcessId = pid,
                ProcessName = processName,
                DisplayName = displayName,
                WindowTitle = title,
                Path = path,
                DirectoryPath = directoryPath,
                OpenPath = canFavorite ? openPath : string.Empty,
                CanFavorite = canFavorite,
                Icon = _iconService.GetIcon(kind, canFavorite ? openPath : path, displayName)
            });

            return true;
        }, 0);

        return new ReadOnlyCollection<ActivityItem>(items);
    }

    private IEnumerable<ActivityItem> CreateDirectOfficeItems(IReadOnlyList<OfficeDocumentInfo> officeDocuments)
    {
        foreach (var document in officeDocuments.Where(document =>
            document.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase)
            && !string.IsNullOrWhiteSpace(document.FullName)
            && Path.IsPathRooted(document.FullName)))
        {
            var hWnd = document.WindowHandle;
            var pid = 0;
            if (hWnd != 0)
            {
                NativeMethods.GetWindowThreadProcessId(hWnd, out var pidValue);
                pid = unchecked((int)pidValue);
            }

            var title = string.IsNullOrWhiteSpace(document.WindowCaption)
                ? document.FileName
                : document.WindowCaption;

            yield return new ActivityItem
            {
                Id = $"office-com:{document.ProcessName}:{hWnd}:{document.FullName}",
                Kind = AppActivityKind.OfficeFile,
                WindowHandle = hWnd,
                ProcessId = pid,
                ProcessName = document.ProcessName,
                DisplayName = document.FileName,
                WindowTitle = title,
                Path = document.FullName,
                DirectoryPath = document.DirectoryPath,
                OpenPath = document.FullName,
                CanFavorite = true,
                Icon = _iconService.GetIcon(AppActivityKind.OfficeFile, document.FullName, document.FileName)
            };
        }
    }

    private IReadOnlyList<OfficeDocumentInfo> GetOfficeDocuments()
    {
        if (DateTime.UtcNow - _officeDocumentsCacheTime < OfficeDocumentCacheDuration)
        {
            return _officeDocumentsCache;
        }

        try
        {
            _officeDocumentsCache = _officeDocumentService.Snapshot();
            _officeDocumentsCacheTime = DateTime.UtcNow;
            return _officeDocumentsCache;
        }
        catch
        {
            _officeDocumentsCache = [];
            _officeDocumentsCacheTime = DateTime.UtcNow;
            return _officeDocumentsCache;
        }
    }

    public CloseRequestResult RequestClose(ActivityItem item)
    {
        if (item.WindowHandle == 0)
        {
            return new CloseRequestResult(false, item.ProcessId, item.WindowHandle, new HashSet<nint>());
        }

        var knownWindows = GetVisibleWindowsForProcess(item.ProcessId);
        var wasSent = NativeMethods.PostMessage(item.WindowHandle, NativeMethods.WmClose, 0, 0);
        return new CloseRequestResult(wasSent, item.ProcessId, item.WindowHandle, knownWindows);
    }

    public async Task<bool> ActivateCloseDialogIfPresentAsync(CloseRequestResult closeRequest)
    {
        if (!closeRequest.WasSent)
        {
            return false;
        }

        await Task.Delay(CloseDialogActivationDelay);

        var dialogHandle = FindNewDialogWindow(closeRequest);
        if (dialogHandle == 0)
        {
            return false;
        }

        _ = NativeMethods.ShowWindow(dialogHandle, NativeMethods.SwRestore);
        return NativeMethods.SetForegroundWindow(dialogHandle);
    }

    public bool Activate(ActivityItem item)
    {
        if (item.WindowHandle == 0)
        {
            return false;
        }

        _ = NativeMethods.ShowWindow(item.WindowHandle, NativeMethods.SwRestore);
        return NativeMethods.SetForegroundWindow(item.WindowHandle);
    }

    private static bool ShouldIncludeWindow(nint hWnd)
    {
        if (!NativeMethods.IsWindowVisible(hWnd))
        {
            return false;
        }

        var exStyle = NativeMethods.GetWindowLong(hWnd, NativeMethods.GwlExStyle);
        return (exStyle & NativeMethods.WsExToolWindow) == 0;
    }

    private static string GetWindowTitle(nint hWnd)
    {
        var length = NativeMethods.GetWindowTextLength(hWnd);
        if (length <= 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder(length + 1);
        _ = NativeMethods.GetWindowText(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    private static (string ProcessName, string ExecutablePath) GetProcessInfo(int pid)
    {
        try
        {
            using var process = Process.GetProcessById(pid);
            var path = string.Empty;
            try
            {
                path = process.MainModule?.FileName ?? string.Empty;
            }
            catch
            {
                path = string.Empty;
            }

            return (process.ProcessName, path);
        }
        catch
        {
            return ($"PID {pid}", string.Empty);
        }
    }

    private bool IsOwnApplication(string processName, string executablePath)
    {
        if (processName.Equals("WinappManagement", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        try
        {
            return !string.IsNullOrWhiteSpace(_currentProcessPath)
                && !string.IsNullOrWhiteSpace(executablePath)
                && string.Equals(
                    Path.GetFullPath(executablePath),
                    Path.GetFullPath(_currentProcessPath),
                    StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static OfficeDocumentInfo? MatchOfficeDocument(
        IReadOnlyList<OfficeDocumentInfo> documents,
        string processName,
        string displayName,
        string title,
        nint windowHandle)
    {
        var candidates = documents
            .Where(document => document.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (candidates.Count == 0)
        {
            return null;
        }

        var handleMatch = candidates.FirstOrDefault(document => document.WindowHandle != 0 && document.WindowHandle == windowHandle);
        if (handleMatch is not null)
        {
            return handleMatch;
        }

        if (processName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) && candidates.Count == 1)
        {
            return candidates[0];
        }

        var normalizedDisplayName = NormalizeOfficeName(displayName);
        var normalizedTitle = NormalizeOfficeName(title);
        var normalizedTitleParts = OfficeNameCandidates(title)
            .Select(NormalizeOfficeName)
            .Where(candidate => !string.IsNullOrWhiteSpace(candidate))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var exactMatch = candidates.FirstOrDefault(document =>
            OfficeNameEquals(document.Name, normalizedDisplayName)
            || OfficeNameEquals(document.FileName, normalizedDisplayName)
            || OfficeNameEquals(FileNameWithoutExtension(document.Name), normalizedDisplayName)
            || OfficeNameEquals(FileNameWithoutExtension(document.FileName), normalizedDisplayName)
            || OfficeNameEquals(document.Name, normalizedTitle)
            || OfficeNameEquals(document.FileName, normalizedTitle)
            || OfficeNameEquals(document.WindowCaption, normalizedTitle)
            || normalizedTitleParts.Any(part =>
                OfficeNameEquals(document.Name, part)
                || OfficeNameEquals(document.FileName, part)
                || OfficeNameEquals(FileNameWithoutExtension(document.Name), part)
                || OfficeNameEquals(FileNameWithoutExtension(document.FileName), part)
                || OfficeNameEquals(document.WindowCaption, part)
                || OfficeNameEquals(FileNameWithoutExtension(document.WindowCaption), part)));
        if (exactMatch is not null)
        {
            return exactMatch;
        }

        var titleMatch = candidates.FirstOrDefault(document =>
            ContainsOfficeName(title, document.Name)
            || ContainsOfficeName(title, document.FileName)
            || ContainsOfficeName(title, document.WindowCaption));
        if (titleMatch is not null)
        {
            return titleMatch;
        }

        return candidates.Count == 1 ? candidates[0] : null;
    }

    private static bool OfficeNameEquals(string left, string normalizedRight)
    {
        var normalizedLeft = NormalizeOfficeName(left);
        if (string.IsNullOrWhiteSpace(normalizedLeft) || string.IsNullOrWhiteSpace(normalizedRight))
        {
            return false;
        }

        return normalizedLeft.Equals(normalizedRight, StringComparison.OrdinalIgnoreCase)
            || FileNameWithoutExtension(normalizedLeft).Equals(FileNameWithoutExtension(normalizedRight), StringComparison.OrdinalIgnoreCase);
    }

    private static bool ContainsOfficeName(string title, string fileName)
    {
        var normalizedTitle = NormalizeOfficeName(title);
        var normalizedFileName = NormalizeOfficeName(fileName);
        var normalizedFileNameWithoutExtension = NormalizeOfficeName(FileNameWithoutExtension(fileName));
        if (!string.IsNullOrWhiteSpace(normalizedFileName)
            && normalizedTitle.Contains(normalizedFileName, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return !string.IsNullOrWhiteSpace(normalizedFileNameWithoutExtension)
            && normalizedTitle.Contains(normalizedFileNameWithoutExtension, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeOfficeName(string value)
    {
        var normalized = value
            .Replace("[兼容模式]", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("[Compatibility Mode]", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("[受保护的视图]", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("[Protected View]", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("[只读]", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("(只读)", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("只读", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Read-Only", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(" - Excel", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(" - Microsoft Excel", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(" - Word", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(" - Microsoft Word", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(" - PowerPoint", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(" - Microsoft PowerPoint", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();

        normalized = RegexCache.ExcelWindowInstanceSuffix().Replace(normalized, string.Empty);

        return normalized.Trim();
    }

    private static string FileNameWithoutExtension(string value)
    {
        try
        {
            return Path.GetFileNameWithoutExtension(value);
        }
        catch
        {
            return value;
        }
    }

    private static IEnumerable<string> OfficeNameCandidates(string title)
    {
        if (string.IsNullOrWhiteSpace(title))
        {
            yield break;
        }

        yield return title;

        foreach (var part in title.Split([" - ", " — ", " – ", " | "], StringSplitOptions.RemoveEmptyEntries))
        {
            yield return part;
        }

        var fileMatch = RegexCache.OfficeFileName().Match(title);
        if (fileMatch.Success)
        {
            yield return fileMatch.Value;
        }
    }

    private static HashSet<nint> GetVisibleWindowsForProcess(int processId)
    {
        var handles = new HashSet<nint>();
        NativeMethods.EnumWindows((hWnd, _) =>
        {
            if (IsVisibleTopLevelWindowForProcess(hWnd, processId))
            {
                handles.Add(hWnd);
            }

            return true;
        }, 0);

        return handles;
    }

    private static nint FindNewDialogWindow(CloseRequestResult closeRequest)
    {
        nint candidate = 0;
        NativeMethods.EnumWindows((hWnd, _) =>
        {
            if (hWnd == closeRequest.OriginalWindowHandle
                || closeRequest.KnownWindows.Contains(hWnd)
                || !IsVisibleTopLevelWindowForProcess(hWnd, closeRequest.ProcessId))
            {
                return true;
            }

            var owner = NativeMethods.GetWindow(hWnd, NativeMethods.GwOwner);
            if (owner != 0 && owner != closeRequest.OriginalWindowHandle)
            {
                return true;
            }

            var title = GetWindowTitle(hWnd);
            if (string.IsNullOrWhiteSpace(title))
            {
                return true;
            }

            candidate = hWnd;
            return false;
        }, 0);

        return candidate;
    }

    private static bool IsVisibleTopLevelWindowForProcess(nint hWnd, int processId)
    {
        if (!NativeMethods.IsWindowVisible(hWnd))
        {
            return false;
        }

        NativeMethods.GetWindowThreadProcessId(hWnd, out var pidValue);
        if (unchecked((int)pidValue) != processId)
        {
            return false;
        }

        var exStyle = NativeMethods.GetWindowLong(hWnd, NativeMethods.GwlExStyle);
        return (exStyle & NativeMethods.WsExToolWindow) == 0;
    }

    private static class RegexCache
    {
        private static readonly Regex ExcelWindowInstanceSuffixRegex = new(@":\d+\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex OfficeFileNameRegex = new(@"[^\s\\/:*?""<>|]+\.(?:docx?|xlsx?|pptx?)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static Regex ExcelWindowInstanceSuffix() => ExcelWindowInstanceSuffixRegex;

        public static Regex OfficeFileName() => OfficeFileNameRegex;
    }

    private IEnumerable<ActivityItem> GetExplorerFolders()
    {
        Type? shellType = Type.GetTypeFromProgID("Shell.Application");
        if (shellType is null)
        {
            yield break;
        }

        object? shell = null;
        try
        {
            shell = Activator.CreateInstance(shellType);
            if (shell is null)
            {
                yield break;
            }

            dynamic windows = ((dynamic)shell).Windows();
            foreach (dynamic window in windows)
            {
                ActivityItem? item = TryCreateExplorerFolder(window);
                if (item is not null)
                {
                    yield return item;
                }
            }
        }
        finally
        {
            if (shell is not null && Marshal.IsComObject(shell))
            {
                Marshal.FinalReleaseComObject(shell);
            }
        }
    }

    private ActivityItem? TryCreateExplorerFolder(dynamic window)
    {
        try
        {
            string fullName = window.FullName;
            if (!fullName.EndsWith("explorer.exe", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            string path = window.Document.Folder.Self.Path;
            if (string.IsNullOrWhiteSpace(path))
            {
                return null;
            }

            nint hWnd = (nint)Convert.ToInt64(window.HWND);
            NativeMethods.GetWindowThreadProcessId(hWnd, out var pidValue);
            var pid = unchecked((int)pidValue);
            var folderName = Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            if (string.IsNullOrWhiteSpace(folderName))
            {
                folderName = path;
            }

            return new ActivityItem
            {
                Id = $"folder:{hWnd}:{path}",
                Kind = AppActivityKind.Folder,
                WindowHandle = hWnd,
                ProcessId = pid,
                ProcessName = "explorer",
                DisplayName = folderName,
                WindowTitle = path,
                Path = path,
                OpenPath = path,
                CanFavorite = true,
                Icon = _iconService.GetIcon(AppActivityKind.Folder, path, folderName)
            };
        }
        catch
        {
            return null;
        }
    }
}

public sealed record CloseRequestResult(
    bool WasSent,
    int ProcessId,
    nint OriginalWindowHandle,
    IReadOnlySet<nint> KnownWindows);
