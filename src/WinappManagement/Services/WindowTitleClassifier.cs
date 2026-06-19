using System.IO;
using System.Text.RegularExpressions;
using WinappManagement.Models;

namespace WinappManagement.Services;

public static partial class WindowTitleClassifier
{
    private static readonly string[] FileExtensions =
    [
        ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".pdf", ".txt", ".md",
        ".csv", ".json", ".xml", ".html", ".htm", ".cs", ".js", ".ts", ".tsx",
        ".py", ".java", ".cpp", ".c", ".h", ".log", ".zip", ".rar", ".7z"
    ];

    public static (ActivityKind Kind, string DisplayName, string Path, string DirectoryPath, string OpenPath) Classify(
        string title,
        string processName,
        string executablePath)
    {
        if (ShouldDisplayAsPlainApplication(processName))
        {
            return (ActivityKind.Application, FriendlyApplicationName(processName), executablePath, string.Empty, executablePath);
        }

        var fileCandidate = ExtractFileCandidate(title);
        if (!string.IsNullOrWhiteSpace(fileCandidate))
        {
            var fileName = Path.GetFileName(fileCandidate);
            var isOffice = IsOfficeFile(fileName) || IsOfficeProcess(processName);
            var isFullPath = Path.IsPathRooted(fileCandidate);
            var displayPath = isFullPath ? fileCandidate : title;
            var directoryPath = isFullPath ? Path.GetDirectoryName(fileCandidate) ?? string.Empty : string.Empty;
            var openPath = isFullPath ? fileCandidate : string.Empty;
            return (isOffice ? ActivityKind.OfficeFile : ActivityKind.File, fileName, displayPath, directoryPath, openPath);
        }

        if (IsOfficeProcess(processName))
        {
            var officeDisplayName = ExtractOfficeTitle(title, processName);
            return (ActivityKind.OfficeFile, officeDisplayName, title, string.Empty, string.Empty);
        }

        var displayName = string.IsNullOrWhiteSpace(processName)
            ? title
            : processName;

        return (ActivityKind.Application, displayName, executablePath, string.Empty, executablePath);
    }

    private static bool ShouldDisplayAsPlainApplication(string processName)
    {
        return processName.Equals("Code", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("msedge", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("chrome", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("firefox", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("brave", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("brave-browser", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("Obsidian", StringComparison.OrdinalIgnoreCase);
    }

    private static string FriendlyApplicationName(string processName)
    {
        if (processName.Equals("Code", StringComparison.OrdinalIgnoreCase))
        {
            return "Visual Studio Code";
        }

        if (processName.Equals("msedge", StringComparison.OrdinalIgnoreCase))
        {
            return "Microsoft Edge";
        }

        if (processName.Equals("chrome", StringComparison.OrdinalIgnoreCase))
        {
            return "Google Chrome";
        }

        if (processName.Equals("firefox", StringComparison.OrdinalIgnoreCase))
        {
            return "Mozilla Firefox";
        }

        if (processName.Equals("brave", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("brave-browser", StringComparison.OrdinalIgnoreCase))
        {
            return "Brave";
        }

        if (processName.Equals("Obsidian", StringComparison.OrdinalIgnoreCase))
        {
            return "Obsidian";
        }

        return processName;
    }

    private static string? ExtractFileCandidate(string title)
    {
        if (string.IsNullOrWhiteSpace(title))
        {
            return null;
        }

        foreach (var candidate in CandidateParts(title))
        {
            var trimmed = candidate.Trim();
            if (FileExtensions.Any(ext => trimmed.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
            {
                return trimmed;
            }
        }

        var match = FileNamePattern().Match(title);
        return match.Success ? match.Value.Trim() : null;
    }

    private static IEnumerable<string> CandidateParts(string title)
    {
        yield return title;

        foreach (var part in title.Split([" - ", " — ", " – ", " | "], StringSplitOptions.RemoveEmptyEntries))
        {
            yield return part;
        }
    }

    [GeneratedRegex(@"[^\s\\/:*?""<>|]+(?:\.docx?|\.xlsx?|\.pptx?|\.pdf|\.txt|\.md|\.csv|\.json|\.xml|\.html?|\.cs|\.jsx?|\.tsx?|\.py|\.java|\.cpp|\.c|\.h|\.log|\.zip|\.rar|\.7z)", RegexOptions.IgnoreCase)]
    private static partial Regex FileNamePattern();

    private static bool IsOfficeFile(string fileName)
    {
        var extension = Path.GetExtension(fileName);
        return extension.Equals(".doc", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".docx", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".xls", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".ppt", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsOfficeProcess(string processName)
    {
        return processName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase);
    }

    private static string ExtractOfficeTitle(string title, string processName)
    {
        var displayName = title
            .Split([" - ", " — ", " – ", " | "], StringSplitOptions.RemoveEmptyEntries)
            .FirstOrDefault();
        if (string.IsNullOrWhiteSpace(displayName))
        {
            return processName;
        }

        return displayName.Trim();
    }
}
