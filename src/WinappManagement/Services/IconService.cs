using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using WinappManagement.Models;

namespace WinappManagement.Services;

public sealed class IconService
{
    private const int MaxCacheSize = 128;
    private readonly object _cacheLock = new();
    private readonly Dictionary<string, LinkedListNode<IconCacheEntry>> _cache = new(StringComparer.OrdinalIgnoreCase);
    private readonly LinkedList<IconCacheEntry> _cacheOrder = new();

    public ImageSource? GetIcon(ActivityKind kind, string path, string displayName)
    {
        var key = CreateKey(kind, path, displayName);
        lock (_cacheLock)
        {
            if (_cache.TryGetValue(key, out var node))
            {
                _cacheOrder.Remove(node);
                _cacheOrder.AddFirst(node);
                return node.Value.Icon;
            }
        }

        var icon = LoadIcon(kind, path, displayName);
        lock (_cacheLock)
        {
            if (_cache.TryGetValue(key, out var existingNode))
            {
                _cacheOrder.Remove(existingNode);
                _cacheOrder.AddFirst(existingNode);
                return existingNode.Value.Icon;
            }

            var node = new LinkedListNode<IconCacheEntry>(new IconCacheEntry(key, icon));
            _cacheOrder.AddFirst(node);
            _cache[key] = node;

            while (_cache.Count > MaxCacheSize && _cacheOrder.Last is not null)
            {
                var lastNode = _cacheOrder.Last;
                _cacheOrder.RemoveLast();
                _cache.Remove(lastNode.Value.Key);
            }
        }

        return icon;
    }

    private static ImageSource? LoadIcon(ActivityKind kind, string path, string displayName)
    {
        var iconPath = ResolveIconPath(kind, path, displayName);
        var attributes = kind == ActivityKind.Folder
            ? NativeMethods.FileAttributeDirectory
            : NativeMethods.FileAttributeNormal;

        var flags = NativeMethods.ShgfiIcon | NativeMethods.ShgfiSmallIcon;
        if (!File.Exists(iconPath) && !Directory.Exists(iconPath))
        {
            flags |= NativeMethods.ShgfiUseFileAttributes;
        }

        var info = new NativeMethods.ShFileInfo();
        var result = NativeMethods.SHGetFileInfo(
            iconPath,
            attributes,
            ref info,
            (uint)System.Runtime.InteropServices.Marshal.SizeOf<NativeMethods.ShFileInfo>(),
            flags);

        if (result == 0 || info.hIcon == 0)
        {
            return null;
        }

        try
        {
            var source = Imaging.CreateBitmapSourceFromHIcon(
                info.hIcon,
                Int32Rect.Empty,
                BitmapSizeOptions.FromWidthAndHeight(20, 20));
            source.Freeze();
            return source;
        }
        finally
        {
            _ = NativeMethods.DestroyIcon(info.hIcon);
        }
    }

    private static string ResolveIconPath(ActivityKind kind, string path, string displayName)
    {
        if (kind == ActivityKind.OfficeFile)
        {
            var existingPath = ExistingPathOrEmpty(path);
            return string.IsNullOrWhiteSpace(existingPath)
                ? OfficeExtension(displayName, path)
                : existingPath;
        }

        var realPath = ExistingPathOrEmpty(path);
        if (!string.IsNullOrWhiteSpace(realPath))
        {
            return realPath;
        }

        return kind switch
        {
            ActivityKind.Folder => Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ActivityKind.File => ExtensionOrDefault(displayName, path),
            ActivityKind.Application => string.IsNullOrWhiteSpace(path) ? Environment.ProcessPath ?? "explorer.exe" : path,
            _ => Environment.ProcessPath ?? "explorer.exe"
        };
    }

    private static string OfficeExtension(string displayName, string path)
    {
        var extension = ExtractOfficeExtension(displayName);
        if (string.IsNullOrWhiteSpace(extension))
        {
            extension = ExtractOfficeExtension(path);
        }

        if (!string.IsNullOrWhiteSpace(extension))
        {
            return $"sample{extension}";
        }

        return "sample.docx";
    }

    private static string ExtensionOrDefault(string displayName, string path)
    {
        var extension = SafeExtension(displayName);
        if (string.IsNullOrWhiteSpace(extension))
        {
            extension = SafeExtension(path);
        }

        return string.IsNullOrWhiteSpace(extension) ? "sample.txt" : $"sample{extension}";
    }

    private static string ExistingPathOrEmpty(string path)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            return File.Exists(path) || Directory.Exists(path) ? path : string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string ExtractOfficeExtension(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var match = Regex.Match(value, @"\.(docx?|xlsx?|pptx?)\b", RegexOptions.IgnoreCase);
        return match.Success ? match.Value.ToLowerInvariant() : string.Empty;
    }

    private static string SafeExtension(string value)
    {
        try
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : Path.GetExtension(value);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string CreateKey(ActivityKind kind, string path, string displayName)
    {
        var iconPath = ResolveIconPath(kind, path, displayName);
        return $"{kind}:{iconPath}";
    }

    private sealed record IconCacheEntry(string Key, ImageSource? Icon);
}
