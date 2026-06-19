using System.Text.Json.Serialization;
using System.Windows.Media;

namespace WinappManagement.Models;

public sealed class FavoriteItem
{
    public required string Id { get; init; }
    public required ActivityKind Kind { get; init; }
    public required string DisplayName { get; init; }
    public required string Path { get; init; }
    public string ProcessName { get; init; } = string.Empty;
    public DateTime CreatedAt { get; init; } = DateTime.UtcNow;
    [JsonIgnore]
    public ImageSource? Icon { get; set; }

    [JsonIgnore]
    public bool IsOfficeFile => Kind == ActivityKind.OfficeFile
        || Path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)
        || Path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
        || Path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)
        || Path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
        || Path.EndsWith(".ppt", StringComparison.OrdinalIgnoreCase)
        || Path.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase);

    [JsonIgnore]
    public string KindLabel => Kind switch
    {
        ActivityKind.Folder => "文件夹",
        ActivityKind.File => "文件",
        ActivityKind.OfficeFile => "Office文件",
        _ => "应用"
    };

    public static FavoriteItem FromActivity(ActivityItem item)
    {
        return new FavoriteItem
        {
            Id = CreateKey(item.Kind, item.OpenPath),
            Kind = item.Kind,
            DisplayName = item.DisplayName,
            Path = item.OpenPath,
            ProcessName = item.ProcessName,
            CreatedAt = DateTime.UtcNow
        };
    }

    public static string CreateKey(ActivityKind kind, string path)
    {
        return $"{kind}:{path.Trim()}".ToUpperInvariant();
    }
}
