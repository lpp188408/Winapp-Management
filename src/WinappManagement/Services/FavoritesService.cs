using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using WinappManagement.Models;

namespace WinappManagement.Services;

public sealed class FavoritesService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly string _filePath;

    public FavoritesService()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var folder = Path.Combine(appData, "WinappManagement");
        _filePath = Path.Combine(folder, "favorites.json");
    }

    public ObservableCollection<FavoriteItem> Load()
    {
        try
        {
            if (!File.Exists(_filePath))
            {
                return [];
            }

            var json = File.ReadAllText(_filePath);
            var items = JsonSerializer.Deserialize<List<FavoriteItem>>(json, JsonOptions) ?? [];
            return new ObservableCollection<FavoriteItem>(items.OrderBy(item => item.DisplayName));
        }
        catch
        {
            return [];
        }
    }

    public void Save(IEnumerable<FavoriteItem> items)
    {
        var folder = Path.GetDirectoryName(_filePath);
        if (!string.IsNullOrWhiteSpace(folder))
        {
            Directory.CreateDirectory(folder);
        }

        var json = JsonSerializer.Serialize(items.OrderBy(item => item.DisplayName).ToList(), JsonOptions);
        File.WriteAllText(_filePath, json);
    }

    public bool Open(FavoriteItem item)
    {
        if (string.IsNullOrWhiteSpace(item.Path))
        {
            return false;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = item.Path,
                UseShellExecute = true
            });
            return true;
        }
        catch
        {
            return false;
        }
    }
}
