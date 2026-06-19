using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Reflection;
using System.Threading;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using Microsoft.Win32;
using WinappManagement.Models;
using WinappManagement.Services;

namespace WinappManagement;

public sealed class MainViewModel : INotifyPropertyChanged, IDisposable
{
    private readonly WindowInventoryService _inventoryService;
    private readonly FavoritesService _favoritesService;
    private readonly IconService _iconService = new();
    private readonly DispatcherTimer _timer;
    private CancellationTokenSource? _delayedRefreshCancellation;
    private bool _isRefreshing;
    private bool _isDisposed;
    private bool _isAutoRefreshActive;
    private string _searchText = string.Empty;
    private ActivityItem? _selectedItem;
    private string _summaryText = "正在读取窗口...";
    private readonly string _appVersion = GetAppVersion();

    public MainViewModel(WindowInventoryService inventoryService)
    {
        _inventoryService = inventoryService;
        _favoritesService = new FavoritesService();
        _timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(4)
        };
        _timer.Tick += (_, _) => _ = RefreshAsync();

        RefreshCommand = new RelayCommand(_ => _ = RefreshAsync(), _ => !_isRefreshing);
        CloseSelectedCommand = new RelayCommand(_ => CloseSelected(), _ => Items.Any(item => item.IsSelected));
        CloseItemCommand = new RelayCommand(parameter => CloseItem(parameter as ActivityItem), parameter => parameter is ActivityItem);
        ActivateItemCommand = new RelayCommand(parameter => ActivateItem(parameter as ActivityItem), parameter => parameter is ActivityItem);
        ToggleSelectionCommand = new RelayCommand(parameter => ToggleSelection(parameter as ActivityItem), parameter => parameter is ActivityItem);
        ToggleFavoriteCommand = new RelayCommand(parameter => ToggleFavorite(parameter as ActivityItem), parameter => parameter is ActivityItem);
        ResolveOfficeDirectoryCommand = new RelayCommand(parameter => ResolveOfficeDirectory(parameter as ActivityItem), parameter => parameter is ActivityItem item && item.IsOfficeFile);
        OpenFavoriteCommand = new RelayCommand(parameter => OpenFavorite(parameter as FavoriteItem), parameter => parameter is FavoriteItem);
        RemoveFavoriteCommand = new RelayCommand(parameter => RemoveFavorite(parameter as FavoriteItem), parameter => parameter is FavoriteItem);

        foreach (var item in _favoritesService.Load())
        {
            item.Icon = _iconService.GetIcon(item.Kind, item.Path, item.DisplayName);
            Favorites.Add(item);
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event EventHandler? FilterChanged;

    public ObservableCollection<ActivityItem> Items { get; } = [];

    public ObservableCollection<FavoriteItem> Favorites { get; } = [];

    public IEnumerable<FavoriteItem> FolderFavorites => Favorites
        .Where(item => item.Kind == ActivityKind.Folder && MatchesFavoriteSearch(item))
        .OrderBy(item => item.DisplayName);

    public IEnumerable<FavoriteItem> OfficeFavorites => Favorites
        .Where(item => item.IsOfficeFile && MatchesFavoriteSearch(item))
        .OrderBy(item => item.DisplayName);

    public IEnumerable<FavoriteItem> ApplicationFavorites => Favorites
        .Where(item => item.Kind == ActivityKind.Application && !item.IsOfficeFile && MatchesFavoriteSearch(item))
        .OrderBy(item => item.DisplayName);

    public int ApplicationCount => Items.Count(item => item.IsApplicationTabItem && MatchesSearch(item));

    public int OfficeCount => Items.Count(item => item.IsOfficeFile && MatchesSearch(item));

    public int FolderFavoriteCount => FolderFavorites.Count();

    public int OfficeFavoriteCount => OfficeFavorites.Count();

    public int ApplicationFavoriteCount => ApplicationFavorites.Count();

    public Visibility FolderFavoritesEmptyVisibility => FolderFavoriteCount == 0 ? Visibility.Visible : Visibility.Collapsed;

    public Visibility OfficeFavoritesEmptyVisibility => OfficeFavoriteCount == 0 ? Visibility.Visible : Visibility.Collapsed;

    public Visibility ApplicationFavoritesEmptyVisibility => ApplicationFavoriteCount == 0 ? Visibility.Visible : Visibility.Collapsed;

    public string AppVersion => $"v{_appVersion}";

    public RelayCommand RefreshCommand { get; }

    public RelayCommand CloseSelectedCommand { get; }

    public RelayCommand CloseItemCommand { get; }

    public RelayCommand ActivateItemCommand { get; }

    public RelayCommand ToggleSelectionCommand { get; }

    public RelayCommand ToggleFavoriteCommand { get; }

    public RelayCommand ResolveOfficeDirectoryCommand { get; }

    public RelayCommand OpenFavoriteCommand { get; }

    public RelayCommand RemoveFavoriteCommand { get; }

    public string SearchText
    {
        get => _searchText;
        set
        {
            if (_searchText == value)
            {
                return;
            }

            _searchText = value;
            OnPropertyChanged(nameof(SearchText));
            FilterChanged?.Invoke(this, EventArgs.Empty);
            UpdateSummary();
            RefreshFavoriteViews();
        }
    }

    public ActivityItem? SelectedItem
    {
        get => _selectedItem;
        set
        {
            if (_selectedItem == value)
            {
                return;
            }

            _selectedItem = value;
            OnPropertyChanged(nameof(SelectedItem));
            CloseSelectedCommand.RaiseCanExecuteChanged();
        }
    }

    public string SummaryText
    {
        get => _summaryText;
        private set
        {
            if (_summaryText == value)
            {
                return;
            }

            _summaryText = value;
            OnPropertyChanged(nameof(SummaryText));
        }
    }

    public void Start()
    {
        ResumeAutoRefresh();
    }

    public void ResumeAutoRefresh()
    {
        if (_isDisposed || _isAutoRefreshActive)
        {
            return;
        }

        _isAutoRefreshActive = true;
        _timer.Start();
        _ = RefreshAsync();
    }

    public void PauseAutoRefresh()
    {
        _isAutoRefreshActive = false;
        _timer.Stop();
    }

    public bool MatchesSearch(ActivityItem item)
    {
        if (string.IsNullOrWhiteSpace(SearchText))
        {
            return true;
        }

        var query = SearchText.Trim();
        return Contains(item.DisplayName, query)
            || Contains(item.WindowTitle, query)
            || Contains(item.Path, query)
            || Contains(item.ProcessName, query)
            || item.ProcessId.ToString().Contains(query, StringComparison.OrdinalIgnoreCase);
    }

    public bool MatchesApplicationTab(ActivityItem item)
    {
        return item.IsApplicationTabItem && MatchesSearch(item);
    }

    public bool MatchesOfficeTab(ActivityItem item)
    {
        return item.IsOfficeFile && MatchesSearch(item);
    }

    public void Dispose()
    {
        _isDisposed = true;
        _timer.Stop();
        _delayedRefreshCancellation?.Cancel();
        _delayedRefreshCancellation?.Dispose();
    }

    private async Task RefreshAsync()
    {
        if (_isRefreshing || _isDisposed)
        {
            return;
        }

        _isRefreshing = true;
        RefreshCommand.RaiseCanExecuteChanged();
        SummaryText = "正在刷新...";

        IReadOnlyList<ActivityItem> snapshot;
        try
        {
            snapshot = await SnapshotOnStaThreadAsync(_inventoryService);
        }
        catch
        {
            if (!_isDisposed)
            {
                SummaryText = "读取窗口失败，请稍后刷新";
            }

            return;
        }
        finally
        {
            _isRefreshing = false;
            if (!_isDisposed)
            {
                RefreshCommand.RaiseCanExecuteChanged();
            }
        }

        if (_isDisposed)
        {
            return;
        }

        try
        {
            ApplySnapshot(snapshot);
        }
        catch
        {
            SummaryText = "更新列表失败，请稍后刷新";
        }
    }

    private void ApplySnapshot(IReadOnlyList<ActivityItem> snapshot)
    {
        var selectedIds = Items.Where(item => item.IsSelected).Select(item => item.Id).ToHashSet(StringComparer.Ordinal);
        var selectedId = SelectedItem?.Id;
        var favoriteIds = Favorites.Select(item => item.Id).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var syncedItems = SyncItems(snapshot, selectedIds, favoriteIds);
        ReplaceItemsInPlace(syncedItems);
        SelectedItem = Items.FirstOrDefault(item => item.Id == selectedId);
        UpdateSummary();
        FilterChanged?.Invoke(this, EventArgs.Empty);
    }

    private static Task<IReadOnlyList<ActivityItem>> SnapshotOnStaThreadAsync(WindowInventoryService inventoryService)
    {
        var completion = new TaskCompletionSource<IReadOnlyList<ActivityItem>>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try
            {
                completion.SetResult(inventoryService.Snapshot());
            }
            catch (Exception exception)
            {
                completion.SetException(exception);
            }
        })
        {
            IsBackground = true,
            Name = "Winapp Management Snapshot"
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }

    private List<ActivityItem> SyncItems(
        IReadOnlyList<ActivityItem> snapshot,
        HashSet<string> selectedIds,
        HashSet<string> favoriteIds)
    {
        var existingItems = Items.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var syncedItems = new List<ActivityItem>(snapshot.Count);

        foreach (var snapshotItem in snapshot)
        {
            if (existingItems.TryGetValue(snapshotItem.Id, out var previousItem))
            {
                PreserveManualOfficePath(previousItem, snapshotItem);
            }

            if (existingItems.TryGetValue(snapshotItem.Id, out var existingItem) && CanReuseItem(existingItem, snapshotItem))
            {
                existingItem.IsFavorite = existingItem.CanFavorite && favoriteIds.Contains(existingItem.FavoriteKey);
                existingItem.IsSelected = selectedIds.Contains(existingItem.Id);
                syncedItems.Add(existingItem);
                continue;
            }

            if (existingItems.TryGetValue(snapshotItem.Id, out existingItem))
            {
                snapshotItem.Status = existingItem.Status;
            }

            snapshotItem.IsFavorite = snapshotItem.CanFavorite && favoriteIds.Contains(snapshotItem.FavoriteKey);
            snapshotItem.IsSelected = selectedIds.Contains(snapshotItem.Id);
            snapshotItem.PropertyChanged += ActivityItemPropertyChanged;
            syncedItems.Add(snapshotItem);
        }

        return syncedItems;
    }

    private static void PreserveManualOfficePath(ActivityItem existingItem, ActivityItem snapshotItem)
    {
        if (!existingItem.IsOfficeFile
            || !snapshotItem.IsOfficeFile
            || !snapshotItem.IsUnknownOfficeDirectory
            || existingItem.IsUnknownOfficeDirectory
            || string.IsNullOrWhiteSpace(existingItem.OpenPath))
        {
            return;
        }

        snapshotItem.ApplyManualOfficePath(existingItem.OpenPath);
    }

    private static bool CanReuseItem(ActivityItem existingItem, ActivityItem snapshotItem)
    {
        return existingItem.Kind == snapshotItem.Kind
            && existingItem.WindowHandle == snapshotItem.WindowHandle
            && existingItem.ProcessId == snapshotItem.ProcessId
            && string.Equals(existingItem.ProcessName, snapshotItem.ProcessName, StringComparison.Ordinal)
            && string.Equals(existingItem.DisplayName, snapshotItem.DisplayName, StringComparison.Ordinal)
            && string.Equals(existingItem.WindowTitle, snapshotItem.WindowTitle, StringComparison.Ordinal)
            && string.Equals(existingItem.Path, snapshotItem.Path, StringComparison.Ordinal)
            && string.Equals(existingItem.DirectoryPath, snapshotItem.DirectoryPath, StringComparison.Ordinal)
            && string.Equals(existingItem.OpenPath, snapshotItem.OpenPath, StringComparison.Ordinal)
            && existingItem.CanFavorite == snapshotItem.CanFavorite;
    }

    private void ReplaceItemsInPlace(IReadOnlyList<ActivityItem> syncedItems)
    {
        for (var targetIndex = 0; targetIndex < syncedItems.Count; targetIndex++)
        {
            var item = syncedItems[targetIndex];
            var currentIndex = IndexOfItem(item.Id);
            if (currentIndex == targetIndex)
            {
                if (!ReferenceEquals(Items[targetIndex], item))
                {
                    Items[targetIndex].PropertyChanged -= ActivityItemPropertyChanged;
                    Items[targetIndex] = item;
                }

                continue;
            }

            if (currentIndex >= 0)
            {
                Items.Move(currentIndex, targetIndex);
                if (!ReferenceEquals(Items[targetIndex], item))
                {
                    Items[targetIndex].PropertyChanged -= ActivityItemPropertyChanged;
                    Items[targetIndex] = item;
                }
            }
            else
            {
                Items.Insert(targetIndex, item);
            }
        }

        for (var index = Items.Count - 1; index >= syncedItems.Count; index--)
        {
            Items[index].PropertyChanged -= ActivityItemPropertyChanged;
            Items.RemoveAt(index);
        }
    }

    private int IndexOfItem(string id)
    {
        for (var index = 0; index < Items.Count; index++)
        {
            if (string.Equals(Items[index].Id, id, StringComparison.Ordinal))
            {
                return index;
            }
        }

        return -1;
    }

    private void CloseSelected()
    {
        var selectedItems = Items.Where(item => item.IsSelected).ToList();
        foreach (var item in selectedItems)
        {
            var closeRequest = _inventoryService.RequestClose(item);
            item.Status = closeRequest.WasSent
                ? "已发送关闭"
                : "关闭失败";
            _ = ActivateCloseDialogIfPresentAsync(closeRequest);
        }

        ScheduleDelayedRefresh();
    }

    private void CloseItem(ActivityItem? item)
    {
        if (item is null)
        {
            return;
        }

        var closeRequest = _inventoryService.RequestClose(item);
        item.Status = closeRequest.WasSent
            ? "已发送关闭"
            : "关闭失败";
        _ = ActivateCloseDialogIfPresentAsync(closeRequest);
        ScheduleDelayedRefresh();
    }

    private async Task ActivateCloseDialogIfPresentAsync(CloseRequestResult closeRequest)
    {
        try
        {
            await _inventoryService.ActivateCloseDialogIfPresentAsync(closeRequest);
        }
        catch
        {
            // A close confirmation prompt is best-effort only; failure should not block the app.
        }
    }

    private void ScheduleDelayedRefresh()
    {
        if (!_isAutoRefreshActive)
        {
            return;
        }

        _delayedRefreshCancellation?.Cancel();
        _delayedRefreshCancellation?.Dispose();
        _delayedRefreshCancellation = new CancellationTokenSource();
        _ = DelayedRefreshAsync(_delayedRefreshCancellation.Token);
    }

    private async Task DelayedRefreshAsync(CancellationToken cancellationToken)
    {
        try
        {
            await Task.Delay(TimeSpan.FromMilliseconds(800), cancellationToken);
        }
        catch (OperationCanceledException)
        {
            return;
        }

        if (!cancellationToken.IsCancellationRequested)
        {
            await RefreshAsync();
        }
    }

    private void ActivateItem(ActivityItem? item)
    {
        if (item is null)
        {
            return;
        }

        item.Status = _inventoryService.Activate(item)
            ? "已切换到窗口"
            : "切换失败";
    }

    private void ToggleSelection(ActivityItem? item)
    {
        if (item is null)
        {
            return;
        }

        item.IsSelected = !item.IsSelected;
        SelectedItem = item;
    }

    private void ToggleFavorite(ActivityItem? item)
    {
        if (item is null)
        {
            return;
        }

        if (!item.CanFavorite || string.IsNullOrWhiteSpace(item.OpenPath))
        {
            MessageBox.Show(
                "当前只能收藏能获取真实路径的应用、文件夹或文件。这个项目暂时无法获取真实路径，所以不能收藏为可打开项。",
                "无法收藏",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var existing = Favorites.FirstOrDefault(favorite => string.Equals(favorite.Id, item.FavoriteKey, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            Favorites.Remove(existing);
            item.IsFavorite = false;
        }
        else
        {
            var favorite = FavoriteItem.FromActivity(item);
            favorite.Icon = _iconService.GetIcon(favorite.Kind, favorite.Path, favorite.DisplayName);
            Favorites.Add(favorite);
            item.IsFavorite = true;
        }

        SaveFavorites();
        UpdateSummary();
        RefreshFavoriteViews();
    }

    private void ResolveOfficeDirectory(ActivityItem? item)
    {
        if (item is null || !item.IsOfficeFile)
        {
            return;
        }

        var dialog = new OpenFileDialog
        {
            Title = "选择这个 Office 窗口对应的文件",
            CheckFileExists = true,
            Multiselect = false,
            Filter = OfficeFileDialogFilter(item),
            InitialDirectory = InitialOfficeDirectory()
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var fullPath = dialog.FileName;
        if (string.IsNullOrWhiteSpace(fullPath) || !File.Exists(fullPath))
        {
            return;
        }

        item.ApplyManualOfficePath(fullPath);
        item.Icon = _iconService.GetIcon(item.Kind, item.OpenPath, item.DisplayName);
        item.IsFavorite = Favorites.Any(favorite => string.Equals(favorite.Id, item.FavoriteKey, StringComparison.OrdinalIgnoreCase));
        item.Status = "已手工补齐目录";

        UpdateSummary();
        RefreshFavoriteViews();
        FilterChanged?.Invoke(this, EventArgs.Empty);
    }

    private static string OfficeFileDialogFilter(ActivityItem item)
    {
        if (item.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase))
        {
            return "Excel 文件 (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|所有文件 (*.*)|*.*";
        }

        if (item.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase))
        {
            return "Word 文件 (*.docx;*.doc)|*.docx;*.doc|所有文件 (*.*)|*.*";
        }

        if (item.ProcessName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase))
        {
            return "PowerPoint 文件 (*.pptx;*.ppt)|*.pptx;*.ppt|所有文件 (*.*)|*.*";
        }

        return "Office 文件 (*.xlsx;*.xls;*.xlsm;*.docx;*.doc;*.pptx;*.ppt)|*.xlsx;*.xls;*.xlsm;*.docx;*.doc;*.pptx;*.ppt|所有文件 (*.*)|*.*";
    }

    private static string InitialOfficeDirectory()
    {
        var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        return Directory.Exists(documents)
            ? documents
            : Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    }

    private void OpenFavorite(FavoriteItem? item)
    {
        if (item is null)
        {
            return;
        }

        if (!_favoritesService.Open(item))
        {
            MessageBox.Show("打开失败。这个收藏的路径可能已经不存在。", "打开收藏", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void RemoveFavorite(FavoriteItem? item)
    {
        if (item is null)
        {
            return;
        }

        Favorites.Remove(item);
        SaveFavorites();

        foreach (var activity in Items.Where(activity => string.Equals(activity.FavoriteKey, item.Id, StringComparison.OrdinalIgnoreCase)))
        {
            activity.IsFavorite = false;
        }

        UpdateSummary();
        RefreshFavoriteViews();
    }

    private void UpdateSummary()
    {
        var visibleCount = Items.Count(MatchesSearch);
        var totalCount = Items.Count;
        var folderCount = Items.Count(item => item.Kind == ActivityKind.Folder);
        var officeCount = Items.Count(item => item.Kind == ActivityKind.OfficeFile);
        var selectedCount = Items.Count(item => item.IsSelected);
        SummaryText = $"当前显示 {visibleCount}/{totalCount} 项，已选中 {selectedCount} 项，文件夹 {folderCount} 个，Office 文件 {officeCount} 个，收藏 {Favorites.Count} 个";
        OnPropertyChanged(nameof(ApplicationCount));
        OnPropertyChanged(nameof(OfficeCount));
        OnPropertyChanged(nameof(FolderFavoriteCount));
        OnPropertyChanged(nameof(OfficeFavoriteCount));
        OnPropertyChanged(nameof(ApplicationFavoriteCount));
        OnPropertyChanged(nameof(FolderFavoritesEmptyVisibility));
        OnPropertyChanged(nameof(OfficeFavoritesEmptyVisibility));
        OnPropertyChanged(nameof(ApplicationFavoritesEmptyVisibility));
        CloseSelectedCommand.RaiseCanExecuteChanged();
    }

    private void SaveFavorites()
    {
        try
        {
            _favoritesService.Save(Favorites);
        }
        catch
        {
            MessageBox.Show("收藏保存失败，请检查当前用户目录是否可写。", "收藏", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        OnPropertyChanged(nameof(Favorites));
    }

    private bool MatchesFavoriteSearch(FavoriteItem item)
    {
        if (string.IsNullOrWhiteSpace(SearchText))
        {
            return true;
        }

        var query = SearchText.Trim();
        return Contains(item.DisplayName, query)
            || Contains(item.Path, query)
            || Contains(item.ProcessName, query)
            || Contains(item.KindLabel, query);
    }

    private void ActivityItemPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ActivityItem.IsSelected))
        {
            UpdateSummary();
            CloseSelectedCommand.RaiseCanExecuteChanged();
        }
    }

    private void RefreshFavoriteViews()
    {
        OnPropertyChanged(nameof(FolderFavorites));
        OnPropertyChanged(nameof(OfficeFavorites));
        OnPropertyChanged(nameof(ApplicationFavorites));
        OnPropertyChanged(nameof(FolderFavoriteCount));
        OnPropertyChanged(nameof(OfficeFavoriteCount));
        OnPropertyChanged(nameof(ApplicationFavoriteCount));
        OnPropertyChanged(nameof(FolderFavoritesEmptyVisibility));
        OnPropertyChanged(nameof(OfficeFavoritesEmptyVisibility));
        OnPropertyChanged(nameof(ApplicationFavoritesEmptyVisibility));
    }

    private static bool Contains(string value, string query)
    {
        return !string.IsNullOrWhiteSpace(value)
            && value.Contains(query, StringComparison.OrdinalIgnoreCase);
    }

    private void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private static string GetAppVersion()
    {
        var version = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;

        return string.IsNullOrWhiteSpace(version) ? "1.0.2" : version;
    }
}
