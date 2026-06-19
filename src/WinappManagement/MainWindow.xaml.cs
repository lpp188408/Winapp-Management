using System.ComponentModel;
using System.Windows;
using System.Windows.Data;
using System.Windows.Controls;
using WinappManagement.Services;

namespace WinappManagement;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();
        _viewModel = new MainViewModel(new WindowInventoryService());
        DataContext = _viewModel;

        Loaded += (_, _) =>
        {
            ApplyFilter();
            UpdateApplicationPathColumnWidth();
            _viewModel.Start();
        };
        Activated += (_, _) => _viewModel.ResumeAutoRefresh();
        Deactivated += (_, _) => _viewModel.PauseAutoRefresh();
        StateChanged += (_, _) =>
        {
            if (WindowState == WindowState.Minimized)
            {
                _viewModel.PauseAutoRefresh();
            }
            else if (IsActive)
            {
                _viewModel.ResumeAutoRefresh();
            }
        };
        Closed += (_, _) => _viewModel.Dispose();
        _viewModel.FilterChanged += (_, _) => ApplyFilter();
    }

    private void ApplyFilter()
    {
        ApplyFilter("ApplicationItemsViewSource", FilterApplicationItem);
        ApplyFilter("OfficeItemsViewSource", FilterOfficeItem);
    }

    private void ApplyFilter(string resourceKey, FilterEventHandler filter)
    {
        if (Resources[resourceKey] is not CollectionViewSource source)
        {
            return;
        }

        source.Filter -= filter;
        source.Filter += filter;
        if (resourceKey == "OfficeItemsViewSource")
        {
            source.SortDescriptions.Clear();
            source.SortDescriptions.Add(new SortDescription("OfficeTypeOrder", ListSortDirection.Ascending));
            source.SortDescriptions.Add(new SortDescription("DisplayName", ListSortDirection.Ascending));
        }

        source.View?.Refresh();
    }

    private void FilterApplicationItem(object sender, FilterEventArgs e)
    {
        e.Accepted = e.Item is Models.ActivityItem item && _viewModel.MatchesApplicationTab(item);
    }

    private void FilterOfficeItem(object sender, FilterEventArgs e)
    {
        e.Accepted = e.Item is Models.ActivityItem item && _viewModel.MatchesOfficeTab(item);
    }

    private void ApplicationDataGridSizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdateApplicationPathColumnWidth();
    }

    private void UpdateApplicationPathColumnWidth()
    {
        if (ApplicationDataGrid.ActualWidth <= 0)
        {
            return;
        }

        const double gridBorderAndScrollbarReserve = 26;
        var fixedColumnsWidth = ApplicationDataGrid.Columns
            .Where(column => column != ApplicationPathColumn)
            .Sum(column => column.ActualWidth > 0 ? column.ActualWidth : column.Width.DisplayValue);
        var availableWidth = ApplicationDataGrid.ActualWidth - fixedColumnsWidth - gridBorderAndScrollbarReserve;
        ApplicationPathColumn.Width = new DataGridLength(Math.Max(300, availableWidth), DataGridLengthUnitType.Pixel);
    }

}
