using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using ExcelDiffTracker.App.Services;
using ExcelDiffTracker.App.ViewModels;

namespace ExcelDiffTracker.App;

public partial class MainWindow : Window
{
    private readonly WindowThemeService _windowTheme;

    public MainWindow(AppServices services)
    {
        InitializeComponent();
        ViewModel = new MainViewModel(services);
        DataContext = ViewModel;
        _windowTheme = WindowThemeService.Attach(this, services.Themes);
    }

    public MainViewModel ViewModel { get; }

    private void OnClosing(object? sender, CancelEventArgs args)
    {
        if (Application.Current is App { IsExiting: false })
        {
            args.Cancel = true;
            Hide();
            ViewModel.ShowToast("Excel Scenario Analysis Tool is still running in the notification area.");
        }
        else
        {
            _windowTheme.Dispose();
        }
    }

    private void OnThemeChanged(object sender, SelectionChangedEventArgs args)
    {
        if (IsLoaded && ViewModel.ApplyThemeCommand.CanExecute(null))
            ViewModel.ApplyThemeCommand.Execute(null);
    }

    private void OnMoreMenuClick(object sender, RoutedEventArgs args)
    {
        if (sender is Button { ContextMenu: { } menu } button)
        {
            menu.PlacementTarget = button;
            menu.Placement = PlacementMode.Left;
            menu.IsOpen = true;
        }
    }
}
