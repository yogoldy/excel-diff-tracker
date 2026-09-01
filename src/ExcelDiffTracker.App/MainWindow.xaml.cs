using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using ExcelDiffTracker.App.Services;
using ExcelDiffTracker.App.ViewModels;

namespace ExcelDiffTracker.App;

public partial class MainWindow : Window
{
    public MainWindow(AppServices services)
    {
        InitializeComponent();
        ViewModel = new MainViewModel(services);
        DataContext = ViewModel;
    }

    public MainViewModel ViewModel { get; }

    private void OnClosing(object? sender, CancelEventArgs args)
    {
        if (Application.Current is App { IsExiting: false })
        {
            args.Cancel = true;
            Hide();
            ViewModel.ShowToast("Excel Diff Tracker is still running in the notification area.");
        }
    }

    private void OnThemeChanged(object sender, SelectionChangedEventArgs args)
    {
        if (IsLoaded && ViewModel.ApplyThemeCommand.CanExecute(null))
            ViewModel.ApplyThemeCommand.Execute(null);
    }
}
