using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using ExcelDiffTracker.App.Services;
using Xunit;

namespace ExcelDiffTracker.App.Tests;

public sealed class OnboardingWindowTests
{
    [Fact]
    public void FirstRunWindowUsesAOneWayProgressBinding()
    {
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                var application = new App();
                application.InitializeComponent();
                using var services = new TestServices();
                var window = new OnboardingWindow(services.Value);
                var progress = Assert.IsType<ProgressBar>(window.FindName("StepProgress"));
                var binding = BindingOperations.GetBinding(progress, ProgressBar.ValueProperty);

                Assert.NotNull(binding);
                Assert.Equal(BindingMode.OneWay, binding.Mode);
            }
            catch (Exception exception)
            {
                failure = exception;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();
    }

    private sealed class TestServices : IDisposable
    {
        public TestServices() => Value = new AppServices();

        public AppServices Value { get; }

        public void Dispose() => Value.DisposeAsync().AsTask().GetAwaiter().GetResult();
    }
}
