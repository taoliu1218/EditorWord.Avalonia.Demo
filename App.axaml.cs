using System;
using System.Reactive;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Threading;
using EditWord.Avalonia.ViewModels;
using EditWord.Avalonia.Views;
using ReactiveUI;

namespace EditWord.Avalonia;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is not IClassicDesktopStyleApplicationLifetime desktop) return;
        
        RxApp.DefaultExceptionHandler = Observer.Create<Exception>(ex =>
        {
            Console.WriteLine(ex.Message);
        });
        Dispatcher.UIThread.UnhandledException += (_, ex) =>
        {
            ex.Handled = true;
            Console.WriteLine(ex.Exception.Message);
        };

        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            if (e.ExceptionObject is Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        };
        
        
        desktop.MainWindow = new MainWindow
        {
            DataContext = new MainWindowViewModel(),
        };

        base.OnFrameworkInitializationCompleted();
    }
}