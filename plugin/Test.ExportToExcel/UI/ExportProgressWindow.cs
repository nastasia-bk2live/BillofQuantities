using System;
using System.Windows;
using System.Windows.Controls;

namespace Test.ExportToExcel.UI
{
    /// <summary>
    /// Простое WPF-окно прогресса с возможностью отмены.
    /// </summary>
    public class ExportProgressWindow : Window
    {
        private readonly ProgressBar _progressBar;
        private readonly TextBlock _statusText;

        public ExportProgressWindow()
        {
            Title = "Экспорт";
            Width = 420;
            Height = 170;
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            var root = new Grid { Margin = new Thickness(12) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            _statusText = new TextBlock
            {
                Text = "Подготовка...",
                Margin = new Thickness(0, 0, 0, 8)
            };

            _progressBar = new ProgressBar
            {
                Minimum = 0,
                Maximum = 100,
                Height = 22,
                Margin = new Thickness(0, 0, 0, 12)
            };

            var cancelButton = new Button
            {
                Content = "Cancel",
                Width = 100,
                Height = 30,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            cancelButton.Click += (sender, args) => { IsCancellationRequested = true; };

            Grid.SetRow(_statusText, 0);
            Grid.SetRow(_progressBar, 1);
            Grid.SetRow(cancelButton, 2);

            root.Children.Add(_statusText);
            root.Children.Add(_progressBar);
            root.Children.Add(cancelButton);

            Content = root;
        }

        public bool IsCancellationRequested { get; private set; }

        public void ReportProgress(int current, int total, string stage)
        {
            total = Math.Max(total, 1);
            var percent = (current * 100.0) / total;

            _statusText.Text = stage + ": " + current + " / " + total;
            _progressBar.Value = percent;
            DoEvents();
        }

        private static void DoEvents()
        {
            var frame = new System.Windows.Threading.DispatcherFrame();
            System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                System.Windows.Threading.DispatcherPriority.Background,
                new Action(() => frame.Continue = false));
            System.Windows.Threading.Dispatcher.PushFrame(frame);
        }
    }
}
