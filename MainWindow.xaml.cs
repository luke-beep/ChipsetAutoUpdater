// <copyright file="MainWindow.xaml.cs" company="nocorp">
// Copyright (c) realies. No rights reserved.
// </copyright>

using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;
using Microsoft.Win32;
using Microsoft.Win32.TaskScheduler;
using Application = System.Windows.Application;
using Cursors = System.Windows.Input.Cursors;
using MessageBox = System.Windows.MessageBox;
using Task = System.Threading.Tasks.Task;

namespace ChipsetAutoUpdater
{
    using Application = Application;
    using Cursors = Cursors;
    using MessageBox = MessageBox;
    using SystemTask = Task;

    /// <summary>
    ///     Interaction logic for MainWindow.xaml.
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string TaskName = "ChipsetAutoUpdaterAutoStart";
        private const int UpdateCheckIntervalHours = 6;

        private NotifyIcon _notifyIcon;


        private string _latestVersionReleasePageUrl;
        private string _latestVersionDownloadFileUrl;
        private string _latestVersionString;
        private CancellationTokenSource _cancellationTokenSource;
        private HttpClient _client;
        private bool _isInitializing = true;
        private DispatcherTimer _updateTimer;

        private string InstalledVersionString
        {
            get => InstalledVersionText.Text;
            set => InstalledVersionText.Text = value;
        }

        private string DetectedChipsetString
        {
            get => ChipsetModelText.Text;
            set => ChipsetModelText.Text = value;
        }

        private bool AutoUpdateEnabled
        {
            get => AutoUpdateCheckBox.IsChecked != null && AutoUpdateCheckBox.IsChecked.Value;
            set => AutoUpdateCheckBox.IsChecked = value;
        }

        public bool AutoStartEnabled
        {
            get => AutoStartCheckBox.IsChecked != null && AutoStartCheckBox.IsChecked.Value;
            set => AutoStartCheckBox.IsChecked = value;
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="MainWindow" /> class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            InitializeNotifyIcon();

            StateChanged += MainWindow_StateChanged;
            Loaded += Window_Loaded;

            HandleArguments();

            InitializeTitle();
            InitializeHttpClient();
            InitializeUpdateTimer();
            _ = RenderView();
        }

        /// <summary>
        ///     Performs cleanup operations when the window is closed.
        ///     Disposes of the notify icon and calls the base OnClosed method.
        /// </summary>
        /// <param name="e">A System.EventArgs that contains the event data.</param>
        protected override void OnClosed(EventArgs e)
        {
            _notifyIcon.Dispose();
            base.OnClosed(e);
        }

        private void HandleArguments()
        {
            var args = Environment.GetCommandLineArgs();
            if (args.Contains("/min")) CloseWindow();

            if (args.Contains("/autoupdate")) AutoUpdateEnabled = true;
        }

        private void InitializeTitle()
        {
            Title = "CAU " +
                    Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyFileVersionAttribute), false)
                        .OfType<AssemblyFileVersionAttribute>().FirstOrDefault()?.Version
                        ?.TrimEnd(".0".ToCharArray()) +
                    " ";
        }

        private void InitializeHttpClient()
        {
            _client = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(1)
            };
            _client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (compatible; AcmeInc/1.0)");
            _client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("text/html"));
            _client.DefaultRequestHeaders.Referrer = new Uri("https://www.amd.com/");
        }

        private void InitializeUpdateTimer()
        {
            _updateTimer = new DispatcherTimer();
            _updateTimer.Tick += async (sender, e) => await CheckForUpdates();
            _updateTimer.Interval = TimeSpan.FromHours(UpdateCheckIntervalHours);
            _updateTimer.Start();
        }

        private async SystemTask CheckForUpdates()
        {
            await RenderView();
            if (_latestVersionString != null && _latestVersionString != InstalledVersionString)
            {
                OpenWindow();
                Activate();
                if (AutoUpdateEnabled) InstallDrivers_Click(this, new RoutedEventArgs());
            }
        }

        private void AutoStartCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (_isInitializing) return;

            if (AutoUpdateEnabled)
            {
                CreateStartupTask();
                AutoUpdateEnabled = true;
            }
            else
            {
                RemoveStartupTask();
                AutoUpdateEnabled = false;
            }
        }

        private void AutoUpdateCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (_isInitializing) return;
            CreateStartupTask();
        }

        private void CreateStartupTask()
        {
            using (var ts = new TaskService())
            {
                var td = ts.NewTask();
                td.RegistrationInfo.Description = "Start ChipsetAutoUpdater at system startup";

                td.Triggers.Add(new LogonTrigger());

                var arguments = "/min";
                if (AutoUpdateEnabled) arguments += " /autoupdate";

                td.Actions.Add(new ExecAction(Assembly.GetExecutingAssembly().Location, arguments));

                td.Principal.RunLevel = TaskRunLevel.Highest;

                ts.RootFolder.RegisterTaskDefinition(TaskName, td);
            }
        }

        private static void RemoveStartupTask()
        {
            using (var ts = new TaskService())
            {
                ts.RootFolder.DeleteTask(TaskName, false);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            using (var ts = new TaskService())
            {
                var task = ts.GetTask(TaskName);
                AutoStartEnabled = task != null;
                AutoUpdateEnabled = task != null && task.Definition.Actions.OfType<ExecAction>()
                    .Any(a => a.Arguments.Contains("/autoupdate"));
            }

            _isInitializing = false;
        }

        private void InitializeNotifyIcon()
        {
            _notifyIcon = new NotifyIcon
            {
                Icon = System.Drawing.Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location),
                Visible = true,
                Text = @"Chipset Auto Updater"
            };

            // Create context menu
            var contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("Open", null, OnOpenClick);
            contextMenu.Items.Add("Exit", null, OnExitClick);

            _notifyIcon.ContextMenuStrip = contextMenu;
            _notifyIcon.DoubleClick += NotifyIcon_DoubleClick;
        }

        private async void MainWindow_StateChanged(object sender, EventArgs e)
        {
            switch (WindowState)
            {
                case WindowState.Minimized:
                    Hide();
                    break;
                case WindowState.Normal:
                    await CheckForUpdates();
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private void NotifyIcon_DoubleClick(object sender, EventArgs e)
        {
            OpenWindow();
        }

        private void OnOpenClick(object sender, EventArgs e)
        {
            OpenWindow();
        }

        private void OpenWindow()
        {
            Show();
            WindowState = WindowState.Normal;
        }

        private void CloseWindow()
        {
            Hide();
            WindowState = WindowState.Minimized;
        }

        private static void OnExitClick(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }

        private async SystemTask RenderView()
        {
            DetectedChipsetString = ChipsetModelMatcher() ?? "Not Detected";
            SetInstalledVersion();

            if (DetectedChipsetString != null)
            {
                var versionData = await FetchLatestVersionData(DetectedChipsetString);
                _latestVersionReleasePageUrl = versionData.Item1;
                _latestVersionDownloadFileUrl = versionData.Item2;
                _latestVersionString = versionData.Item3;
            }

            if (_latestVersionString != null)
            {
                LatestVersionText.Text = _latestVersionString;
                LatestVersionText.TextDecorations = TextDecorations.Underline;
                LatestVersionText.Cursor = Cursors.Hand;
                LatestVersionText.MouseLeftButtonDown += (s, e) =>
                {
                    if (_latestVersionReleasePageUrl != null)
                        Process.Start(new ProcessStartInfo(_latestVersionReleasePageUrl) { UseShellExecute = true });
                };
                InstallDriversButton.IsEnabled = true;
            }
            else
            {
                LatestVersionText.Text = "Error fetching";
                LatestVersionText.TextDecorations = null;
                LatestVersionText.Cursor = Cursors.Arrow;
                LatestVersionText.MouseLeftButtonDown -= (s, e) => { };
                InstallDriversButton.IsEnabled = false;
            }
        }

        private static string ChipsetModelMatcher()
        {
            const string registryPath = @"HARDWARE\DESCRIPTION\System\BIOS";
            using (var key = Registry.LocalMachine.OpenSubKey(registryPath))
            {
                if (key == null) return null;
                var product = key.GetValue("BaseBoardProduct")?.ToString().ToUpper() ?? string.Empty;
                const string
                    pattern =
                        @"\b[A-Z]{1,2}\d{2,4}[A-Z]?\b"; // Define the regex pattern to match chipset models like A/B/X followed by three digits and optional E
                var match = Regex.Match(product, pattern);
                if (match.Success) return match.Value;
            }

            return null;
        }

        private static string GetInstalledAmdChipsetVersion()
        {
            const string registryPath = @"Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
            using (var key = Registry.LocalMachine.OpenSubKey(registryPath))
            {
                if (key == null) return null;
                foreach (var subkeyName in key.GetSubKeyNames())
                    using (var subkey = key.OpenSubKey(subkeyName))
                    {
                        if (subkey != null)
                        {
                            var publisher = subkey.GetValue("Publisher") as string;
                            var displayName = subkey.GetValue("DisplayName") as string;
                            if (publisher == "Advanced Micro Devices, Inc." &&
                                displayName == "AMD Chipset Software")
                                return subkey.GetValue("DisplayVersion") as string;
                        }
                    }
            }

            return null;
        }

        private async Task<Tuple<string, string, string>> FetchLatestVersionData(string chipset)
        {
            try
            {
                const string driversPageUrl = "https://www.amd.com/en/support/download/drivers.html";
                var driversPageContent = await _client.GetStringAsync(driversPageUrl);
                var driversPagePattern = $@"https://[^""&]+{Regex.Escape(chipset)}\.html";
                var driversUrlMatch = Regex.Match(driversPageContent, driversPagePattern, RegexOptions.IgnoreCase);
                if (driversUrlMatch.Success)
                {
                    var chipsetPageContent = await _client.GetStringAsync(driversUrlMatch.Value);
                    const string chipsetDownloadFileUrlPattern = @"https://[^""]+\.exe";
                    var chipsetDownloadFileUrlMatch = Regex.Match(chipsetPageContent, chipsetDownloadFileUrlPattern,
                        RegexOptions.IgnoreCase);
                    const string chipsetReleasePageUrlPattern =
                        @"/en/resources/support-articles/release-notes/.*chipset[^""]+";
                    var chipsetReleasePageUrlMatch = Regex.Match(chipsetPageContent, chipsetReleasePageUrlPattern,
                        RegexOptions.IgnoreCase);
                    if (chipsetDownloadFileUrlMatch.Success && chipsetReleasePageUrlMatch.Success)
                        return Tuple.Create($@"https://www.amd.com{chipsetReleasePageUrlMatch.Value}",
                            chipsetDownloadFileUrlMatch.Value,
                            chipsetDownloadFileUrlMatch.Value.Split('_').Last().Replace(".exe", string.Empty));
                }
            }
            catch (Exception)
            {
                // ignored
            }

            return Tuple.Create<string, string, string>(null, null, null);
        }

        private async void InstallDrivers_Click(object sender, RoutedEventArgs e)
        {
            var localFilePath = Path.Combine(Path.GetTempPath(), $"amd_chipset_software_{_latestVersionString}.exe");
            try
            {
                InstallDriversButton.IsEnabled = false;
                InstallDriversButton.Visibility = Visibility.Collapsed;
                DownloadProgressBar.Visibility = Visibility.Visible;
                CancelButton.Visibility = Visibility.Visible;
                _cancellationTokenSource = new CancellationTokenSource();

                await DownloadFileAsync(_latestVersionDownloadFileUrl, localFilePath, _cancellationTokenSource.Token);

                if (_cancellationTokenSource.Token
                    .IsCancellationRequested) return;

                CancelButton.IsEnabled = false;
                CancelButton.Content = "Installing...";
                var startInfo = new ProcessStartInfo
                {
                    FileName = localFilePath,
                    Arguments = "-INSTALL",
                    Verb = "runas",
                    UseShellExecute = true,
                    CreateNoWindow = true
                };
                var process = Process.Start(startInfo);
                await SystemTask.Run(() => MonitorProcess(process));
                await SystemTask.Run(() => { process?.WaitForExit(); });
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"Error downloading or installing drivers: {ex.Message}");
            }
            finally
            {
                InstallDriversButton.IsEnabled = true;
                InstallDriversButton.Visibility = Visibility.Visible;
                DownloadProgressBar.Value = 0;
                CancelButton.Visibility = Visibility.Collapsed;
                CancelButton.Content = "Cancel Download";
                CleanUpFile(localFilePath);
            }
        }

        private async SystemTask DownloadFileAsync(string requestUri, string destinationFilePath,
            CancellationToken cancellationToken)
        {
            try
            {
                using (var response = await _client.GetAsync(requestUri, HttpCompletionOption.ResponseHeadersRead,
                           cancellationToken))
                {
                    _ = response.EnsureSuccessStatusCode();
                    var totalBytes = response.Content.Headers.ContentLength ?? -1L;
                    var canReportProgress = totalBytes != -1;
                    using (var contentStream = await response.Content.ReadAsStreamAsync())
                    using (var fileStream = new FileStream(destinationFilePath, FileMode.Create, FileAccess.Write,
                               FileShare.None, 8192, true))
                    {
                        var totalRead = 0L;
                        var buffer = new byte[8192];
                        var isMoreToRead = true;
                        do
                        {
                            var read = await contentStream.ReadAsync(buffer, 0, buffer.Length, cancellationToken);
                            if (read == 0)
                            {
                                isMoreToRead = false;
                            }
                            else
                            {
                                await fileStream.WriteAsync(buffer, 0, read, cancellationToken);
                                totalRead += read;
                                if (!canReportProgress) continue;

                                var progress = (int)(totalRead * 100 / totalBytes);
                                DownloadProgressBar.Value = progress;
                            }
                        } while (isMoreToRead);
                    }
                }
            }
            catch (TaskCanceledException)
            {
                if (File.Exists(destinationFilePath))
                    try
                    {
                        File.Delete(destinationFilePath);
                    }
                    catch (Exception ex)
                    {
                        _ = MessageBox.Show($"Error deleting file: {ex.Message}");
                    }
            }
            catch (HttpRequestException httpRequestEx)
            {
                _ = MessageBox.Show($"HTTP request error: {httpRequestEx.Message}");
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"Error during download: {ex.Message}");
                CleanUpFile(destinationFilePath);
            }
        }

        private void CancelDownload_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource?.Cancel();
        }

        private async SystemTask MonitorProcess(Process process)
        {
            while (!process.HasExited)
            {
                await Dispatcher.InvokeAsync(SetInstalledVersion);
                await SystemTask.Delay(1000);
            }
        }

        private void SetInstalledVersion()
        {
            InstalledVersionString = GetInstalledAmdChipsetVersion() ?? "Not Installed";
        }

        private void CleanUpFile(string destinationFilePath)
        {
            if (!File.Exists(destinationFilePath)) return;

            try
            {
                File.Delete(destinationFilePath);
            }
            catch (Exception deleteEx)
            {
                _ = MessageBox.Show($"Error deleting file: {deleteEx.Message}");
            }
        }
    }
}