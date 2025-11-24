using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DefectClassification.GUI.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DefectClassification.GUI.ViewModels
{
    public partial class MainWindowViewModel : ViewModelBase
    {
        [ObservableProperty]
        private string _filePath = string.Empty;

        [ObservableProperty]
        private double _lambdaValue = 10.0;

        [ObservableProperty]
        private bool _isStaticMode = true;

        [ObservableProperty]
        private bool _isDynamicMode = false;

        [ObservableProperty]
        private bool _createBackup = true;

        [ObservableProperty]
        private double _progressValue = 0;

        [ObservableProperty]
        private string _statusMessage = "Готов к работе";

        [ObservableProperty]
        private IBrush _statusMessageBrush = Brushes.Black;


        [ObservableProperty]
        private string _outputFilePath = string.Empty;

        [ObservableProperty]
        private bool _hasError = false;

        [ObservableProperty]
        private bool _isProcessing = false;

        [ObservableProperty]
        private bool _canProcess = false;

        [ObservableProperty]
        private bool _isDone = false;

        [ObservableProperty]
        private ObservableCollection<TubeConfiguration> _tubeConfigurations = new();

        [ObservableProperty]
        private TubeConfiguration? _selectedTubeConfiguration;

        [ObservableProperty]
        private int _newTubeNumber = 1;

        [ObservableProperty]
        private double _newWallThickness = 10.0;
        [ObservableProperty]
        private ObservableCollection<LogEntry> _logEntries = new();

        [ObservableProperty]
        private bool _hasLogs = false;

        // Color definitions
        private static readonly IBrush SuccessBrush = new SolidColorBrush(Color.FromRgb(34, 139, 34));  // Green
        private static readonly IBrush ErrorBrush = new SolidColorBrush(Color.FromRgb(220, 20, 60));     // Red
        private static readonly IBrush WarningBrush = new SolidColorBrush(Color.FromRgb(255, 140, 0));   // Orange
        private static readonly IBrush InfoBrush = new SolidColorBrush(Color.FromRgb(0, 100, 200));      // Blue
        private static readonly IBrush NormalBrush = Brushes.Black;                                      // Black

        private const string SuccessColor = "#228B22";
        private const string ErrorColor = "#DC143C";
        private const string WarningColor = "#FF8C00";
        private const string InfoColor = "#0064C8";
        private const string NormalColor = "#000000";

        public MainWindowViewModel()
        {
            // Initialize with some default tube configurations
            InitializeDefaultTubes();
            AddLog("Приложение запущено", LogLevel.Info);

        }

        private void InitializeDefaultTubes()
        {
            // Add 10 default tubes as examples
            for (int i = 1; i <= 10; i++)
            {
                TubeConfigurations.Add(new TubeConfiguration(i, 10.0));
            }
        }

        [RelayCommand]
        private void ClearLogs()
        {
            LogEntries.Clear();
            HasLogs = false;
            AddLog("Журнал очищен", LogLevel.Info);
        }
        [RelayCommand]
        private async Task ExportLogs()
        {
            if (LogEntries.Count == 0)
            {
                SetStatusMessage("⚠ Журнал пуст - нечего экспортировать", MessageType.Warning);
                return;
            }

            try
            {
                var topLevel = TopLevel.GetTopLevel(App.MainWindow);
                if (topLevel == null) return;

                var file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
                {
                    Title = "Сохранить журнал",
                    DefaultExtension = ".txt",
                    SuggestedFileName = $"log_{DateTime.Now:yyyyMMdd_HHmmss}.txt",
                    FileTypeChoices = new[]
                    {
                        new FilePickerFileType("Text Files") { Patterns = new[] { "*.txt" } }
                    }
                });

                if (file != null)
                {
                    var logText = string.Join(Environment.NewLine,
                        LogEntries.Select(e => e.FormattedMessage));

                    await File.WriteAllTextAsync(file.Path.LocalPath, logText);

                    SetStatusMessage($"✓ Журнал сохранён: {Path.GetFileName(file.Path.LocalPath)}", MessageType.Success);
                }
            }
            catch (Exception ex)
            {
                SetStatusMessage($"❌ Ошибка сохранения журнала: {ex.Message}", MessageType.Error);
            }
        }

        private void AddLog(string message, LogLevel level)
        {
            var color = level switch
            {
                LogLevel.Success => SuccessColor,
                LogLevel.Error => ErrorColor,
                LogLevel.Warning => WarningColor,
                LogLevel.Info => InfoColor,
                _ => NormalColor
            };

            var entry = new LogEntry(message, level, color);

            // Add on UI thread
            Dispatcher.UIThread.Post(() =>
            {
                LogEntries.Add(entry);
                HasLogs = LogEntries.Count > 0;
            });
        }

        [RelayCommand]
        private void AddTube()
        {
            // Check if tube number already exists
            if (TubeConfigurations.Any(t => t.TubeNumber == NewTubeNumber))
            {
                SetStatusMessage($"❌ Трубка #{NewTubeNumber} уже существует в списке", MessageType.Error);
                return;
            }

            TubeConfigurations.Add(new TubeConfiguration(NewTubeNumber, NewWallThickness));

            // Sort by tube number
            var sorted = TubeConfigurations.OrderBy(t => t.TubeNumber).ToList();
            TubeConfigurations.Clear();
            foreach (var tube in sorted)
            {
                TubeConfigurations.Add(tube);
            }

            SetStatusMessage($"✓ Добавлена трубка #{NewTubeNumber} с толщиной стенки {NewWallThickness} мм", MessageType.Success);

            // Increment for next tube
            NewTubeNumber++;
        }

        [RelayCommand]
        private void RemoveTube()
        {
            if (SelectedTubeConfiguration != null)
            {
                var tubeNum = SelectedTubeConfiguration.TubeNumber;
                TubeConfigurations.Remove(SelectedTubeConfiguration);
                SetStatusMessage($"✓ Удалена трубка #{tubeNum}", MessageType.Info);
                SelectedTubeConfiguration = null;
            }
            else
            {
                SetStatusMessage("⚠ Выберите трубку для удаления", MessageType.Warning);
            }
        }

        [RelayCommand]
        private void ClearAllTubes()
        {
            if (TubeConfigurations.Count == 0)
            {
                SetStatusMessage("⚠ Список трубок уже пуст", MessageType.Warning);
                return;
            }

            TubeConfigurations.Clear();
            SetStatusMessage("✓ Все трубки удалены", MessageType.Info);
        }

        [RelayCommand]
        private async Task SelectFile()
        {
            try
            {
                var topLevel = TopLevel.GetTopLevel(App.MainWindow);
                if (topLevel == null)
                {
                    SetStatusMessage("❌ Ошибка инициализации окна", MessageType.Error);
                    return;
                }

                var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
                {
                    Title = "Выберите Excel файл",
                    AllowMultiple = false,
                    FileTypeFilter = new[]
                    {
                        new FilePickerFileType("Excel Files")
                        {
                            Patterns = new[] { "*.xlsx" }
                        }
                    }
                });

                if (files.Count > 0)
                {
                    FilePath = files[0].Path.LocalPath;

                    // Validate file exists and is readable
                    if (!File.Exists(FilePath))
                    {
                        SetStatusMessage($"❌ Файл не найден: {Path.GetFileName(FilePath)}", MessageType.Error);
                        CanProcess = false;
                        return;
                    }



                    var extension = Path.GetExtension(FilePath).ToLower();
                    if (extension != ".xlsx")
                    {
                        SetStatusMessage($"❌ Неверный формат файла: {extension}(поддерживается только .xsls", MessageType.Error);
                        CanProcess = false;
                        return;
                    }
                       CanProcess = true;
                    IsDone = false;

                }
            }
            catch (Exception ex)
            {
                SetStatusMessage($"❌ Ошибка выбора файла: {ex.Message}", MessageType.Error);
                CanProcess = false;
            }
        }

        [RelayCommand]
        private async Task ProcessFile()
        {
            SetStatusMessage("▶ Начинаю обработку...", MessageType.Info);

            // Validate file selection
            if (string.IsNullOrWhiteSpace(FilePath))
            {
                SetStatusMessage("❌ Пожалуйста, выберите файл", MessageType.Error);
                return;
            }
            var extension = Path.GetExtension(FilePath).ToLower();
            if (extension != ".xlsx")
            {
                SetStatusMessage($"❌ Неверный формат файла: {extension}(поддерживается только .xsls", MessageType.Error);
                CanProcess = false;
                return;
            }
            if (!File.Exists(FilePath))
            {
                SetStatusMessage($"❌ Файл не найден: {Path.GetFileName(FilePath)}", MessageType.Error);
                CanProcess = false;
                return;
            }

            // Validate dynamic mode has tube configurations
            if (IsDynamicMode && TubeConfigurations.Count == 0)
            {
                SetStatusMessage("❌ Добавьте хотя бы одну конфигурацию трубки для динамического режима", MessageType.Error);
                return;
            }

            IsProcessing = true;
            CanProcess = false;
            ProgressValue = 0;
            HasError = false;

            try
            {
                // Check file access before processing

                    SetStatusMessage("Проверяю доступность файла...", MessageType.Info);



                var processor = new Services.ExcelProcessor();

                var progress = new Progress<(int current, int total, string message)>(p =>
                {
                    ProgressValue = p.total > 0 ? (double)p.current / p.total * 100 : 0;

                    // Detect error/warning messages and color them
                    if (p.message.Contains("❌") || p.message.ToLower().Contains("ошибка"))
                    {
                        SetStatusMessage(p.message, MessageType.Error);
                    }
                    else if (p.message.Contains("⚠") || p.message.Contains("⚠️"))
                    {
                        SetStatusMessage(p.message, MessageType.Warning);
                    }
                    else
                    {
                        SetStatusMessage(p.message, MessageType.Info);
                    }
                });

                // Create tube thickness dictionary for dynamic mode
                Dictionary<int, double>? tubeThicknesses = null;
                if (IsDynamicMode)
                {
                    tubeThicknesses = TubeConfigurations.ToDictionary(
                        t => t.TubeNumber,
                        t => t.WallThickness);
                }

                await Task.Run(() => processor.ProcessFile(
                    FilePath,
                    LambdaValue,
                    IsStaticMode,
                    CreateBackup,
                    tubeThicknesses,
                    progress));

                SetStatusMessage("✅ Обработка завершена успешно!", MessageType.Success);
                ProgressValue = 100;
                IsDone = true;
                OutputFilePath = FilePath;

            }
            catch (FileNotFoundException ex)
            {
                SetStatusMessage($"❌ Файл не найден: {ex.Message}", MessageType.Error);
                ProgressValue = 0;
            }
            catch (UnauthorizedAccessException)
            {
                SetStatusMessage($"❌ Недостаточно прав для доступа к файлу", MessageType.Error);
                ProgressValue = 0;
            }
            catch (IOException ex)
            {
                if (ex.Message.Contains("being used by another process"))
                {
                    SetStatusMessage($"❌ Файл открыт в другой программе. Закройте Excel и попробуйте снова.", MessageType.Error);
                }
                else
                {
                    SetStatusMessage($"❌ Ошибка ввода/вывода: {ex.Message}", MessageType.Error);
                }
                ProgressValue = 0;
            }
            catch (InvalidOperationException ex)
            {
                SetStatusMessage($"❌ Ошибка операции: {ex.Message}", MessageType.Error);
                ProgressValue = 0;
            }
            catch (Exception ex)
            {
                SetStatusMessage($"❌ Неожиданная ошибка: {ex.Message}", MessageType.Error);
                ProgressValue = 0;
            }
            finally
            {
                IsProcessing = false;
                CanProcess = !string.IsNullOrWhiteSpace(FilePath) && File.Exists(FilePath);
            }
        }

        [RelayCommand]
        private void OpenFile()
        {
            SetStatusMessage("ℹ Открытие файла...", MessageType.Info);
            if (string.IsNullOrWhiteSpace(OutputFilePath) || !File.Exists(OutputFilePath))
            {
                SetStatusMessage("❌ Файл не найден. Выполните обработку сначала.", MessageType.Error);
                return;
            }

            try
            {
                // Cross-platform file opening
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    // Windows: use explorer
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = OutputFilePath,
                        UseShellExecute = true
                    });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    // Linux: use xdg-open
                    Process.Start("xdg-open", OutputFilePath);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    // macOS: use open
                    Process.Start("open", OutputFilePath);
                }

                SetStatusMessage($"✓ Открыт файл: {Path.GetFileName(OutputFilePath)}", MessageType.Success);
            }
            catch (Exception ex)
            {
                SetStatusMessage($"❌ Ошибка открытия файла: {ex.Message}", MessageType.Error);
            }
        }


        partial void OnIsStaticModeChanged(bool value)
        {
            if (value)
            {
                IsDynamicMode = false;
                SetStatusMessage("ℹ Выбран статический режим - единая толщина стенки для всех трубок", MessageType.Info);
            }
        }

        partial void OnIsDynamicModeChanged(bool value)
        {
            if (value)
            {
                IsStaticMode = false;
                SetStatusMessage("ℹ Выбран динамический режим - индивидуальная толщина стенки для каждой трубки", MessageType.Info);
            }
        }

        private void SetStatusMessage(string message, MessageType type)
        {
            StatusMessage = message;
            HasError = type == MessageType.Error;

            StatusMessageBrush = type switch
            {
                MessageType.Success => SuccessBrush,
                MessageType.Error => ErrorBrush,
                MessageType.Warning => WarningBrush,
                MessageType.Info => InfoBrush,
                MessageType.Normal => NormalBrush,
                _ => NormalBrush
            };

            // Also add to log history
            var logLevel = type switch
            {
                MessageType.Success => LogLevel.Success,
                MessageType.Error => LogLevel.Error,
                MessageType.Warning => LogLevel.Warning,
                MessageType.Info => LogLevel.Info,
                _ => LogLevel.Normal
            };

            AddLog(message, logLevel);
        }

        private enum MessageType
        {
            Normal,
            Success,
            Error,
            Warning,
            Info
        }
    }
}