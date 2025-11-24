using CommunityToolkit.Mvvm.ComponentModel;
using System;
using Tmds.DBus.Protocol;

namespace DefectClassification.GUI.Models
{
    /// <summary>
    /// Представляет запись в журнале событий
    /// </summary>
    public partial class LogEntry : ObservableObject
    {
        [ObservableProperty]
        private DateTime _timestamp;

        [ObservableProperty]
        private string _message = string.Empty;

        [ObservableProperty]
        private LogLevel _level;

        [ObservableProperty]
        private string _color = "#000000";

        public LogEntry()
        {
            Timestamp = DateTime.Now;
        }

        public LogEntry(string message, LogLevel level, string color)
        {
            Timestamp = DateTime.Now;
            Message = message;
            Level = level;
            Color = color;
        }

        /// <summary>
        /// Форматированная строка времени
        /// </summary>
        public string TimeStamp => Timestamp.ToString("HH:mm:ss");

        /// <summary>
        /// Форматированная полная запись
        /// </summary>
        public string FormattedMessage => $"[{TimeStamp}] {Message}";
    }

    /// <summary>
    /// Уровень важности лога
    /// </summary>
    public enum LogLevel
    {
        Normal,
        Info,
        Success,
        Warning,
        Error
    }
}