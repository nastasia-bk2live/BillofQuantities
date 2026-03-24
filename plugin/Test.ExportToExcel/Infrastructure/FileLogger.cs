using System;
using System.IO;
using System.Text;

namespace Test.ExportToExcel.Infrastructure
{
    /// <summary>
    /// Простой файловый логгер для фиксации ошибок и ключевых шагов экспорта.
    /// </summary>
    public class FileLogger
    {
        private readonly string _logFilePath;

        public FileLogger()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var logDirectory = Path.Combine(appData, "Test.ExportToExcel");
            Directory.CreateDirectory(logDirectory);
            _logFilePath = Path.Combine(logDirectory, "export.log");
        }

        public void Info(string message)
        {
            Write("INFO", message);
        }

        public void Error(string message, Exception exception = null)
        {
            var details = exception == null ? message : message + Environment.NewLine + exception;
            Write("ERROR", details);
        }

        private void Write(string level, string message)
        {
            var builder = new StringBuilder();
            builder.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            builder.Append(" [");
            builder.Append(level);
            builder.Append("] ");
            builder.AppendLine(message);

            File.AppendAllText(_logFilePath, builder.ToString(), Encoding.UTF8);
        }
    }
}
