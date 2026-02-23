using System;
using System.IO;

namespace Mdr.Revit.Infra.Logging
{
    public sealed class PluginLogger
    {
        private readonly object _sync = new object();
        private readonly string _logDirectory;

        public PluginLogger(string logDirectory)
        {
            if (string.IsNullOrWhiteSpace(logDirectory))
            {
                throw new ArgumentException("Log directory is required.", nameof(logDirectory));
            }

            _logDirectory = logDirectory;
            Directory.CreateDirectory(_logDirectory);
        }

        public void Info(string message)
        {
            Write("INFO", message);
        }

        public void Error(string message)
        {
            Write("ERROR", message);
        }

        private void Write(string level, string message)
        {
            string safeMessage = message ?? string.Empty;
            string correlation = CorrelationContext.CurrentRunUid;
            string line =
                DateTimeOffset.UtcNow.ToString("o") +
                " [" + level + "]" +
                (string.IsNullOrWhiteSpace(correlation) ? string.Empty : " [run=" + correlation + "]") +
                " " + safeMessage;

            string filePath = Path.Combine(_logDirectory, DateTime.UtcNow.ToString("yyyyMMdd") + ".log");

            lock (_sync)
            {
                File.AppendAllText(filePath, line + Environment.NewLine);
            }
        }
    }
}
