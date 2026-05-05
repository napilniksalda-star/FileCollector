// FileCollector V3.0 - logger that batches log lines on a timer to avoid 1-Invoke-per-message.
// V2.1 called Control.Invoke for every LogMessage from worker threads inside a tight loop -
// when the loop logged thousands of times the UI thread became the bottleneck.

using System;
using System.Collections.Concurrent;
using System.Text;
using System.Windows.Forms;

namespace FileCollector.Core.Logging
{
    public sealed class BatchUiLogger : IAppLogger, IDisposable
    {
        private readonly Action<string> _appendOnUi;
        private readonly System.Windows.Forms.Timer _timer;
        private readonly ConcurrentQueue<string> _queue = new();
        private LogLevel _minLevel;
        private bool _disposed;

        /// <summary>
        /// Creates a batched UI logger. Must be constructed on the UI thread (the timer ticks on it).
        /// </summary>
        /// <param name="appendOnUi">Receives multi-line text to append to the UI log control.</param>
        /// <param name="flushIntervalMs">How often pending lines are flushed to the UI.</param>
        /// <param name="minLevel">Minimum level to emit. Debug is suppressed by default in V3.0.</param>
        public BatchUiLogger(Action<string> appendOnUi, int flushIntervalMs = 120, LogLevel minLevel = LogLevel.Info)
        {
            _appendOnUi = appendOnUi ?? throw new ArgumentNullException(nameof(appendOnUi));
            _minLevel = minLevel;
            _timer = new System.Windows.Forms.Timer { Interval = flushIntervalMs };
            _timer.Tick += (_, __) => Flush();
            _timer.Start();
        }

        public LogLevel MinimumLevel
        {
            get => _minLevel;
            set => _minLevel = value;
        }

        public void Log(LogLevel level, string message)
        {
            if (level < _minLevel) return;
            _queue.Enqueue($"{DateTime.Now:HH:mm:ss} [{LevelTag(level)}] {message}");
        }

        public void Flush()
        {
            if (_queue.IsEmpty) return;

            // Drain up to N lines per tick - keeps UI responsive when the queue is large.
            const int maxPerTick = 500;
            var sb = new StringBuilder();
            int n = 0;
            while (n < maxPerTick && _queue.TryDequeue(out var line))
            {
                sb.Append(line).Append(Environment.NewLine);
                n++;
            }
            if (sb.Length > 0)
                _appendOnUi(sb.ToString());
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _timer.Stop();
            _timer.Dispose();
            // One last flush so we don't lose buffered messages.
            try { Flush(); } catch { /* swallow during shutdown */ }
        }

        private static string LevelTag(LogLevel l) => l switch
        {
            LogLevel.Debug => "DBG",
            LogLevel.Info  => "INF",
            LogLevel.Warn  => "WRN",
            LogLevel.Error => "ERR",
            _              => "???"
        };
    }
}
