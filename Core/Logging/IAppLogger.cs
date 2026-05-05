// FileCollector V3.0 - logger abstraction. Lets Core/ services log without depending on WinForms.

namespace FileCollector.Core.Logging
{
    public enum LogLevel
    {
        Debug,
        Info,
        Warn,
        Error
    }

    public interface IAppLogger
    {
        void Log(LogLevel level, string message);
    }

    public static class LoggerExtensions
    {
        public static void Debug(this IAppLogger l, string msg) => l.Log(LogLevel.Debug, msg);
        public static void Info(this IAppLogger l, string msg)  => l.Log(LogLevel.Info,  msg);
        public static void Warn(this IAppLogger l, string msg)  => l.Log(LogLevel.Warn,  msg);
        public static void Error(this IAppLogger l, string msg) => l.Log(LogLevel.Error, msg);
    }
}
