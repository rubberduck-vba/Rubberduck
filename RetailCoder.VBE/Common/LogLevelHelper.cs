using System.Linq;
using NLog;
using NLog.Config;
using System;
using System.Collections.Generic;

namespace Rubberduck.Common
{
    public static class LogLevelHelper
    {
        private static readonly Lazy<IEnumerable<LogLevel>> _logLevels = new Lazy<IEnumerable<LogLevel>>(GetLogLevels);

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static string LogHeader;
        private static bool LogHeaderWritten;

        public static IEnumerable<LogLevel> LogLevels
        {
            get
            {
                return _logLevels.Value;
            }
        }

        private static IEnumerable<LogLevel> GetLogLevels()
        {
            var logLevels = new List<LogLevel>();
            logLevels.Add(LogLevel.Off);
            for (int logLevelOrdinal = 0; logLevelOrdinal <= 5; logLevelOrdinal++)
            {
                logLevels.Add(LogLevel.FromOrdinal(logLevelOrdinal));
            }
            return logLevels;
        }

        public static int MinLogLevel()
        {
            return GetLogLevels().Min(lvl => lvl.Ordinal);
        }

        public static int MaxLogLevel()
        {
            return GetLogLevels().Max(lvl => lvl.Ordinal);
        }

        public static void SetDebugInfo(String value)
        {
            LogHeader = value;
            LogHeaderWritten = false;
        }

        public static void SetMinimumLogLevel(LogLevel minimumLogLevel)
        {
            if (LogManager.GlobalThreshold == minimumLogLevel && LogHeaderWritten == true)
            {
                return;
            }
            if (LogHeaderWritten == true)
            {
                Logger.Log(LogLevel.Info, "Minimum log level changing from " + 
                    LogManager.GlobalThreshold.Name +
                    " to " + minimumLogLevel.Name);
            }
            var loggingRules = LogManager.Configuration.LoggingRules;
            foreach (var loggingRule in loggingRules)
            {
                ClearLogLevels(loggingRule);
            }
            if (minimumLogLevel == LogLevel.Off)
            {
                LogManager.DisableLogging();
                LogManager.GlobalThreshold = LogLevel.Off;
                LogManager.ReconfigExistingLoggers();
                return;
            }
            LogManager.EnableLogging();
            foreach (var loggingRule in loggingRules)
            {
                foreach (var logLevel in LogLevels)
                {
                    if (logLevel != LogLevel.Off && logLevel >= minimumLogLevel)
                    {
                        loggingRule.EnableLoggingForLevel(logLevel);
                    }
                }
            }
            LogManager.GlobalThreshold = minimumLogLevel;
            LogManager.ReconfigExistingLoggers();
            if (LogHeaderWritten == false)
            {
                Logger.Log(minimumLogLevel, LogHeader);
                LogHeaderWritten = true;
            }
        }

        private static void ClearLogLevels(LoggingRule loggingRule)
        {
            foreach (var logLevel in LogLevels)
            {
                if (logLevel != LogLevel.Off && loggingRule.IsLoggingEnabledForLevel(logLevel))
                {
                    loggingRule.DisableLoggingForLevel(logLevel);
                }
            }
        }
    }
}
