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

        public static void SetMinimumLogLevel(LogLevel minimumLogLevel)
        {
            var loggingRules = LogManager.Configuration.LoggingRules;
            foreach (var loggingRule in loggingRules)
            {
                ClearLogLevels(loggingRule);
            }
            if (minimumLogLevel == LogLevel.Off)
            {
                LogManager.DisableLogging();
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
            LogManager.ReconfigExistingLoggers();
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
