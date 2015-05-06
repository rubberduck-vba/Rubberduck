using NLog;
using NLog.Config;
using NLog.Targets;

namespace Rubberduck.Logging
{
    public static class LoggingConfigurator
    {
        public static void ConfigureParserLogger()
        {
            var config = new LoggingConfiguration();

            var target = new FileTarget
            {
                FileName = "${specialfolder:folder=ApplicationData}/Rubberduck/logs/parser.log",
                Layout = "${longdate} ${uppercase:${level}} ${newline} ${message} ${newline} ${newline} ${exception:format=tostring} ${newline}"
            };
            config.AddTarget("parser", target);

            var rule = new LoggingRule("*", LogLevel.Trace, target);
            config.LoggingRules.Add(rule);

            LogManager.ThrowExceptions = true;
            LogManager.Configuration = config;
        }
    }
}
