using System.Diagnostics;
using NLog;

namespace Rubberduck.Parsing.Common
{
    public class ParsingStageTimer
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static ParsingStageTimer StartNew()
        {
            var timer = new ParsingStageTimer();
            timer.Start();
            return timer;
        }

        private readonly Stopwatch _stopwatch = new Stopwatch();

        public void Start() => _stopwatch.Start();
        public void Stop() => _stopwatch.Stop();
        public void Restart() => _stopwatch.Restart();

        public void Log(string message)
        {
            Logger.Info(message, _stopwatch.ElapsedMilliseconds);
        }
    }
}