using System;
using System.Diagnostics;
using Ninject.Extensions.Interception;
using Ninject.Infrastructure.Language;
using NLog;

namespace Rubberduck.Root
{
    /// <summary>
    /// An attribute that makes an intercepted method call log the duration of its execution.
    /// </summary>
    public class TimedCallInterceptAttribute : Attribute { }

    /// <summary>
    /// An interceptor that logs the duration of an intercepted invocation.
    /// </summary>
    public class TimedCallLoggerInterceptor : InterceptorBase
    {
        private readonly ILogger _logger = LogManager.GetCurrentClassLogger();
        private readonly Stopwatch _stopwatch = new Stopwatch();
        private bool _running;

        protected override void BeforeInvoke(IInvocation invocation)
        {
            _running = invocation.Request.Method.HasAttribute<TimedCallInterceptAttribute>();
            if(!_running) { return; }

            _stopwatch.Reset();
            _stopwatch.Start();
        }

        protected override void AfterInvoke(IInvocation invocation)
        {
            if (!_running) { return; }

            _stopwatch.Stop();
            _logger.Trace("Intercepted invocation of '{0}.{1}' ran for {2}ms",
                invocation.Request.Target.GetType().Name, invocation.Request.Method.Name, _stopwatch.ElapsedMilliseconds);
        }
    }
}
