using System;
using Ninject.Extensions.Interception;
using NLog;

namespace Rubberduck.Root
{
    /// <summary>
    /// An interceptor that logs an unhandled exception.
    /// </summary>
    public class FatalExceptionInterceptor : InterceptorBase
    {
        private readonly ILogger _logger = LogManager.GetCurrentClassLogger();

        protected override void OnError(IInvocation invocation, Exception exception)
        {
            _logger.Fatal(exception);
            throw new InterceptedException(exception);
        }
    }
}