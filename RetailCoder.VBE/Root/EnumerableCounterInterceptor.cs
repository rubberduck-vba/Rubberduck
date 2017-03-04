using System;
using System.Collections.Generic;
using System.Linq;
using Ninject.Extensions.Interception;
using Ninject.Infrastructure.Language;
using NLog;

namespace Rubberduck.Root
{
    /// <summary>
    /// An attribute that makes an intercepted method call log the number of items returned.
    /// </summary>
    public class EnumerableCounterInterceptAttribute : Attribute { }

    /// <summary>
    /// An interceptor that logs the number of items returned by an intercepted invocation that returns any IEnumerable{T}.
    /// </summary>
    public class EnumerableCounterInterceptor<T> : InterceptorBase
    {
        private readonly ILogger _logger = LogManager.GetCurrentClassLogger();

        protected override void AfterInvoke(IInvocation invocation)
        {
            if (!invocation.Request.Method.HasAttribute<EnumerableCounterInterceptAttribute>()) { return; }

            var result = invocation.ReturnValue as IEnumerable<T>;
            if (result != null)
            {
                _logger.Trace("Intercepted invocation of '{0}.{1}' returned {2} objects.",
                    invocation.Request.Target.GetType().Name, invocation.Request.Method.Name, result.Count());
            }
        }
    }
}