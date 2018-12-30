using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NLog;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    public class ComMocked : ICustomQueryInterface
    {
        private static readonly ILogger Logger = LogManager.GetCurrentClassLogger();
        private readonly object _target;
        private readonly IEnumerable<Type> _supportedTypes;

        internal ComMocked(object target, IEnumerable<Type> supportedTypes)
        {
            _target = target;
            _supportedTypes = supportedTypes;
        }

        private static readonly Guid iidIDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            try
            {
                var result = IntPtr.Zero;
                var searchIid = iid;

                if (iid == iidIDispatch)
                {
                    // TODO: find a better way to get IDispatch - it is not possible to return
                    // a IDispatch because it is in a object hierarchy where the parent object
                    // is not COM visible. Returning IUnknown seems to work because of the happy
                    // accident that the IDispatch's v-table usually follows IUnknown's v-table.
                    result = Marshal.GetIUnknownForObject(_target);
                }
                else
                {
                    var type = _supportedTypes.SingleOrDefault(x => x.GUID == searchIid);
                    if (type != null)
                    {
                        result = Marshal.GetComInterfaceForObject(_target, type);
                    }
                }

                ppv = result;
                return result == IntPtr.Zero
                    ? CustomQueryInterfaceResult.NotHandled
                    : CustomQueryInterfaceResult.Handled;
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, $"Failed to perform IQueryInterface call on {nameof(ComMocked)}. IID requested was {{{iid}}}.");
                ppv = IntPtr.Zero;
                return CustomQueryInterfaceResult.Failed;
            }
        }
    }
}