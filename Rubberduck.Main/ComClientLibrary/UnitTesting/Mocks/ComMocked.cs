using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Moq;
using NLog;
using Rubberduck.Resources.Registration;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IComMockedGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IComMocked : IMocked
    {
        new ComMock Mock { get; }
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.ComMockedGuid),
        ProgId(RubberduckProgId.ComMockedProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IComMocked))
    ]
    public class ComMocked : IComMocked, ICustomQueryInterface
    {
        private static readonly ILogger Logger = LogManager.GetCurrentClassLogger();
        private readonly IEnumerable<Type> _supportedTypes;

        internal ComMocked(ComMock mock, IEnumerable<Type> supportedTypes)
        {
            Mock = mock;
            _supportedTypes = supportedTypes;
        }

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            try
            {
                var result = IntPtr.Zero;
                var searchIid = iid; // Cannot use ref parameters directly in LINQ

                if (iid == new Guid(RubberduckGuid.ComMockedGuid))
                {
                    result = Marshal.GetIUnknownForObject(this);
                }
                else
                {
                    // Apparently some COM objects have multiple interface implementations using same GUID
                    // so first result should suffice to avoid exception when using single. 
                    var type = _supportedTypes.FirstOrDefault(x => x.GUID == searchIid);
                    if (type != null)
                    {
                        result = Marshal.GetComInterfaceForObject(Mock.Mock.Object, type);
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

        public ComMock Mock { get; }

        Mock IMocked.Mock => Mock.Mock;
    }
}