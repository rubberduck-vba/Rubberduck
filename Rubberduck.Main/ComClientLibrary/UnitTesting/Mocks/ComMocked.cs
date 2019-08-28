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
        private readonly IEnumerable<Type> _supportedInterfaces;
        
        // Not using auto-property as that leads to ambiguity. For COM compatibility,
        // this backs the public field `Mock`, which hides the `Moq.IMocked.Mock` 
        // property that returns a non-COM-visible `Moq.Mock` object. 
        private readonly ComMock _comMock; 

        internal ComMocked(ComMock mock, IEnumerable<Type> supportedInterfaces)
        {
            _comMock = mock;
            _supportedInterfaces = supportedInterfaces;
        }

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            try
            {
                var result = IntPtr.Zero;
                var searchIid = iid; // Cannot use ref parameters directly in LINQ

                if (iid == new Guid(RubberduckGuid.ComMockedGuid) || iid == new Guid(RubberduckGuid.IComMockedGuid))
                {
                    result = Marshal.GetIUnknownForObject(this);
                }
                else if (iid == new Guid(RubberduckGuid.IID_IDispatch) && !string.IsNullOrWhiteSpace(Mock.Project))
                {
                    // We cannot return IDispatch directly for VBA types but we can return the IUnknown in its place,
                    // which is sufficient for COM's needs. 

                    var pObject = Marshal.GetComInterfaceForObject(_comMock.Mock.Object, _comMock.MockedType);
                    searchIid = new Guid(RubberduckGuid.IID_IUnknown);
                    var hr = Marshal.QueryInterface(pObject, ref searchIid, out result);
                    Marshal.Release(pObject);
                    if (hr < 0)
                    {
                        ppv = IntPtr.Zero;
                        return CustomQueryInterfaceResult.Failed;
                    }
                }
                else
                {
                    // Apparently some COM objects have multiple interface implementations using same GUID
                    // so first result should suffice to avoid exception when using single. 
                    var type = _supportedInterfaces.FirstOrDefault(x => x.GUID == searchIid);
                    if (type != null)
                    {
                        // Ensure that we return the actual Moq.Mock.Object, not the ComMocked object.
                        result = Marshal.GetComInterfaceForObject(_comMock.Mock.Object, type);
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

        // ReSharper disable once ConvertToAutoPropertyWhenPossible -- Leads to ambiguous naming; see comments above
        public ComMock Mock => _comMock;

        Mock IMocked.Mock => _comMock.Mock;
    }
}