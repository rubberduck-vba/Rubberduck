using Rubberduck.Com.Extensions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using static Vanara.PInvoke.OleAut32;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using PARAMDESC = System.Runtime.InteropServices.ComTypes.PARAMDESC;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.Com
{
    public abstract class TypeApiWalker<T>
    {
        protected IEnumerable<T> Visitors;
        
        protected void ExecuteVisit(Action<T> action)
        {
            foreach (var visitor in Visitors)
            {
                action.Invoke(visitor);
            }
        }

        protected void EnumerateCustomData(Action<IntPtr> getAllCustData, Action<Guid, object> enumerationAction)
        {
            var ptrCustData = IntPtr.Zero;
            getAllCustData.Invoke(ptrCustData);

            if(ptrCustData == IntPtr.Zero)
            {
                return;
            }

            var custData = Marshal.PtrToStructure<CUSTDATA>(ptrCustData);
            var offset = Marshal.SizeOf(custData.cCustData);
            var ptrSize = Marshal.SizeOf(typeof(IntPtr));

            for (var i = 0; i < custData.cCustData; i++)
            {
                var ptrCustDataItem = ptrCustData + offset + (i * ptrSize);
                var guid = Marshal.PtrToStructure<Guid>(ptrCustDataItem);
                var ptrValue = ptrCustDataItem + Marshal.SizeOf(typeof(Guid));
                var value = Marshal.GetObjectForNativeVariant(ptrValue);

                enumerationAction.Invoke(guid, value);
            }
            ClearCustData(ref custData);
        }
    }
}
