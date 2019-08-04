using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged
{
    /// <summary>
    /// An internal representation of the <see cref="ITypeLib"/> object hosted by the VBE.
    /// Also provides Prev/Next pointers, exposing a double linked list of all loaded project ITypeLibs
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct VBETypeLibObj
    {
        IntPtr _vTable1;     // ITypeLib vtable
        IntPtr _vTable2;
        IntPtr _vTable3;
        public IntPtr Prev;
        public IntPtr Next;
    }

    /// <summary>
    /// An enumerable class for iterating over the double linked list of <see cref="ITypeLib"/>s provided by the VBE 
    /// </summary>
    internal sealed class VBETypeLibsIterator : IEnumerable<ITypeLibWrapper>, IEnumerator<ITypeLibWrapper>
    {
        private IntPtr _currentTypeLibPtr;
        private VBETypeLibObj _currentTypeLibStruct;
        private bool _isStart;

        public VBETypeLibsIterator(IntPtr typeLibPtr)
        {
            _currentTypeLibPtr = typeLibPtr;
            _currentTypeLibStruct = StructHelper.ReadStructureSafe<VBETypeLibObj>(_currentTypeLibPtr);
            Reset();
        }

        public void Dispose()
        {
            // nothing to do here, we don't own anything that needs releasing
        }

        IEnumerator IEnumerable.GetEnumerator() => this;
        public IEnumerator<ITypeLibWrapper> GetEnumerator() => this;

        ITypeLibWrapper IEnumerator<ITypeLibWrapper>.Current => TypeApiFactory.GetTypeLibWrapper(_currentTypeLibPtr, addRef: true);
        object IEnumerator.Current => TypeApiFactory.GetTypeLibWrapper(_currentTypeLibPtr, addRef: true);

        public void Reset()  // walk back to the first project in the chain
        {
            while (_currentTypeLibStruct.Prev != IntPtr.Zero)
            {
                _currentTypeLibPtr = _currentTypeLibStruct.Prev;
                _currentTypeLibStruct = StructHelper.ReadStructureSafe<VBETypeLibObj>(_currentTypeLibPtr);
            }
            _isStart = true;
        }

        public bool MoveNext()
        {
            if (_isStart)
            {
                _isStart = false;  // MoveNext is called before accessing the very first item
                return true;
            }

            if (_currentTypeLibStruct.Next == IntPtr.Zero) return false;

            _currentTypeLibPtr = _currentTypeLibStruct.Next;
            _currentTypeLibStruct = StructHelper.ReadStructureSafe<VBETypeLibObj>(_currentTypeLibPtr);
            return true;
        }
    }
}
