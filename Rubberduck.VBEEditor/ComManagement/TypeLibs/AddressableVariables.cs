using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    // AddressableVariables are created in unmanaged memory space, designed to aid creating addressable
    // content when passing to/from interop code (as IntPtr addresses)

    // IAddressableVariableBase<T> can  represent a single element, or a contiguous array,
    // and allows the derived classes to implement marshalling for the elements (e.g. string<->BSTR etc)
    public abstract class IAddressableVariableBase<TUnmarshalled, TMarshalled> : IDisposable
    {
        public readonly IntPtr _address;
        public readonly int _elementCount;       // 1 for singular elements
        private readonly bool _ownedMemory;      // true if WE allocated the memory (default)

        // marshalling provided by the derived classes
        public abstract TMarshalled MarshalFrom(TUnmarshalled input);
        public abstract TUnmarshalled MarshalTo(TMarshalled input);
        public abstract void MarshalRelease(TUnmarshalled input);

        // UnmarshalledValue: dereference the Address, without marshalling
        //(for contiguous arrays this returns the first element of the array)
        public TUnmarshalled UnmarshalledValue
        {
            get => GetArrayElementUnmarshalled(0);
            set => SetArrayElementUnmarshalled(0, value);
        }

        // Value: dereference the Address, with marshalling 
        // (for contiguous arrays this returns the first element of the array)
        public TMarshalled Value
        {
            get => GetArrayElement(0);
            set => SetArrayElement(0, value);
        }

        // the individual size of each element
        public int ElementSize => Marshal.SizeOf(typeof(TUnmarshalled));

        // array accessors, unmarshalled
        public TUnmarshalled GetArrayElementUnmarshalled(int elementIndex)
        {
            var elementAddress = _address + (ElementSize * elementIndex);
            return StructHelper.ReadStructureUnsafe<TUnmarshalled>(elementAddress);
        }
        public void SetArrayElementUnmarshalled(int elementIndex, TUnmarshalled value)
        {
            var elementAddress = _address + (ElementSize * elementIndex);
            Marshal.StructureToPtr<TUnmarshalled>(value, elementAddress, false);
        }

        // array accessors, marshalled
        public TMarshalled GetArrayElement(int elementIndex)
            => MarshalFrom(GetArrayElementUnmarshalled(elementIndex));

        public void SetArrayElement(int elementIndex, TMarshalled value)
            => SetArrayElementUnmarshalled(elementIndex, MarshalTo(value));

        // GetArray: grab a full copy of the array content as a c# array, including marshalling of all elements
        public TMarshalled[] GetArray(int maxCopyElements = 0)
        {
            if (maxCopyElements > _elementCount) throw new InvalidOperationException();
            if (maxCopyElements == 0) maxCopyElements = _elementCount;

            TMarshalled[] retVal = new TMarshalled[maxCopyElements];
            var index = 0;
            while (index < maxCopyElements)
            {
                retVal[index] = GetArrayElement(index);
                index++;
            }
            return retVal;
        }

        // CopyArrayTo: copy the internal contiguous array to a c# array, including marshalling of all elements
        // this assumes input array is of the correct size
        public void CopyArrayTo(TMarshalled[] copyTo)
        {
            var index = 0;
            while (index < _elementCount)
            {
                copyTo[index] = GetArrayElement(index);
                index++;
            }
        }

        public IAddressableVariableBase(int contiguousArrayElementCount = 1, IntPtr alreadyAllocatedMem = default)
        {
            _elementCount = contiguousArrayElementCount;
            var sizeOf = ElementSize * _elementCount;
            _ownedMemory = alreadyAllocatedMem == IntPtr.Zero;
            if (_ownedMemory)
            {
                _address = Marshal.AllocHGlobal(sizeOf);
                Marshal.Copy(new byte[sizeOf], 0, _address, sizeOf); // nullify the data
            }
            else
            {
                _address = alreadyAllocatedMem;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }

            // call the derived MarshallRelease for each element (e.g. Marshal.FreeBSTR for strings etc)
            var index = 0;
            while (index < _elementCount)
            {
                MarshalRelease(GetArrayElementUnmarshalled(index++));
            }

            if (_ownedMemory && (_address != IntPtr.Zero)) Marshal.FreeHGlobal(_address);
            _isDisposed = true;
        }
    }

    // AddressableVariableSimple is ideal for basic types, like int, short, that require no marshalling or special handling
    public class AddressableVariableSimple<TBasicType> : IAddressableVariableBase<TBasicType, TBasicType>
    {
        public AddressableVariableSimple(int contiguousArrayElementCount = 1,
                                    IntPtr alreadyAllocatedMem = default)
            : base(contiguousArrayElementCount, alreadyAllocatedMem) { }

        public override TBasicType MarshalFrom(TBasicType input) { return input; }  // no marshalling for basic types
        public override TBasicType MarshalTo(TBasicType input) { return input; }    // no marshalling for basic types
        public override void MarshalRelease(TBasicType input) { }                   // no cleanup for basic types
    }

    // this one is for marshalling string <-> BSTR
    public class AddressableVariableBSTR : IAddressableVariableBase<IntPtr, string>
    {
        public AddressableVariableBSTR(int contigiousArrayElementCount)
            : base(contigiousArrayElementCount) { }

        public override string MarshalFrom(IntPtr input)
            => (input != IntPtr.Zero) ? Marshal.PtrToStringBSTR(input) : null;
        public override IntPtr MarshalTo(string input)
            => (input != null) ? Marshal.StringToBSTR(input) : IntPtr.Zero;
        public override void MarshalRelease(IntPtr input)
        {
            if (input != IntPtr.Zero) Marshal.FreeBSTR(input);
        }
    }

    // this one is for marshalling com-interface-ptr <-> object
    public class AddressableVariableObject<T> : IAddressableVariableBase<IntPtr, T>
    {
        public AddressableVariableObject(int contigiousArrayElementCount)
            : base(contigiousArrayElementCount) { }

        public override T MarshalFrom(IntPtr input)
            => (input != IntPtr.Zero) ? (T)Marshal.GetObjectForIUnknown(input) : default(T);

        public override IntPtr MarshalTo(T input)
            => (input != null) ? Marshal.GetIUnknownForObject(input) : IntPtr.Zero;

        public override void MarshalRelease(IntPtr input)
        {
            if (input != IntPtr.Zero) Marshal.Release(input);
        }
    }

    // A pointer version, particularly useful for creating out-only pointers, 
    // making the content easily deferencable once set on the unmanaged side
    public class AddressableVariablePtr<T> : IAddressableVariableBase<IntPtr, AddressableVariableSimple<T>>
    {
        public override AddressableVariableSimple<T> MarshalFrom(IntPtr input)
            => new AddressableVariableSimple<T>(alreadyAllocatedMem: UnmarshalledValue);
        public override IntPtr MarshalTo(AddressableVariableSimple<T> input)
            => UnmarshalledValue = input._address;
        public override void MarshalRelease(IntPtr input) { }
    }

    // helpers for creating AddressableVariables
    static class AddressableVariables
    {
        public static AddressableVariableSimple<T> Create<T>(int elementCount = 1)
            => new AddressableVariableSimple<T>(elementCount);

        // CreateOutPtr has an additional layer of indirection
        public static AddressableVariablePtr<T> CreatePtrTo<T>()
            => new AddressableVariablePtr<T>();

        public static AddressableVariableBSTR CreateBSTR(int elementCount = 1)
            => new AddressableVariableBSTR(elementCount);

        public static AddressableVariableObject<T> CreateObjectPtr<T>(int elementCount = 1)
            => new AddressableVariableObject<T>(elementCount);
    }
}
