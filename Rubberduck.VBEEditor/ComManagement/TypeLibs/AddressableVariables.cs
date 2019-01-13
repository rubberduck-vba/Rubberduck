using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// AddressableVariables aid creating and handling unmanaged data
    /// </summary>
    /// <remarks>
    /// AddressableVariables are created in unmanaged memory space, designed to aid creating addressable
    /// content when passing to/from interop code (as IntPtr addresses)
    /// 
    /// Memory is allocated on the heap, and alignment of the data is on 8-byte boundaries as is standard 
    /// for the Windows heap allocators, and this is ample for our use cases.
    /// 
    /// IAddressableVariableBase<T> can  represent a single element, or a contiguous array,
    /// and allows the derived classes to implement marshalling for the elements (e.g. string<->BSTR etc)
    /// </remarks>
    public abstract class IAddressableVariableBase<TUnmarshalled, TMarshalled> : IDisposable
    {
        public IntPtr Address { get; private set; }
        public int ElementCount { get; private set; }       // 1 for singular elements
        private bool _ownedMemory;      // true if WE allocated the memory, and it hasn't been extracted from us (default)

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
            var elementAddress = Address + (ElementSize * elementIndex);
            return StructHelper.ReadStructureUnsafe<TUnmarshalled>(elementAddress);
        }
        public void SetArrayElementUnmarshalled(int elementIndex, TUnmarshalled value)
        {
            var elementAddress = Address + (ElementSize * elementIndex);
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
            if (maxCopyElements > ElementCount) throw new InvalidOperationException();
            if (maxCopyElements == 0) maxCopyElements = ElementCount;

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
            while (index < ElementCount)
            {
                copyTo[index] = GetArrayElement(index);
                index++;
            }
        }

        public IAddressableVariableBase(int contiguousArrayElementCount = 1, IntPtr alreadyAllocatedMem = default)
        {
            ElementCount = contiguousArrayElementCount;
            var sizeOf = ElementSize * ElementCount;
            _ownedMemory = alreadyAllocatedMem == IntPtr.Zero;
            if (_ownedMemory)
            {
                Address = Marshal.AllocHGlobal(sizeOf);
                Marshal.Copy(new byte[sizeOf], 0, Address, sizeOf); // nullify the data
            }
            else
            {
                Address = alreadyAllocatedMem;
            }
        }

        // Extract the address and ensure no cleanup
        public IntPtr Extract()
        {
            _ownedMemory = false;
            return Address;
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

            if (_ownedMemory)
            {
                // call the derived MarshallRelease for each element (e.g. Marshal.FreeBSTR for strings etc)
                var index = 0;
                while (index < ElementCount)
                {
                    MarshalRelease(GetArrayElementUnmarshalled(index++));
                }

                if (Address != IntPtr.Zero) Marshal.FreeHGlobal(Address);
            }
            _isDisposed = true;
        }
    }

    /// <summary>
    /// AddressableVariableSimple is ideal for basic types, like int, short, that require no marshalling or special handling
    /// </summary>
    public class AddressableVariableSimple<TBasicType> : IAddressableVariableBase<TBasicType, TBasicType>
    {
        public AddressableVariableSimple(int contiguousArrayElementCount = 1,
                                    IntPtr alreadyAllocatedMem = default)
            : base(contiguousArrayElementCount, alreadyAllocatedMem) { }

        public override TBasicType MarshalFrom(TBasicType input) { return input; }  // no marshalling for basic types
        public override TBasicType MarshalTo(TBasicType input) { return input; }    // no marshalling for basic types
        public override void MarshalRelease(TBasicType input) { }                   // no cleanup for basic types
    }

    /// <summary>
    /// AddressableVariableBSTR handles marshalling between string and BSTR
    /// </summary>
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

    /// <summary>
    /// AddressableVariableObject handles marshalling between COM interface pointers, and object
    /// </summary>
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

    /// <summary>
    /// AddressableVariablePtr is ideal for creating out-only pointers, making the content easily 
    /// deferencable once set on the unmanaged side. Designed for simple types with no content marshalling.
    /// </summary>
    public class AddressableVariablePtr<T> : IAddressableVariableBase<IntPtr, AddressableVariableSimple<T>>
    {
        public override AddressableVariableSimple<T> MarshalFrom(IntPtr input)
            => new AddressableVariableSimple<T>(alreadyAllocatedMem: UnmarshalledValue);
        public override IntPtr MarshalTo(AddressableVariableSimple<T> input)
            => UnmarshalledValue = input.Address;
        public override void MarshalRelease(IntPtr input) { }
    }

    /// <summary>
    /// AddressableVariables exposes helpers for creating the common derived classes 
    /// </summary>
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
