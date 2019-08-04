using System;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged
{
    /// <summary>
    /// AddressableVariables aids in creating and handling unmanaged data
    /// </summary>
    /// <remarks>
    /// AddressableVariables are created in unmanaged memory space, designed to aid creating addressable
    /// content when passing to/from interop code (as IntPtr addresses)
    /// 
    /// Memory is allocated on the heap, and alignment of the data is on 8-byte boundaries as is standard 
    /// for the Windows heap allocators, and this is ample for our use cases.
    /// 
    /// <see cref="AddressableVariableBase{TUnmarshalled,TMarshalled}"/> can  represent a single element, or a contiguous array,
    /// and allows the derived classes to implement marshalling for the elements (e.g. string<->BSTR etc)
    ///
    /// Use static <see cref="AddressableVariables"/> to create objects simply.
    /// </remarks>
    internal abstract class AddressableVariableBase<TUnmarshalled, TMarshalled> : IDisposable
    {
        public IntPtr Address { get; }
        public int ElementCount { get; }    // 1 for singular elements
        private bool _ownedMemory;          // true if WE allocated the memory, and it hasn't been extracted from us (default)

        // marshalling provided by the derived classes
        protected abstract TMarshalled MarshalFrom(TUnmarshalled input);
        protected abstract TUnmarshalled MarshalTo(TMarshalled input);
        protected abstract void MarshalRelease(TUnmarshalled input);

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
        public readonly int ElementSize;

        // array accessors, unmarshalled
        public TUnmarshalled GetArrayElementUnmarshalled(int elementIndex)
        {
            var elementAddress = Address + (ElementSize * elementIndex);
            return StructHelper.ReadStructureUnsafe<TUnmarshalled>(elementAddress);
        }
        public void SetArrayElementUnmarshalled(int elementIndex, TUnmarshalled value)
        {
            var elementAddress = Address + (ElementSize * elementIndex);
            RdMarshal.StructureToPtr<TUnmarshalled>(value, elementAddress, true);
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

        protected AddressableVariableBase(int contiguousArrayElementCount = 1, IntPtr alreadyAllocatedMem = default)
        {
            ElementSize = RdMarshal.SizeOf(typeof(TUnmarshalled));

            ElementCount = contiguousArrayElementCount;
            var sizeOf = ElementSize * ElementCount;

            // Some methods may provide the pointers to us; in this case, we shouldn't 
            // deallocate the pointers since we don't own it. 
            _ownedMemory = alreadyAllocatedMem == IntPtr.Zero;
            if (_ownedMemory)
            {
                Address = RdMarshal.AllocHGlobal(sizeOf);
                RdMarshal.Copy(new byte[sizeOf], 0, Address, sizeOf); // nullify the data
            }
            else
            {
                Address = alreadyAllocatedMem;
            }
        }

        /// <summary>
        /// Extract the address and ensure no cleanup; the caller is now the owner of the
        /// unmanaged variable and must release it itself.
        /// </summary>
        /// <returns><see cref="IntPtr"/> representing the address of the unmanaged variable</returns>
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
                // call the derived MarshalRelease for each element (e.g. Marshal.FreeBSTR for strings etc)
                var index = 0;
                while (index < ElementCount)
                {
                    MarshalRelease(GetArrayElementUnmarshalled(index++));
                }

                if (Address != IntPtr.Zero) RdMarshal.FreeHGlobal(Address);
            }
            _isDisposed = true;
        }
    }

    /// <summary>
    /// AddressableVariableSimple is ideal for basic types, like int, short, that require no marshalling or special handling
    /// </summary>
    internal class AddressableVariableSimple<TBasicType> : AddressableVariableBase<TBasicType, TBasicType>
    {
        public AddressableVariableSimple(int contiguousArrayElementCount = 1,
                                    IntPtr alreadyAllocatedMem = default)
            : base(contiguousArrayElementCount, alreadyAllocatedMem) { }

        protected override TBasicType MarshalFrom(TBasicType input) { return input; }  // no marshalling for basic types
        protected override TBasicType MarshalTo(TBasicType input) { return input; }    // no marshalling for basic types
        protected override void MarshalRelease(TBasicType input) { }                   // no cleanup for basic types
    }

    /// <summary>
    /// AddressableVariableBSTR handles marshalling between string and BSTR
    /// </summary>
    internal class AddressableVariableBSTR : AddressableVariableBase<IntPtr, string>
    {
        public AddressableVariableBSTR(int contigiousArrayElementCount)
            : base(contigiousArrayElementCount) { }

        protected override string MarshalFrom(IntPtr input)
            => (input != IntPtr.Zero) ? RdMarshal.PtrToStringBSTR(input) : null;
        protected override IntPtr MarshalTo(string input)
            => (input != null) ? RdMarshal.StringToBSTR(input) : IntPtr.Zero;
        protected override void MarshalRelease(IntPtr input)
        {
            if (input != IntPtr.Zero) RdMarshal.FreeBSTR(input);
        }
    }

    /// <summary>
    /// AddressableVariableObject handles marshalling between COM interface pointers, and object
    /// </summary>
    internal class AddressableVariableObject<T> : AddressableVariableBase<IntPtr, T>
    {
        public AddressableVariableObject(int contigiousArrayElementCount)
            : base(contigiousArrayElementCount) { }

        protected override T MarshalFrom(IntPtr input)
            => (input != IntPtr.Zero) ? (T)RdMarshal.GetObjectForIUnknown(input) : default(T);

        protected override IntPtr MarshalTo(T input)
            => (input != null) ? RdMarshal.GetIUnknownForObject(input) : IntPtr.Zero;

        protected override void MarshalRelease(IntPtr input)
        {
            if (input != IntPtr.Zero) RdMarshal.Release(input);
        }
    }

    /// <summary>
    /// AddressableVariablePtr is ideal for creating out-only pointers, making the content easily 
    /// deferencable once set on the unmanaged side. Designed for simple types with no content marshalling.
    /// </summary>
    internal class AddressableVariablePtr<T> : AddressableVariableBase<IntPtr, AddressableVariableSimple<T>>
    {
        protected override AddressableVariableSimple<T> MarshalFrom(IntPtr input)
            => new AddressableVariableSimple<T>(alreadyAllocatedMem: UnmarshalledValue);
        protected override IntPtr MarshalTo(AddressableVariableSimple<T> input)
            => UnmarshalledValue = input.Address;
        protected override void MarshalRelease(IntPtr input) { }
    }

    /// <summary>
    /// AddressableVariables exposes helpers for creating the common derived classes 
    /// </summary>
    internal static class AddressableVariables
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
