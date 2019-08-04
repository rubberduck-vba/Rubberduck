using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged
{
    /// <summary>
    /// Windows API constants used by VirtualQuery, see https://msdn.microsoft.com/en-us/library/windows/desktop/aa366786.aspx
    /// </summary>
    internal enum ALLOCATION_PROTECTION : uint
    {
        PAGE_EXECUTE = 0x00000010,
        PAGE_EXECUTE_READ = 0x00000020,
        PAGE_EXECUTE_READWRITE = 0x00000040,
        PAGE_EXECUTE_WRITECOPY = 0x00000080,
        PAGE_NOACCESS = 0x00000001,
        PAGE_READONLY = 0x00000002,
        PAGE_READWRITE = 0x00000004,
        PAGE_WRITECOPY = 0x00000008,
        PAGE_GUARD = 0x00000100,
        PAGE_NOCACHE = 0x00000200,
        PAGE_WRITECOMBINE = 0x00000400
    }

    /// <summary>
    /// Windows API structure used by VirtualQuery, see https://msdn.microsoft.com/en-us/library/windows/desktop/aa366775.aspx
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct MEMORY_BASIC_INFORMATION
    {
        public IntPtr BaseAddress;
        public IntPtr AllocationBase;
        public uint AllocationProtect;
        public IntPtr RegionSize;
        public uint State;
        public ALLOCATION_PROTECTION Protect;
        public uint Type;
    }

    /// <summary>
    /// Exposes some special routines for dealing with unmanaged memory
    /// </summary>
    internal static class UnmanagedMemoryHelper
    {
        /// <summary>
        /// Windows API call used for memory range validation
        /// </summary>
        [DllImport("kernel32.dll")]
        public static extern IntPtr VirtualQuery(IntPtr lpAddress, out MEMORY_BASIC_INFORMATION lpBuffer, IntPtr dwLength);

        /// <summary>
        /// Do our best to validate that the input memory address is actually a COM object
        /// </summary>
        /// <param name="comObjectPtr">the input memory address to check</param>
        /// <returns>false means definitely not a valid COM object.  true means _probably_ a valid COM object</returns>
        public static bool ValidateComObject(IntPtr comObjectPtr)
        {
            // Is it a valid memory address, with at least one accessible vTable ptr
            if (!IsValidMemoryRange(comObjectPtr, IntPtr.Size)) return false;

            var vTablePtr = RdMarshal.ReadIntPtr(comObjectPtr);

            // And for a COM object, we need a valid vtable, with at least 3 vTable entries (for IUnknown)
            if (!IsValidMemoryRange(vTablePtr, IntPtr.Size * 3)) return false;

            var firstvTableEntry = RdMarshal.ReadIntPtr(vTablePtr);

            // And lets check the first vTable entry actually points to EXECUTABLE memory
            // (we could check all 3 initial IUnknown entries, but we want to be reasonably  
            // efficient and we can never 100% guarantee our result anyway.)
            if (IsValidMemoryRange(firstvTableEntry, 1, checkIsExecutable: true))
            {
                // As best as we can tell, it looks to be a valid COM object
                return true;
            }
            else
            {
                // One of the validation checks failed.  The COM object is definitely not a valid COM object.
                return false;
            }
        }

        /// <summary>
        /// Validate a memory address range
        /// </summary>
        /// <param name="memOffset">the input memory address to check</param>
        /// <param name="size">the minimum size of data we are expecting to be available next to memOffset</param>
        /// <param name="checkIsExecutable">optionally check if the memory address points to EXECUTABLE memory</param>
        public static bool IsValidMemoryRange(IntPtr memOffset, int size, bool checkIsExecutable = false)
        {
            if (memOffset == IntPtr.Zero) return false;

            var memInfo = new MEMORY_BASIC_INFORMATION();
            var sizeOfMemInfo = new IntPtr(RdMarshal.SizeOf(memInfo));

            // most of the time, a bad pointer will fail here
            if (VirtualQuery(memOffset, out memInfo, sizeOfMemInfo) != sizeOfMemInfo)
            {
                return false;
            }

            // check the memory area is not a guard page, or otherwise inaccessible
            if ((memInfo.Protect.HasFlag(ALLOCATION_PROTECTION.PAGE_NOACCESS)) ||
                (memInfo.Protect.HasFlag(ALLOCATION_PROTECTION.PAGE_GUARD)))
            {
                return false;
            }

            // We've confirmed the base memory address is valid, and is accessible.
            // Finally just check the full address RANGE is also valid (i.e. the end point of the structure we're reading)
            var validMemAddressEnd = memInfo.BaseAddress.ToInt64() + memInfo.RegionSize.ToInt64();
            var endOfStructPtr = memOffset.ToInt64() + size;
            if (endOfStructPtr > validMemAddressEnd) return false;

            if (checkIsExecutable)
            {
                // We've been asked to check if the memory address is marked as containing executable code
                return memInfo.Protect.HasFlag(ALLOCATION_PROTECTION.PAGE_EXECUTE) ||
                        memInfo.Protect.HasFlag(ALLOCATION_PROTECTION.PAGE_EXECUTE_READ) ||
                        memInfo.Protect.HasFlag(ALLOCATION_PROTECTION.PAGE_EXECUTE_READWRITE) ||
                        memInfo.Protect.HasFlag(ALLOCATION_PROTECTION.PAGE_EXECUTE_WRITECOPY);
            }
            return true;
        }
    }

    /// <summary>
    /// Encapsulates reading unmanaged memory into managed structures
    /// </summary>
    internal static class StructHelper
    {
        /// <summary>
        /// Takes a COM object, and reads the unmanaged memory given by its pointer, allowing us to read internal fields
        /// </summary>
        /// <typeparam name="T">the type of structure to return</typeparam>
        /// <param name="comObj">the COM object</param>
        /// <returns>the requested structure T</returns>
        public static T ReadComObjectStructure<T>(object comObj)
        {
            // Reads a COM object as a structure to copy its internal fields
            if (!RdMarshal.IsComObject(comObj))
            {
                throw new ArgumentException("Expected a COM object");
            }

            var referencesPtr = RdMarshal.GetIUnknownForObjectInContext(comObj);
            if (referencesPtr == IntPtr.Zero)
            {
                throw new InvalidOperationException("Cannot access the TypeLib API from this thread.  TypeLib API must be accessed from the main thread.");
            }
            var retVal = ReadStructureSafe<T>(referencesPtr);
            RdMarshal.Release(referencesPtr);
            return retVal;
        }

        /// <summary>
        /// Takes an unmanaged memory address and reads the unmanaged memory given by its pointer
        /// </summary>
        /// <typeparam name="T">the type of structure to return</typeparam>
        /// <param name="memAddress">the unmanaged memory address to read</param>
        /// <returns>the requested structure T</returns>
        /// <remarks>use this over ReadStructureSafe for efficiency when there is no doubt about the validity of the pointed to data</remarks>
        public static T ReadStructureUnsafe<T>(IntPtr memAddress)
        {
            // We catch the most basic mistake of passing a null pointer here as it virtually costs nothing to check, 
            // but no other checks are made as to the validity of the pointer. 
            if (memAddress == IntPtr.Zero) throw new ArgumentException("Unexpected null pointer.");
            return (T)RdMarshal.PtrToStructure(memAddress, typeof(T));
        }


        /// <summary>
        /// Takes an unmanaged memory address and reads the unmanaged memory given by its pointer, 
        /// with memory address validation for added protection
        /// </summary>
        /// <typeparam name="T">the type of structure to return</typeparam>
        /// <param name="memAddress">the unamanaged memory address to read</param>
        /// <returns>the requested structure T</returns>
        public static T ReadStructureSafe<T>(IntPtr memAddress)
        {
            if (UnmanagedMemoryHelper.IsValidMemoryRange(memAddress, RdMarshal.SizeOf(typeof(T))))
            {
                return (T)RdMarshal.PtrToStructure(memAddress, typeof(T));
            }

            throw new ArgumentException("Bad data pointer - unable to read structure data.");
        }
    }
}
