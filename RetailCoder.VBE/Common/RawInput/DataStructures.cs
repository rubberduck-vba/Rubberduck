using System;
using System.Runtime.InteropServices;

namespace RawInput_dll
{
    // ReSharper disable FieldCanBeMadeReadOnly.Global

    [StructLayout(LayoutKind.Explicit)]
    public struct DeviceInfo
    {
        [FieldOffset(0)]
        public int Size;
        [FieldOffset(4)]
        public int Type;

        [FieldOffset(8)]
        public DeviceInfoMouse MouseInfo;
        [FieldOffset(8)]
        public DeviceInfoKeyboard KeyboardInfo;
        [FieldOffset(8)]
        public DeviceInfoHid HIDInfo;

        public override string ToString()
        {
            return string.Format("DeviceInfo\n Size: {0}\n Type: {1}\n", Size, Type);
        }
    }

    public struct DeviceInfoMouse
    {
        // ReSharper disable MemberCanBePrivate.Global
        public uint Id;                         // Identifier of the mouse device
        public uint NumberOfButtons;            // Number of buttons for the mouse
        public uint SampleRate;                 // Number of data points per second.
        public bool HasHorizontalWheel;         // True is mouse has wheel for horizontal scrolling else false.
        // ReSharper restore MemberCanBePrivate.Global
        public override string ToString()
        {
            return string.Format("MouseInfo\n Id: {0}\n NumberOfButtons: {1}\n SampleRate: {2}\n HorizontalWheel: {3}\n", Id, NumberOfButtons, SampleRate, HasHorizontalWheel);
        }
    }

    public struct DeviceInfoKeyboard
    {
        public uint Type;                       // Type of the keyboard
        public uint SubType;                    // Subtype of the keyboard
        public uint KeyboardMode;               // The scan code mode
        public uint NumberOfFunctionKeys;       // Number of function keys on the keyboard
        public uint NumberOfIndicators;         // Number of LED indicators on the keyboard
        public uint NumberOfKeysTotal;          // Total number of keys on the keyboard

        public override string ToString()
        {
            return string.Format("DeviceInfoKeyboard\n Type: {0}\n SubType: {1}\n KeyboardMode: {2}\n NumberOfFunctionKeys: {3}\n NumberOfIndicators {4}\n NumberOfKeysTotal: {5}\n", 
                                                             Type, 
                                                             SubType, 
                                                             KeyboardMode,
                                                             NumberOfFunctionKeys, 
                                                             NumberOfIndicators, 
                                                             NumberOfKeysTotal);
        }
    }

    public struct DeviceInfoHid
    {
        public uint VendorID;       // Vendor identifier for the HID
        public uint ProductID;      // Product identifier for the HID
        public uint VersionNumber;  // Version number for the device
        public ushort UsagePage;    // Top-level collection Usage page for the device
        public ushort Usage;        // Top-level collection Usage for the device

        public override string ToString()
        {
            return string.Format("HidInfo\n VendorID: {0}\n ProductID: {1}\n VersionNumber: {2}\n UsagePage: {3}\n Usage: {4}\n", VendorID, ProductID, VersionNumber, UsagePage, Usage);
        }
    }

    struct BroadcastDeviceInterface
    {
        // ReSharper disable NotAccessedField.Global
        // ReSharper disable UnusedField.Compiler
        public Int32 DbccSize;
        public BroadcastDeviceType BroadcastDeviceType;
        public Int32 DbccReserved;
        public Guid DbccClassguid;
        public char DbccName;
        // ReSharper restore NotAccessedField.Global
        // ReSharper restore UnusedField.Compiler
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Rawinputdevicelist
    {
        public IntPtr hDevice;
        public uint dwType;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct RawData
    {
        [FieldOffset(0)]
        internal Rawmouse mouse;
        [FieldOffset(0)]
        internal Rawkeyboard keyboard;
        [FieldOffset(0)]
        internal Rawhid hid;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct  InputData
    {
        public Rawinputheader header;           // 64 bit header size: 24  32 bit the header size: 16
        public RawData data;                    // Creating the rest in a struct allows the header size to align correctly for 32/64 bit
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct Rawinputheader
    {
        public uint dwType;                     // Type of raw input (RIM_TYPEHID 2, RIM_TYPEKEYBOARD 1, RIM_TYPEMOUSE 0)
        public uint dwSize;                     // Size in bytes of the entire input packet of data. This includes RAWINPUT plus possible extra input reports in the RAWHID variable length array. 
        public IntPtr hDevice;                  // A handle to the device generating the raw input data. 
        public IntPtr wParam;                   // RIM_INPUT 0 if input occurred while application was in the foreground else RIM_INPUTSINK 1 if it was not.

        public override string ToString()
        {
            return string.Format("RawInputHeader\n dwType : {0}\n dwSize : {1}\n hDevice : {2}\n wParam : {3}", dwType, dwSize, hDevice, wParam);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Rawhid
    {
        public uint dwSizHid;
        public uint dwCount;
        public byte bRawData;

        public override string ToString()
        {
            return string.Format("Rawhib\n dwSizeHid : {0}\n dwCount : {1}\n bRawData : {2}\n", dwSizHid, dwCount, bRawData);
        }
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct Rawmouse
    {
        [FieldOffset(0)]
        public ushort usFlags;
        [FieldOffset(4)]
        public uint ulButtons;
        [FieldOffset(4)]
        public ushort usButtonFlags;
        [FieldOffset(6)]
        public ushort usButtonData;
        [FieldOffset(8)]
        public uint ulRawButtons;
        [FieldOffset(12)]
        public int lLastX;
        [FieldOffset(16)]
        public int lLastY;
        [FieldOffset(20)]
        public uint ulExtraInformation;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Rawkeyboard
    {
        public ushort Makecode;                 // Scan code from the key depression
        public ushort Flags;                    // One or more of RI_KEY_MAKE, RI_KEY_BREAK, RI_KEY_E0, RI_KEY_E1
        private readonly ushort Reserved;       // Always 0    
        public ushort VKey;                     // Virtual Key Code
        public uint Message;                    // Corresponding Windows message for exmaple (WM_KEYDOWN, WM_SYASKEYDOWN etc)
        public uint ExtraInformation;           // The device-specific addition information for the event (seems to always be zero for keyboards)

        public override string ToString()
        {
            return string.Format("Rawkeyboard\n Makecode: {0}\n Makecode(hex) : {0:X}\n Flags: {1}\n Reserved: {2}\n VKeyName: {3}\n Message: {4}\n ExtraInformation {5}\n", 
                                                Makecode, Flags, Reserved, VKey, Message, ExtraInformation);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct RawInputDevice
    {
        internal HidUsagePage UsagePage;
        internal HidUsage Usage;
        internal RawInputDeviceFlags Flags;
        internal IntPtr Target;

        public override string ToString()
        {
            return string.Format("{0}/{1}, flags: {2}, target: {3}", UsagePage, Usage, Flags, Target);
        }
    }
}
