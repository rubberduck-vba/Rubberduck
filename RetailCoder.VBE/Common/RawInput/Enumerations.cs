using System;

namespace RawInput_dll
{
    // ReSharper disable UnusedMember.Global
    // ReSharper disable InconsistentNaming

    public enum DataCommand : uint
    {
        RID_HEADER = 0x10000005, // Get the header information from the RAWINPUT structure.
        RID_INPUT = 0x10000003   // Get the raw data from the RAWINPUT structure.
    }

    public static class DeviceType
    {
        public const int RimTypeMouse = 0;
        public const int RimTypekeyboard = 1;
        public const int RimTypeHid = 2;
    }
    
    internal enum RawInputDeviceInfo : uint
    {
        RIDI_DEVICENAME = 0x20000007,
        RIDI_DEVICEINFO = 0x2000000b,
        PREPARSEDDATA = 0x20000005
    }

    enum BroadcastDeviceType
    {
        DBT_DEVTYP_OEM = 0,
        DBT_DEVTYP_DEVNODE = 1,
        DBT_DEVTYP_VOLUME = 2,
        DBT_DEVTYP_PORT = 3,
        DBT_DEVTYP_NET = 4,
        DBT_DEVTYP_DEVICEINTERFACE = 5,
        DBT_DEVTYP_HANDLE = 6,
    }

    enum DeviceNotification
    {
        DEVICE_NOTIFY_WINDOW_HANDLE = 0x00000000,           // The hRecipient parameter is a window handle
        DEVICE_NOTIFY_SERVICE_HANDLE = 0x00000001,          // The hRecipient parameter is a service status handle
        DEVICE_NOTIFY_ALL_INTERFACE_CLASSES = 0x00000004    // Notifies the recipient of device interface events for all device interface classes. (The dbcc_classguid member is ignored.)
                                                            // This value can be used only if the dbch_devicetype member is DBT_DEVTYP_DEVICEINTERFACE.
    }

    [Flags]
    internal enum RawInputDeviceFlags
    {
        NONE = 0,                   // No flags
        REMOVE = 0x00000001,        // Removes the top level collection from the inclusion list. This tells the operating system to stop reading from a device which matches the top level collection. 
        EXCLUDE = 0x00000010,       // Specifies the top level collections to exclude when reading a complete usage page. This flag only affects a TLC whose usage page is already specified with PageOnly.
        PAGEONLY = 0x00000020,      // Specifies all devices whose top level collection is from the specified UsagePage. Note that Usage must be zero. To exclude a particular top level collection, use Exclude.
        NOLEGACY = 0x00000030,      // Prevents any devices specified by UsagePage or Usage from generating legacy messages. This is only for the mouse and keyboard.
        INPUTSINK = 0x00000100,     // Enables the caller to receive the input even when the caller is not in the foreground. Note that WindowHandle must be specified.
        CAPTUREMOUSE = 0x00000200,  // Mouse button click does not activate the other window.
        NOHOTKEYS = 0x00000200,     // Application-defined keyboard device hotkeys are not handled. However, the system hotkeys; for example, ALT+TAB and CTRL+ALT+DEL, are still handled. By default, all keyboard hotkeys are handled. NoHotKeys can be specified even if NoLegacy is not specified and WindowHandle is NULL.
        APPKEYS = 0x00000400,       // Application keys are handled.  NoLegacy must be specified.  Keyboard only.
        
        // Enables the caller to receive input in the background only if the foreground application does not process it. 
        // In other words, if the foreground application is not registered for raw input, then the background application that is registered will receive the input.
        EXINPUTSINK = 0x00001000,
        DEVNOTIFY = 0x00002000
    }

    public enum HidUsagePage : ushort
    {
        UNDEFINED = 0x00,   // Unknown usage page
        GENERIC = 0x01,     // Generic desktop controls
        SIMULATION = 0x02,  // Simulation controls
        VR = 0x03,          // Virtual reality controls
        SPORT = 0x04,       // Sports controls
        GAME = 0x05,        // Games controls
        KEYBOARD = 0x07,    // Keyboard controls
    }

    public enum HidUsage : ushort
    {
        Undefined = 0x00,       // Unknown usage
        Pointer = 0x01,         // Pointer
        Mouse = 0x02,           // Mouse
        Joystick = 0x04,        // Joystick
        Gamepad = 0x05,         // Game Pad
        Keyboard = 0x06,        // Keyboard
        Keypad = 0x07,          // Keypad
        SystemControl = 0x80,   // Muilt-axis Controller
        Tablet = 0x80,          // Tablet PC controls
        Consumer = 0x0C,        // Consumer
    }

    [Flags]
    public enum UsButtonFlags : ushort
    {
        None = 0,
        RI_MOUSE_LEFT_BUTTON_DOWN = 1,
        RI_MOUSE_LEFT_BUTTON_UP = 2,
        RI_MOUSE_RIGHT_BUTTON_DOWN = 4,
        RI_MOUSE_RIGHT_BUTTON_UP = 8,
        RI_MOUSE_MIDDLE_BUTTON_DOWN = 16,
        RI_MOUSE_MIDDLE_BUTTON_UP = 32,
        RI_MOUSE_BUTTON_4_DOWN = 64,
        RI_MOUSE_BUTTON_4_UP = 128,
        RI_MOUSE_BUTTON_5_DOWN = 256,
        RI_MOUSE_BUTTON_5_UP = 512,
        RI_MOUSE_WHEEL = 1024
    }
}
