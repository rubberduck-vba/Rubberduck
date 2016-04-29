using System;

namespace Rubberduck.Common.WinAPI
{
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
}
