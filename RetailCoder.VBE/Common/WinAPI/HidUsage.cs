using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Common.WinAPI
{
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
}
