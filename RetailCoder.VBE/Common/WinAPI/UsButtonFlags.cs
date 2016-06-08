using System;

namespace Rubberduck.Common.WinAPI
{
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
