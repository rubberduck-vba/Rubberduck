using System;

namespace Rubberduck.Common.WinAPI
{
    static public class Win32
    {
        public const int KEYBOARD_OVERRUN_MAKE_CODE = 0xFF;
        
        internal const int VK_SHIFT = 0x10;
        internal const int RI_KEY_MAKE = 0x00;      // Key Down
        internal const int RI_KEY_BREAK = 0x01;     // Key Up
        internal const int RI_KEY_E0 = 0x02;        // Left version of the key
        internal const int RI_KEY_E1 = 0x04;        // Right version of the key. Only seems to be set for the Pause/Break key.        
        internal const int VK_CONTROL = 0x11;
        internal const int VK_MENU = 0x12;
        internal const int VK_ZOOM = 0xFB;
        internal const int VK_LSHIFT = 0xA0;
        internal const int VK_RSHIFT = 0xA1;
        internal const int VK_LCONTROL = 0xA2;
        internal const int VK_RCONTROL = 0xA3;
        internal const int VK_LMENU = 0xA4;
        internal const int VK_RMENU = 0xA5;
        internal const int SC_SHIFT_R = 0x36;

        public static string GetDeviceType(DeviceType device)
        {
            string type;
            switch (device)
            {
                case DeviceType.RIM_TYPE_MOUSE:
                    type = "MOUSE";
                    break;
                case DeviceType.RIM_TYPE_KEYBOARD:
                    type = "KEYBOARD";
                    break;
                case DeviceType.RIM_TYPE_HID:
                    type = "HID";
                    break;
                default:
                    type = "UNKNOWN";
                    break;
            }
            return type;
        }

        public static string GetDeviceDescription(string device)
        {
            string description;
            try
            {
                var key = RegistryAccess.GetDeviceKey(device);
                description = key.GetValue("DeviceDesc").ToString();
                description = description.Substring(description.IndexOf(';') + 1);
            }
            catch (Exception)
            {
                description = "Device is malformed unable to look up in the registry";
            }
            return description;
        }
    }
}
