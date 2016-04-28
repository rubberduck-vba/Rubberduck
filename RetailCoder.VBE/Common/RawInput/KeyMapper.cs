using System.Globalization;
using System.Windows.Forms;

namespace RawInput_dll
{
    public static class KeyMapper
    {
        // I prefer to have control over the key mapping
        // This mapping could be loading from file to allow mapping changes without a recompile
        public  static string GetKeyName(int value)
        {
            switch (value)
            {
                case 0x41: return "A";
                case 0x6b: return "Add";
                case 0x40000: return "Alt";
                case 0x5d: return "Apps";
                case 0xf6: return "Attn";
                case 0x42: return "B";
                case 8: return "Back";
                case 0xa6: return "BrowserBack";
                case 0xab: return "BrowserFavorites";
                case 0xa7: return "BrowserForward";
                case 0xac: return "BrowserHome";
                case 0xa8: return "BrowserRefresh";
                case 170: return "BrowserSearch";
                case 0xa9: return "BrowserStop";
                case 0x43: return "C";
                case 3: return "Cancel";
                case 20: return "Capital";
                //case 20:      return "CapsLock";
                case 12: return "Clear";
                case 0x20000: return "Control";
                case 0x11: return "ControlKey";
                case 0xf7: return "Crsel";
                case 0x44: return "D";
                case 0x30: return "D0";
                case 0x31: return "D1";
                case 50: return "D2";
                case 0x33: return "D3";
                case 0x34: return "D4";
                case 0x35: return "D5";
                case 0x36: return "D6";
                case 0x37: return "D7";
                case 0x38: return "D8";
                case 0x39: return "D9";
                case 110: return "Decimal";
                case 0x2e: return "Delete";
                case 0x6f: return "Divide";
                case 40: return "Down";
                case 0x45: return "E";
                case 0x23: return "End";
                case 13: return "Enter";
                case 0xf9: return "EraseEof";
                case 0x1b: return "Escape";
                case 0x2b: return "Execute";
                case 0xf8: return "Exsel";
                case 70: return "F";
                case 0x70: return "F1";
                case 0x79: return "F10";
                case 0x7a: return "F11";
                case 0x7b: return "F12";
                case 0x7c: return "F13";
                case 0x7d: return "F14";
                case 0x7e: return "F15";
                case 0x7f: return "F16";
                case 0x80: return "F17";
                case 0x81: return "F18";
                case 130: return "F19";
                case 0x71: return "F2";
                case 0x83: return "F20";
                case 0x84: return "F21";
                case 0x85: return "F22";
                case 0x86: return "F23";
                case 0x87: return "F24";
                case 0x72: return "F3";
                case 0x73: return "F4";
                case 0x74: return "F5";
                case 0x75: return "F6";
                case 0x76: return "F7";
                case 0x77: return "F8";
                case 120: return "F9";
                case 0x18: return "FinalMode";
                case 0x47: return "G";
                case 0x48: return "H";
                case 0x15: return "HanguelMode";
                //case 0x15:    return "HangulMode";
                case 0x19: return "HanjaMode";
                case 0x2f: return "Help";
                case 0x24: return "Home";
                case 0x49: return "I";
                case 30: return "IMEAceept";
                case 0x1c: return "IMEConvert";
                case 0x1f: return "IMEModeChange";
                case 0x1d: return "IMENonconvert";
                case 0x2d: return "Insert";
                case 0x4a: return "J";
                case 0x17: return "JunjaMode";
                case 0x4b: return "K";
                //case 0x15:    return "KanaMode";
                //case 0x19:    return "KanjiMode";
                case 0xffff: return "KeyCode";
                case 0x4c: return "L";
                case 0xb6: return "LaunchApplication1";
                case 0xb7: return "LaunchApplication2";
                case 180: return "LaunchMail";
                case 1: return "LButton";
                case 0xa2: return "LControl";
                case 0x25: return "Left";
                case 10: return "LineFeed";
                case 0xa4: return "LMenu";
                case 160: return "LShift";
                case 0x5b: return "LWin";
                case 0x4d: return "M";
                case 4: return "MButton";
                case 0xb0: return "MediaNextTrack";
                case 0xb3: return "MediaPlayPause";
                case 0xb1: return "MediaPreviousTrack";
                case 0xb2: return "MediaStop";
                case 0x12: return "Menu";
                // case 65536:  return "Modifiers";
                case 0x6a: return "Multiply";
                case 0x4e: return "N";
                case 0x22: return "Next";
                case 0xfc: return "NoName";
                case 0: return "None";
                case 0x90: return "NumLock";
                case 0x60: return "NumPad0";
                case 0x61: return "NumPad1";
                case 0x62: return "NumPad2";
                case 0x63: return "NumPad3";
                case 100: return "NumPad4";
                case 0x65: return "NumPad5";
                case 0x66: return "NumPad6";
                case 0x67: return "NumPad7";
                case 0x68: return "NumPad8";
                case 0x69: return "NumPad9";
                case 0x4f: return "O";
                case 0xdf: return "Oem8";
                case 0xe2: return "OemBackslash";
                case 0xfe: return "OemClear";
                case 0xdd: return "OemCloseBrackets";
                case 0xbc: return "OemComma";
                case 0xbd: return "OemMinus";
                case 0xdb: return "OemOpenBrackets";
                case 190: return "OemPeriod";
                case 220: return "OemPipe";
                case 0xbb: return "Oemplus";
                case 0xbf: return "OemQuestion";
                case 0xde: return "OemQuotes";
                case 0xba: return "OemSemicolon";
                case 0xc0: return "Oemtilde";
                case 80: return "P";
                case 0xfd: return "Pa1";
                // case 0x22:   return "PageDown";
                // case 0x21:   return "PageUp";
                case 0x13: return "Pause";
                case 250: return "Play";
                case 0x2a: return "Print";
                case 0x2c: return "PrintScreen";
                case 0x21: return "Prior";
                case 0xe5: return "ProcessKey";
                case 0x51: return "Q";
                case 0x52: return "R";
                case 2: return "RButton";
                case 0xa3: return "RControl";
                //case 13:      return "Return";
                case 0x27: return "Right";
                case 0xa5: return "RMenu";
                case 0xa1: return "RShift";
                case 0x5c: return "RWin";
                case 0x53: return "S";
                case 0x91: return "Scroll";
                case 0x29: return "Select";
                case 0xb5: return "SelectMedia";
                case 0x6c: return "Separator";
                case 0x10000: return "Shift";
                case 0x10: return "ShiftKey";
                //case 0x2c:    return "Snapshot";
                case 0x20: return "Space";
                case 0x6d: return "Subtract";
                case 0x54: return "T";
                case 9: return "Tab";
                case 0x55: return "U";
                case 0x26: return "Up";
                case 0x56: return "V";
                case 0xae: return "VolumeDown";
                case 0xad: return "VolumeMute";
                case 0xaf: return "VolumeUp";
                case 0x57: return "W";
                case 0x58: return "X";
                case 5: return "XButton1";
                case 6: return "XButton2";
                case 0x59: return "Y";
                case 90: return "Z";
                case 0xfb: return "Zoom";
            }

            return value.ToString(CultureInfo.InvariantCulture).ToUpper();
        }

        // If you prefer the virtualkey converted into a Microsoft virtualkey code use this
        public static string GetMicrosoftKeyName(int virtualKey)
        {
            return new KeysConverter().ConvertToString(virtualKey);
        }
    }
}
