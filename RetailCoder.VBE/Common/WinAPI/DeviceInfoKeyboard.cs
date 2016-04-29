namespace Rubberduck.Common.WinAPI
{
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
}
