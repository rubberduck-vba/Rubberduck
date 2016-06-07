namespace Rubberduck.Common.WinAPI
{
    public struct DeviceInfoMouse
    {
        public uint Id;                         // Identifier of the mouse device
        public uint NumberOfButtons;            // Number of buttons for the mouse
        public uint SampleRate;                 // Number of data points per second.
        public bool HasHorizontalWheel;         // True is mouse has wheel for horizontal scrolling else false.
        public override string ToString()
        {
            return string.Format("MouseInfo\n Id: {0}\n NumberOfButtons: {1}\n SampleRate: {2}\n HorizontalWheel: {3}\n", Id, NumberOfButtons, SampleRate, HasHorizontalWheel);
        }
    }
}
