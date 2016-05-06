using System;

namespace Rubberduck.Common.WinAPI
{
    public interface IRawDevice
    {
        void ProcessRawInput(InputData _rawBuffer);
    }
}
