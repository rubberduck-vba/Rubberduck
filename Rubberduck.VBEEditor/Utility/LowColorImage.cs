using System;
using System.Diagnostics;
using System.IO;

namespace Rubberduck.VBEditor.Utility
{
    public class LowColorImage
    {
        private readonly byte[] _lowColorImageData;
        private int _offset;
        
        public byte[] ButtonFace { get; }
        public byte[] ButtonMask { get; }
        public byte[] DeviceIndependentBitmap { get; }
        public byte[] SystemDrawingBitmap { get; }
        public byte[] Bitmap { get; }
        public byte[] Format17 { get; }
        public bool IsValid { get; }

        public LowColorImage(byte[] lowColorImageData)
        {
            try
            {
                if (lowColorImageData == null)
                {
                    return;
                }
                _lowColorImageData = lowColorImageData;

                ButtonFace = GetNextSegment();
                ButtonMask = GetNextSegment();
                DeviceIndependentBitmap = GetNextSegment();
                SystemDrawingBitmap = GetNextSegment();
                Bitmap = GetNextSegment();
                Format17 = GetNextSegment();

                IsValid = true;
            }
            catch (Exception exception)
            {
                Debug.Assert(false, "Unable to load low color image", exception.ToString());
            }
        }

        private byte[] GetNextSegment()
        {
            byte[] ret;
            var size = BitConverter.ToInt32(_lowColorImageData, _offset);
            _offset += sizeof(int);

            using (var memoryStream = new MemoryStream(_lowColorImageData, _offset, size))
            {
                ret = memoryStream.ToArray();
            }
            _offset += size;
            return ret;
        }
    }
}
