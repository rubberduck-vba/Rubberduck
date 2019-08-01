using System;
using System.Globalization;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// A class that represents a reference within a VBA project
    /// </summary>
    internal class TypeLibReference : ITypeLibReference
    {
        private readonly ITypeLibVBEExtensions _vbeTypeLib;
        private readonly int _typeLibIndex;

        public string RawString { get; }
        public Guid GUID { get; }
        public uint MajorVersion { get; }
        public uint MinorVersion { get; }
        public uint LCID { get; }
        public string Path { get; }
        public string Name { get; }

        public TypeLibReference(ITypeLibVBEExtensions vbeTypeLib, int typeLibIndex, string referenceStringRaw)
        {
            _vbeTypeLib = vbeTypeLib;
            _typeLibIndex = typeLibIndex;

            // Example: "*\G{000204EF-0000-0000-C000-000000000046}#4.1#9#C:\PROGRA~2\COMMON~1\MICROS~1\VBA\VBA7\VBE7.DLL#Visual Basic For Applications"
            // LibidReference defined at https://msdn.microsoft.com/en-us/library/dd922767(v=office.12).aspx
            // The string is split into 5 parts, delimited by #

            RawString = referenceStringRaw;

            var referenceStringParts = referenceStringRaw.Split(new char[] { '#' }, 5);
            if (referenceStringParts.Length != 5)
            {
                throw new ArgumentException($"Invalid reference string got {referenceStringRaw}.  Expected 5 parts.");
            }

            GUID = Guid.Parse(referenceStringParts[0].Substring(3));
            var versionSplit = referenceStringParts[1].Split(new char[] { '.' }, 2);
            if (versionSplit.Length != 2)
            {
                throw new ArgumentException($"Invalid reference string got {referenceStringRaw}.  Invalid version string.");
            }
            MajorVersion = uint.Parse(versionSplit[0], NumberStyles.AllowHexSpecifier);
            MinorVersion = uint.Parse(versionSplit[1], NumberStyles.AllowHexSpecifier);

            LCID = uint.Parse(referenceStringParts[2], NumberStyles.AllowHexSpecifier);
            Path = referenceStringParts[3];
            Name = referenceStringParts[4];
        }

        public ITypeLibWrapper TypeLib => _vbeTypeLib.GetVBEReferenceTypeLibByIndex(_typeLibIndex);
    }
}
