using System;
using Kavod.ComReflection;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class RegisteredLibraryModel : VbaReferenceModel
    {
        private readonly LibraryRegistration _registration;

        internal RegisteredLibraryModel(LibraryRegistration registration)
        {
            _registration = registration;
        }

        public override string FilePath { get { return _registration.FilePath; } }

        public override string Name { get { return _registration.Name; } }

        public override short MajorVersion { get { return _registration.MajorVersion; } }

        public override short MinorVersion { get { return _registration.MinorVersion; } }

        public override Guid Guid { get { return _registration.Guid; } }
    }
}