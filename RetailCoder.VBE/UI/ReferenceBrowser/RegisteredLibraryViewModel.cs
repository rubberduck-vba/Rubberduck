using System;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class RegisteredLibraryViewModel : ViewModelBase, IComparable<RegisteredLibraryViewModel>
    {
        private readonly VBProject _activeProject;
        private bool _isActiveReference;
        private bool _referenceIsRemovable = true;

        public RegisteredLibraryViewModel(VbaReferenceModel model, VBProject activeProject)
        {
            _activeProject = activeProject;
            Model = model;

            _isActiveReference = GetIsActiveProjectReference();
            if (IsActiveProjectReference)
            {
                var reference = GetActiveProjectReferenceByFilePath(FilePath);
                CanRemoveReference = !reference.BuiltIn;
            }
        }

        public VbaReferenceModel Model { get; private set; }

        public string FilePath
        {
            get { return Model.FilePath; }
        }

        public string Name
        {
            get { return Model.Name; }
        }

        public bool IsActiveProjectReference
        {
            get { return _isActiveReference; }
            set
            {
                if (value == _isActiveReference)
                {
                    return;
                }

                if (value)
                {
                    AddReferenceToActiveProject();
                    _isActiveReference = true;
                }
                else
                {
                    // this may not succeed, so we can't just change the value.
                    var result = RemoveReferenceFromActiveProject();
                    if (result)
                    {
                        _isActiveReference = false;
                    }
                    else
                    {
                        // TODO warn the user that they cannot remove this reference.
                        // The user shouldn't be able to remove a builtin reference anyway.
                    }
                }
                OnPropertyChanged();
            }
        }

        public bool CanRemoveReference
        {
            get { return _referenceIsRemovable; }
            set
            {
                if (value == _referenceIsRemovable)
                {
                    return;
                }
                _referenceIsRemovable = value;
                OnPropertyChanged();
            }
        }

        public Guid Guid
        {
            get { return Model.Guid; }
        }

        public short MinorVersion
        {
            get { return Model.MinorVersion; }
        }

        public short MajorVersion
        {
            get { return Model.MajorVersion; }
        }

        private void AddReferenceToActiveProject()
        {
            _activeProject.References.AddFromFile(FilePath);
        }

        private bool RemoveReferenceFromActiveProject()
        {
            // Note that it is possible that we may not be able to remove some references that
            // are used by the VBE.  Trying to do so throws a COMException.
            var reference = GetActiveProjectReferenceByFilePath(FilePath);
            try
            {
                _activeProject.References.Remove(reference);
                return true;
            }
            catch (COMException ex)
            {
                // check it is the COM Exception we were expecting.
                if (ex.Message.ToLower().Contains("default reference"))
                {
                    return false;
                }
                throw;
            }
        }

        private bool GetIsActiveProjectReference()
        {
            return GetActiveProjectReferenceByFilePath(FilePath) != null;
        }

        private Reference GetActiveProjectReferenceByFilePath(string filePath)
        {
            return _activeProject.References
                .OfType<Reference>()
                .SingleOrDefault(r => r.FullPath == filePath);
        }

        #region IComparable

        public int CompareTo(RegisteredLibraryViewModel other)
        {
            if (CanRemoveReference && !other.CanRemoveReference)
            {
                return 1;
            }
            if (!CanRemoveReference && other.CanRemoveReference)
            {
                return -1;
            }
            if (IsActiveProjectReference && !other.IsActiveProjectReference)
            {
                return -1;
            }
            if (!IsActiveProjectReference && other.IsActiveProjectReference)
            {
                return 1;
            }
            return string.Compare(this.Name, other.Name, StringComparison.InvariantCultureIgnoreCase);
        }

        #endregion
    }
}