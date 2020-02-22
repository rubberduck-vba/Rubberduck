using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBProject : SafeComWrapper<VB.VBProject>, IVBProject
    {
        public VBProject(VB.VBProject target, bool rewrapping = false)
            :base(target, rewrapping)
        {
        }

        public IApplication Application => new Application(null);
		
        public IApplication Parent => new Application(null);

        public string ProjectId
        {
            get
            {
                try
                {
                    return IsWrappingNullReference ? string.Empty : Target.ReadProperty("Rubberduck", "ProjectId");
                }
                catch (COMException)
                {
                    return string.Empty;
                }
            }
        } 


        public string HelpFile
        {
            get => IsWrappingNullReference ? string.Empty : Target.HelpFile;
            set { if (!IsWrappingNullReference) Target.HelpFile = value; }
        }

        public string Description 
        {
            get => IsWrappingNullReference ? string.Empty : Target.Description;
            set { if (!IsWrappingNullReference) Target.Description = value; } 
        }

        public string Name
        {
            get => IsWrappingNullReference ? string.Empty : Target.Name;
            set { if (!IsWrappingNullReference) Target.Name = value; }
        }

        public EnvironmentMode Mode
        {
            get
            {
                if (IsWrappingNullReference)
                {
                    return 0;
                }

                // Note - the value returned by the EbMode function does NOT match with the EnvironmentMode enum, hence remapped below.
                var ebMode = EbMode();
                switch (ebMode)
                {
                    case 0:
                        return EnvironmentMode.Design;
                    case 1:
                        return EnvironmentMode.Run;
                    case 2:
                        return EnvironmentMode.Break;
                    default:
                        Debug.Assert(false, $"Unexpected value '{ebMode}' returned from EbMode");
                        return EnvironmentMode.Design;
                }                
            }
        }                   

        public IVBProjects Collection => new VBProjects(IsWrappingNullReference ? null : Target.Collection);

        public IReferences References => new References(IsWrappingNullReference ? null : Target.References);

        public IVBComponents VBComponents => new VBComponents(IsWrappingNullReference ? null : Target.VBComponents);

        public ProjectProtection Protection => ProjectProtection.Unprotected; // VB6 does not allow project protection

        public bool IsSaved => !IsWrappingNullReference && Target.Saved;

        public ProjectType Type => IsWrappingNullReference ? 0 : (ProjectType)Target.Type;

        public string FileName => IsWrappingNullReference ? string.Empty : Target.FileName;

        public string BuildFileName => IsWrappingNullReference ? string.Empty : Target.BuildFileName;

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public void SaveAs(string fileName)
        {
            if (!IsWrappingNullReference) Target.SaveAs(fileName);
        }

        public void MakeCompiledFile()
        {
            if (!IsWrappingNullReference) Target.MakeCompiledFile();
        }

        public override bool Equals(ISafeComWrapper<VB.VBProject> other)
        {
            // This is only safe in VB6 because project names must be unique within a session (which is not true of VBA)
            // Need to do it this way as reference equality fails when comparing to VBProjects.StartProject
            return IsEqualIfNull(other) || (other?.Target != null && other.Target.Name == Target.Name);
        }

        public bool Equals(IVBProject other)
        {
            return Equals(other as SafeComWrapper<VB.VBProject>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 
                : HashCode.Compute(Target);
        }

        public IReadOnlyList<string> ComponentNames()
        {
            var names = new List<string>();
            using (var components = VBComponents)
            {
                foreach (var component in components)
                {
                    using (component)
                    {
                        names.Add(component.Name);
                    }
                }

                return names.ToArray();
            }
        }

        public void AssignProjectId()
        {
            if (IsWrappingNullReference || !string.IsNullOrEmpty(ProjectId))
            {
                return;
            }
            Target.WriteProperty("Rubberduck", "ProjectId", Guid.NewGuid().ToString());
        }

        /// <summary>
        /// Exports all code modules in the VbProject to a destination directory. Files are given the same name as their parent code Module name and file extensions are based on what type of code Module it is.
        /// </summary>
        /// <param name="folder">The destination directory path.</param>
        public void ExportSourceFiles(string folder)
        {
            foreach (var component in VBComponents)
            {
                component.ExportAsSourceFile(folder);
            }
        }

        private string _displayName;
        /// <summary>
        /// WARNING: This property might have has side effects. If the filename cannot be accessed, it changes the ActiveVBProject, which causes a flicker in the VBE.
        /// This should only be called if it is *absolutely* necessary.
        /// </summary>
        public string ProjectDisplayName
        {
            get
            {
                if (_displayName != null)
                {
                    return _displayName;
                }

                if (IsWrappingNullReference)
                {
                    _displayName = string.Empty;
                    return _displayName;
                }

                _displayName = DisplayNameFromFileName();

                if (string.IsNullOrEmpty(_displayName))
                {
                    _displayName = DisplayNameFromWindowCaption();
                }

                if (string.IsNullOrEmpty(_displayName))
                {
                    _displayName = DisplayNameFromBuildFileName();
                }

                return _displayName;
            }
        }

        private string DisplayNameFromFileName()
        {
            return Path.GetFileName(FileName);
        }

        private string DisplayNameFromBuildFileName()
        {
            var pseudoDllName = Path.GetFileName(BuildFileName);
            return pseudoDllName == null || pseudoDllName.Length <= 4 //Should not happen as the string should always end in .DLL.
                ? string.Empty
                : pseudoDllName.Substring(0, pseudoDllName.Length - 4);
        }

        private static readonly Regex CaptionProjectRegex = new Regex(@"^(?<project>.+)(?:\s-\s)Microsoft Visual");
        private string DisplayNameFromWindowCaption()
        {
            using (var vbe = VBE)
            using (var activeProject = vbe.ActiveVBProject)
            using (var mainWindow = vbe.MainWindow)
            {
                try
                {
                    if (ProjectId != activeProject.ProjectId)
                    {
                        vbe.ActiveVBProject = this;
                    }

                    var caption = mainWindow.Caption;
                    if (CaptionProjectRegex.IsMatch(caption))
                    {
                        return CaptionProjectRegex
                            .Matches(caption)[0]
                            .Groups["project"]
                            .Value;
                    }

                    return string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);

        private const string DllName = "vba6.dll";
        [DllImport(DllName)]
        private static extern int EbMode();
    }
}