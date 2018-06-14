using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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
                using (var vbe = VBE)
                using (var mainWindow = vbe.MainWindow)
                {
                    var caption = mainWindow.Caption.ToLowerInvariant();

                    if (caption.EndsWith("[design]"))
                    {
                        return EnvironmentMode.Design;
                    }
                    if (caption.EndsWith("[break]"))
                    {
                        return EnvironmentMode.Break;
                    }
                    if (caption.EndsWith("[run]"))
                    {
                        return EnvironmentMode.Run;
                    }                    

                    // Shouldn't ever get here
                    Debug.Assert(false, "Unable to detect environment mode from main window caption");
                    return EnvironmentMode.Run; // Return most pessimistic result
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
            return IsEqualIfNull(other) || (other != null && other.Target == Target);
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
            return VBComponents.Select(component => component.Name).ToArray();
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

        private static readonly Regex CaptionProjectRegex = new Regex(@"^(?:[^-]+)(?:\s-\s)(?<project>.+)(?:\s-\s.*)?$");
        private static readonly Regex OpenModuleRegex = new Regex(@"^(?<project>.+)(?<module>\s-\s\[.*\((Code|UserForm)\)\])$");
        private string _displayName;
        /// <summary>
        /// WARNING: This property has side effects. It changes the ActiveVBProject, which causes a flicker in the VBE.
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

                var vbe = VBE;
                var activeProject = vbe.ActiveVBProject;
                var mainWindow = vbe.MainWindow;
                {
                    try
                    {
                        if (Target.HelpFile != activeProject.HelpFile)
                        {
                            vbe.ActiveVBProject = this;
                        }

                        var caption = mainWindow.Caption;
                        if (CaptionProjectRegex.IsMatch(caption))
                        {
                            caption = CaptionProjectRegex.Matches(caption)[0].Groups["project"].Value;
                            _displayName = OpenModuleRegex.IsMatch(caption)
                                ? OpenModuleRegex.Matches(caption)[0].Groups["project"].Value
                                : caption;
                        }
                    }
                    catch
                    {
                        _displayName = string.Empty;
                    }
                    return _displayName;
                }
            }
        }
    }
}