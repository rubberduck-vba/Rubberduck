using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBProject : SafeComWrapper<VB.VBProject>, IVBProject
    {
        public VBProject(VB.VBProject vbProject)
            :base(vbProject)
        {
        }

        public IApplication Application
        {
            get { return new Application(IsWrappingNullReference ? null : Target.Application); }
        }

        public IApplication Parent
        {
            get { return new Application(IsWrappingNullReference ? null : Target.Parent); }
        }

        public string ProjectId { get { return HelpFile; } }

        public string HelpFile
        {
            get { return IsWrappingNullReference ? string.Empty : Target.HelpFile; }
            set { if (!IsWrappingNullReference) Target.HelpFile = value; }
        }

        public string Description 
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Description; }
            set { if (!IsWrappingNullReference) Target.Description = value; } 
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Name; }
            set { if (!IsWrappingNullReference) Target.Name = value; }
        }

        public EnvironmentMode Mode
        {
            get { return IsWrappingNullReference ? 0 : (EnvironmentMode)Target.Mode; }
        }

        public IVBProjects Collection
        {
            get { return new VBProjects(IsWrappingNullReference ? null : Target.Collection); }
        }

        public IReferences References
        {
            get { return new References(IsWrappingNullReference ? null : Target.References); }
        }

        public IVBComponents VBComponents
        {
            get { return new VBComponents(IsWrappingNullReference ? null : Target.VBComponents); }
        }

        public ProjectProtection Protection
        {
            get { return IsWrappingNullReference ? 0 : (ProjectProtection)Target.Protection; }
        }

        public bool IsSaved
        {
            get { return !IsWrappingNullReference && Target.Saved; }
        }

        public ProjectType Type
        {
            get { return IsWrappingNullReference ? 0 : (ProjectType)Target.Type; }
        }

        public string FileName
        {
            get
            {
                try
                {
                    return IsWrappingNullReference ? string.Empty : Target.FileName;
                }
                catch (System.IO.IOException)
                {
                    // thrown by the VBIDE API when wrapped VBProject has no filename yet.
                    return string.Empty;
                }
            }
        }

        public string BuildFileName
        {
            get { return IsWrappingNullReference ? string.Empty : Target.BuildFileName; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public void SaveAs(string fileName)
        {
            if (!IsWrappingNullReference) Target.SaveAs(fileName);
        }

        public void MakeCompiledFile()
        {
            if (!IsWrappingNullReference) Target.MakeCompiledFile();
        }

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        if (Protection == ProjectProtection.Unprotected)
        //        {
        //            References.Release();
        //            VBComponents.Release();
        //        }
        //        base.Release(final);
        //    }
        //}

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
            //assign a hashcode if no helpfile is present
            if (string.IsNullOrEmpty(HelpFile))
            {
                HelpFile = GetHashCode().ToString();
            }

            //loop until the helpfile is unique for this host session
            while (!IsProjectIdUnique())
            {
                HelpFile = (GetHashCode() ^ HelpFile.GetHashCode()).ToString();
            }
        }

        private bool IsProjectIdUnique()
        {
            return VBE.VBProjects.Count(project => project.HelpFile == HelpFile) == 1;
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