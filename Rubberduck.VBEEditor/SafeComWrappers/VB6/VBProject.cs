using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBProject : SafeComWrapper<VB.VBProject>, IVBProject
    {
        public VBProject(VB.VBProject vbProject)
            :base(vbProject)
        {
        }

        public IApplication Application
        {
            get { throw new NotImplementedException(); }
        }

        public IApplication Parent
        {
            get { throw new NotImplementedException(); }
        }

        public string ProjectId { get { return HelpFile; } }

        public string HelpFile
        {
            get { return IsWrappingNullReference ? string.Empty : Target.HelpFile; }
            set { Target.HelpFile = value; }
        }

        public string Description 
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Description; }
            set { Target.Description = value; } 
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Name; }
            set { Target.Name = value; }
        }

        public EnvironmentMode Mode
        {
            get { throw new NotImplementedException(); }
        }

        public IVBProjects Collection
        {
            get { return new VBProjects(IsWrappingNullReference ? null : Target.Collection); }
        }

        public IReferences References
        {
            get { throw new NotImplementedException(); }
        }

        public IVBComponents VBComponents
        {
            get { throw new NotImplementedException(); }
        }

        public ProjectProtection Protection
        {
            get { throw new NotImplementedException(); }
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
            get { return IsWrappingNullReference ? string.Empty : Target.FileName; }
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
            Target.SaveAs(fileName);
        }

        public void MakeCompiledFile()
        {
            Target.MakeCompiledFile();
        }

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        References.Release();
        //        VBComponents.Release();
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

        public string ProjectDisplayName { get; private set; }
    }
}