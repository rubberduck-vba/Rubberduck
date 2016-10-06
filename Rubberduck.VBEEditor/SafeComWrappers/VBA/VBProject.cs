using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBProject : SafeComWrapper<Microsoft.Vbe.Interop.VBProject>, IVBProject
    {
        public VBProject(Microsoft.Vbe.Interop.VBProject vbProject)
            :base(vbProject)
        {
        }

        public IApplication Application
        {
            get { return new Application(IsWrappingNullReference ? null : ComObject.Application); }
        }

        public IApplication Parent
        {
            get { return new Application(IsWrappingNullReference ? null : ComObject.Parent); }
        }

        public string HelpFile
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.HelpFile; }
            set { ComObject.HelpFile = value; }
        }

        public int HelpContextId
        {
            get { return IsWrappingNullReference ? 0 : ComObject.HelpContextID; }
            set { ComObject.HelpContextID = value; }
        }

        public string Description 
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Description; }
            set { ComObject.Description = value; } 
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Name; }
            set { ComObject.Name = value; }
        }

        public EnvironmentMode Mode
        {
            get { return IsWrappingNullReference ? 0 : (EnvironmentMode)ComObject.Mode; }
        }

        public IVBProjects Collection
        {
            get { return new VBProjects(IsWrappingNullReference ? null : ComObject.Collection); }
        }

        public IReferences References
        {
            get { return new References(IsWrappingNullReference ? null : ComObject.References); }
        }

        public IVBComponents VBComponents
        {
            get { return new VBComponents(IsWrappingNullReference ? null : ComObject.VBComponents); }
        }

        public ProjectProtection Protection
        {
            get { return IsWrappingNullReference ? 0 : (ProjectProtection)ComObject.Protection; }
        }

        public bool IsSaved
        {
            get { return !IsWrappingNullReference && ComObject.Saved; }
        }

        public ProjectType Type
        {
            get { return IsWrappingNullReference ? 0 : (ProjectType)ComObject.Type; }
        }

        public string FileName
        {
            get { return IsWrappingNullReference ? String.Empty : ComObject.FileName; }
        }

        public string BuildFileName
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.BuildFileName; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public void SaveAs(string fileName)
        {
            ComObject.SaveAs(fileName);
        }

        public void MakeCompiledFile()
        {
            ComObject.MakeCompiledFile();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                References.Release();
                VBComponents.Release();
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.VBProject> other)
        {
            return IsEqualIfNull(other) || (other != null && other.ComObject == ComObject);
        }

        public bool Equals(IVBProject other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.VBProject>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
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
    }
}