using System;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class VBProject : SafeComWrapper<Microsoft.Vbe.Interop.VBProject>
    {
        public VBProject(Microsoft.Vbe.Interop.VBProject vbProject)
            :base(vbProject)
        {
        }

        public void SaveAs(string fileName)
        {
            Invoke(() => ComObject.SaveAs(fileName));
        }

        public void MakeCompiledFile()
        {
            Invoke(() => ComObject.MakeCompiledFile());
        }

        public Application Application { get { return new Application(InvokeResult(() => ComObject.Application)); } }

        public Application Parent { get { return new Application(InvokeResult(() => ComObject.Parent)); } }

        public string HelpFile
        {
            get { return InvokeResult(() => ComObject.HelpFile); }
            set { Invoke(() => ComObject.HelpFile = value); }
        }

        public int HelpContextID
        {
            get { return InvokeResult(() => ComObject.HelpContextID); }
            set  { Invoke(() => ComObject.HelpContextID = value); }
        }

        public string Description 
        {
            get { return InvokeResult(() => ComObject.Description); }
            set { Invoke(() => ComObject.Description = value); } 
        }

        public References References { get { return new References(InvokeResult(() => ComObject.References)); } }

        public string Name
        {
            get { return InvokeResult(() => ComObject.Name); }
            set { Invoke(() => ComObject.Name = value); }
        }

        public EnvironmentMode Mode { get { return (EnvironmentMode)InvokeResult(() => ComObject.Mode); } }

        public VBProjects Collection
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public ProjectProtection Protection { get { return (ProjectProtection)InvokeResult(() => ComObject.Protection); } }

        public bool Saved { get { return InvokeResult(() => ComObject.Saved); } }

        public VBComponents VBComponents
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public ProjectType Type { get { return (ProjectType)InvokeResult(() => ComObject.Type); } }

        public string FileName { get { return InvokeResult(() => ComObject.FileName); } }

        public string BuildFileName { get { return InvokeResult(() => ComObject.BuildFileName); } }

        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }
    }
}