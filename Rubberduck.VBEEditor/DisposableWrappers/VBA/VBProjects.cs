using System.Collections;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class VBProjects : SafeComWrapper<Microsoft.Vbe.Interop.VBProjects>, IEnumerable<VBProject>
    {
        public VBProjects(Microsoft.Vbe.Interop.VBProjects comObject) 
            : base(comObject)
        {
        }

        public VBProject Add(ProjectType type)
        {
            return new VBProject(InvokeResult(() => ComObject.Add((vbext_ProjectType)type)));
        }

        public void Remove(VBProject project)
        {
            Invoke(() => ComObject.Remove(project.ComObject));
        }

        public VBProject Open(string path)
        {
            return new VBProject(InvokeResult(() => ComObject.Open(path)));
        }

        public VBProject Item(object index)
        {
            return new VBProject(InvokeResult(() => ComObject.Item(index)));
        }

        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }
        public VBE Parent { get { return new VBE(InvokeResult(() => ComObject.Parent)); } }
        public int Count { get { return InvokeResult(() => ComObject.Count); } }

        IEnumerator<VBProject> IEnumerable<VBProject>.GetEnumerator()
        {
            return new ComWrapperEnumerator<Microsoft.Vbe.Interop.VBProjects, VBProject>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new ComWrapperEnumerator<Microsoft.Vbe.Interop.VBProjects, VBProject>(ComObject);
        }
    }
}