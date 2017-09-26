using System;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBComponent : ISafeComWrapper, IEquatable<IVBComponent>
    {
        ComponentType Type { get; }
        ICodeModule CodeModule { get; }
        IVBE VBE { get; }
        IVBComponents Collection { get; }
        IProperties Properties { get; }
        IControls Controls { get; }
        IControls SelectedControls { get; }
        bool IsSaved { get; }
        bool HasDesigner { get; }
        bool HasOpenDesigner { get; }
        string DesignerId { get; }
        string Name { get; set; }
        IWindow DesignerWindow();
        void Activate();
        void Export(string path);
        string ExportAsSourceFile(string folder, bool tempFile = false);

        IVBProject ParentProject { get; }
    }
}