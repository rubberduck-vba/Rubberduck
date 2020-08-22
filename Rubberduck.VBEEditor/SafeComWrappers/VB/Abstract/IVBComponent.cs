using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBComponent : ISafeComWrapper, IEquatable<IVBComponent>
    {
        ComponentType Type { get; }
        bool HasCodeModule { get; }
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
        string ExportAsSourceFile(string folder, bool isTempFile = false, bool specialCaseDocumentModules = true);
        int FileCount { get; }
        string GetFileName(short index);
        IVBProject ParentProject { get; }
        int ContentHash();

        QualifiedModuleName QualifiedModuleName { get; }
    }
}