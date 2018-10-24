using System;

namespace Rubberduck.VBEditor.Events
{
    public class ComponentEventArgs : EventArgs
    {
        public ComponentEventArgs(QualifiedModuleName qualifiedModuleName)
        {
            QualifiedModuleName = qualifiedModuleName;
        }

        public string ProjectId => QualifiedModuleName.ProjectId;
        public QualifiedModuleName QualifiedModuleName { get; }
    }
}