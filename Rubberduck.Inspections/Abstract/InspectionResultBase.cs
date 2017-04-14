using System.IO;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class InspectionResultBase : IInspectionResult, INavigateSource, IExportable
    {
        protected InspectionResultBase(IInspection inspection, Declaration target)
            : this(inspection, target.QualifiedName.QualifiedModuleName, target.Context, target)
        { }
        
        protected InspectionResultBase(IInspection inspection, QualifiedModuleName qualifiedName, QualifiedMemberName? qualifiedMemberName, ParserRuleContext context)
        {
            Inspection = inspection;
            QualifiedName = qualifiedName;
            QualifiedMemberName = qualifiedMemberName;
            Context = context;
        }
        
        protected InspectionResultBase(IInspection inspection, QualifiedModuleName qualifiedName, ParserRuleContext context, Declaration target)
        {
            Inspection = inspection;
            QualifiedName = qualifiedName;
            Context = context;
            Target = target;

            QualifiedMemberName = GetQualifiedMemberName(target);
        }

        private QualifiedMemberName? GetQualifiedMemberName(Declaration target)
        {
            if (string.IsNullOrEmpty(target?.QualifiedName.QualifiedModuleName.ComponentName))
            {
                return null;
            }

            if (target.DeclarationType.HasFlag(DeclarationType.Member))
            {
                return target.QualifiedName;
            }

            return GetQualifiedMemberName(target.ParentDeclaration);
        }
        
        public IInspection Inspection { get; }

        public abstract string Description { get; }

        protected QualifiedModuleName QualifiedName { get; }

        public QualifiedMemberName? QualifiedMemberName { get; }

        public ParserRuleContext Context { get; }

        public Declaration Target { get; }

        /// <summary>
        /// Gets the information needed to select the target instruction in the VBE.
        /// </summary>
        public virtual QualifiedSelection QualifiedSelection
        {
            get
            {
                return Context == null
                    ? Target.QualifiedSelection
                    : new QualifiedSelection(QualifiedName, Context.GetSelection());
            }
        }

        public virtual int CompareTo(IInspectionResult other)
        {
            return Inspection.CompareTo(other.Inspection);
        }

        /// <summary>
        /// WARNING: This property can have side effects. It can change the ActiveVBProject if the result has a null Declaration, 
        /// which causes a flicker in the VBE. This should only be called if it is *absolutely* necessary.
        /// </summary>
        public string ToClipboardString()
        {           
            var module = QualifiedSelection.QualifiedName;
            var documentName = Target != null ? Target.ProjectDisplayName : string.Empty;
            if (string.IsNullOrEmpty(documentName))
            {
                var component = module.Component;
                documentName = component != null ? component.ParentProject.ProjectDisplayName : string.Empty;
            }
            if (string.IsNullOrEmpty(documentName))
            {
                documentName = Path.GetFileName(module.ProjectPath);
            }

            return string.Format(
                InspectionsUI.QualifiedSelectionInspection,
                Inspection.Severity,
                Description,
                "(" + documentName + ")",
                module.ProjectName,
                module.ComponentName,
                QualifiedSelection.Selection.StartLine);
        }

        public virtual NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(QualifiedSelection);
        }

        public int CompareTo(object obj)
        {
            return CompareTo(obj as IInspectionResult);
        }

        public object[] ToArray()
        {
            var module = QualifiedSelection.QualifiedName;
            return new object[] { Inspection.Severity.ToString(), module.ProjectName, module.ComponentName, Description, QualifiedSelection.Selection.StartLine, QualifiedSelection.Selection.StartColumn };
        }
    }
}
