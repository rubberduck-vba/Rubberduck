using System;
using System.IO;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class InspectionResult : IInspectionResult, INavigateSource, IExportable
    {
        public InspectionResult(IInspection inspection, string description, Declaration target)
            : this(inspection, description, new QualifiedContext<ParserRuleContext>(target.QualifiedName.QualifiedModuleName, target.Context), target)
        { }

        public InspectionResult(IInspection inspection, string description, RubberduckParserState state, IdentifierReference reference)
            : this(inspection, description, new QualifiedContext<ParserRuleContext>(reference.QualifiedModuleName, reference.Context), reference.Declaration, false)
        {
            QualifiedMemberName = GetQualifiedMemberName(state, reference);
        }

        public InspectionResult(IInspection inspection, string description, RubberduckParserState state, QualifiedContext context)
        {
            Inspection = inspection;
            Description = description.Capitalize();
            QualifiedName = context.ModuleName;
            QualifiedMemberName = GetQualifiedMemberName(state, context);
            Context = context.Context;

            QualifiedSelection = new QualifiedSelection(QualifiedName, Context.GetSelection());
            _navigationArgs = new Lazy<NavigateCodeEventArgs>(() => new NavigateCodeEventArgs(QualifiedSelection));
        }

        public InspectionResult(IInspection inspection, string description, QualifiedContext context, Declaration target, bool navigateToTarget = true)
        {
            Inspection = inspection;
            Description = description.Capitalize();
            QualifiedName = context.ModuleName;
            Context = context.Context;
            Target = target;
            QualifiedSelection = navigateToTarget
                    ? Target.QualifiedSelection
                    : new QualifiedSelection(QualifiedName, Context.GetSelection());
            _navigationArgs = new Lazy<NavigateCodeEventArgs>(() => new NavigateCodeEventArgs(QualifiedSelection));

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

        private QualifiedMemberName? GetQualifiedMemberName(RubberduckParserState state, QualifiedContext context)
        {
            var members = state.DeclarationFinder.Members(context.ModuleName);
            return members.SingleOrDefault(m => m.Selection.Contains(context.Context.GetSelection()))?.QualifiedName;
        }

        private QualifiedMemberName? GetQualifiedMemberName(RubberduckParserState state, IdentifierReference reference)
        {
            var members = state.DeclarationFinder.Members(reference.QualifiedModuleName);
            return members.SingleOrDefault(m => m.Selection.Contains(reference.Selection))?.QualifiedName;
        }

        public IInspection Inspection { get; }

        public string Description { get; }

        public QualifiedModuleName QualifiedName { get; }

        public QualifiedMemberName? QualifiedMemberName { get; }

        public ParserRuleContext Context { get; }

        public Declaration Target { get; }

        /// <summary>
        /// Gets the information needed to select the target instruction in the VBE.
        /// </summary>
        public QualifiedSelection QualifiedSelection { get; }

        public int CompareTo(IInspectionResult other)
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

        private readonly Lazy<NavigateCodeEventArgs> _navigationArgs;
        public NavigateCodeEventArgs GetNavigationArgs() => _navigationArgs.Value;

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
