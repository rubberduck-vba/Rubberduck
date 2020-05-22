using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class InspectionResultBase : IInspectionResult, INavigateSource
    {
        protected InspectionResultBase(IInspection inspection,
            string description,
            QualifiedModuleName qualifiedName,
            ParserRuleContext context,
            Declaration target,
            QualifiedSelection qualifiedSelection,
            QualifiedMemberName? qualifiedMemberName,
            ICollection<string> disabledQuickFixes = null)
        {
            Inspection = inspection;
            Description = description?.Capitalize();
            QualifiedName = qualifiedName;
            Context = context;
            Target = target;
            QualifiedSelection = qualifiedSelection;
            QualifiedMemberName = qualifiedMemberName;
            DisabledQuickFixes = disabledQuickFixes ?? new List<string>();
        }

        public IInspection Inspection { get; }
        public string Description { get; }
        public QualifiedModuleName QualifiedName { get; }
        public QualifiedMemberName? QualifiedMemberName { get; }
        public ParserRuleContext Context { get; }
        public Declaration Target { get; }
        public ICollection<string> DisabledQuickFixes { get; }

        public virtual bool ChangesInvalidateResult(ICollection<QualifiedModuleName> modifiedModules)
        {
            return modifiedModules.Contains(QualifiedName) 
                   || Inspection.ChangesInvalidateResult(this, modifiedModules);
        }

        /// <summary>
        /// Gets the information needed to select the target instruction in the VBE.
        /// </summary>
        public QualifiedSelection QualifiedSelection { get; }

        public int CompareTo(IInspectionResult other)
        {
            return Inspection.CompareTo(other.Inspection);
        }

        private NavigateCodeEventArgs _navigationArgs;
        public NavigateCodeEventArgs GetNavigationArgs()
        {
            if (_navigationArgs != null)
            {
                return _navigationArgs;
            }

            _navigationArgs = new NavigateCodeEventArgs(QualifiedSelection);
            return _navigationArgs;
        }

        public int CompareTo(object obj)
        {
            return CompareTo(obj as IInspectionResult);
        }
    }
}
