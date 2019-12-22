using System.Collections.Generic;
using System.IO;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Interaction.Navigation;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class InspectionResultBase : IInspectionResult, INavigateSource
    {
        protected InspectionResultBase(IInspection inspection,
            string description,
            QualifiedModuleName qualifiedName,
            ParserRuleContext context,
            Declaration target,
            QualifiedSelection qualifiedSelection,
            QualifiedMemberName? qualifiedMemberName,
            dynamic properties)
        {
            Inspection = inspection;
            Description = description?.Capitalize();
            QualifiedName = qualifiedName;
            Context = context;
            Target = target;
            QualifiedSelection = qualifiedSelection;
            QualifiedMemberName = qualifiedMemberName;
            Properties = properties ?? new PropertyBag();
        }

        public IInspection Inspection { get; }
        public string Description { get; }
        public QualifiedModuleName QualifiedName { get; }
        public QualifiedMemberName? QualifiedMemberName { get; }
        public ParserRuleContext Context { get; }
        public Declaration Target { get; }
        public dynamic Properties { get; }

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
