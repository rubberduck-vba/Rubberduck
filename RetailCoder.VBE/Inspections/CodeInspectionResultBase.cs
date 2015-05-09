using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public abstract class CodeInspectionResultBase : ICodeInspectionResult
    {
        /// <summary>
        /// Creates a comment inspection result.
        /// </summary>
        protected CodeInspectionResultBase(string inspection, CodeInspectionSeverity type, CommentNode comment)
            : this(inspection, type, comment.QualifiedSelection.QualifiedName, null, comment)
        { }

        /// <summary>
        /// Creates an inspection result.
        /// </summary>
        protected CodeInspectionResultBase(string inspection, CodeInspectionSeverity type, QualifiedModuleName qualifiedName, ParserRuleContext context, CommentNode comment = null)
        {
            if (context == null && comment == null)
                throw new ArgumentNullException("[context] and [comment] cannot both be null.");

            _name = inspection;
            _type = type;
            _qualifiedName = qualifiedName;
            _context = context;
            _comment = comment;
        }

        private readonly string _name;
        /// <summary>
        /// Gets a string containing the name of the code inspection.
        /// </summary>
        public string Name { get { return _name; } }

        private readonly CodeInspectionSeverity _type;
        /// <summary>
        /// Gets the severity of the code issue.
        /// </summary>
        public CodeInspectionSeverity Severity { get { return _type; } }

        private readonly QualifiedModuleName _qualifiedName;
        protected QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private readonly ParserRuleContext _context;
        public ParserRuleContext Context { get { return _context; } }

        private readonly CommentNode _comment;
        public CommentNode Comment { get { return _comment; } }

        /// <summary>
        /// Gets the information needed to select the target instruction in the VBE.
        /// </summary>
        public virtual QualifiedSelection QualifiedSelection
        {
            get
            {
                return _context == null 
                    ? _comment.QualifiedSelection 
                    : new QualifiedSelection(_qualifiedName, _context.GetSelection());
            }
        }

        /// <summary>
        /// Gets all available "quick fixes" for a code inspection result.
        /// </summary>
        /// <returns>Returns a <c>Dictionary&lt;string&gt;, Action&lt;VBE&gt;</c>
        /// where the keys are descriptions for each quick fix, and
        /// each value is a method returning <c>void</c> and taking a <c>VBE</c> parameter.</returns>
        public abstract IDictionary<string, Action<VBE>> GetQuickFixes();

        public VBComponent FindComponent(VBE vbe)
        {
            var vbProject = vbe.VBProjects.Cast<VBProject>()
                .SingleOrDefault(project => project.Protection != vbext_ProjectProtection.vbext_pp_locked
                                         && project.Equals(QualifiedName.Project));

            if (vbProject == null)
            {
                return null;
            }

            return vbProject.VBComponents.Cast<VBComponent>()
                .SingleOrDefault(component => component.Equals(QualifiedName.Component));
        }
    }
}
