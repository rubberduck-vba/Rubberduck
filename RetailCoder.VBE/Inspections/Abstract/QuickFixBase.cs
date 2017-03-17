using System.Globalization;
using System.Windows.Threading;
using Antlr4.Runtime;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class QuickFixBase : IQuickFix
    {
        private readonly ParserRuleContext _context;
        private readonly QualifiedSelection _selection;
        private readonly string _description;

        public QuickFixBase(ParserRuleContext context, QualifiedSelection selection, string description)
        {
            Dispatcher.CurrentDispatcher.Thread.CurrentCulture = CultureInfo.CurrentUICulture;
            Dispatcher.CurrentDispatcher.Thread.CurrentUICulture = CultureInfo.CurrentUICulture;

            _context = context;
            _selection = selection;
            _description = description;
        }

        public string Description { get { return _description; } }
        public ParserRuleContext Context { get { return _context; } }
        public QualifiedSelection Selection { get { return _selection; } }

        public bool IsCancelled { get; set; }

        public abstract void Fix();

        /// <summary>
        /// Indicates whether this quickfix can be applied to all inspection results in module.
        /// </summary>
        public virtual bool CanFixInModule { get { return true; } }

        /// <summary>
        /// Indicates whether this quickfix can be applied to all inspection results in procedure.
        /// </summary>
        public virtual bool CanFixInProcedure { get { return true; } }

        /// <summary>
        /// Indicates whether this quickfix can be applied to all inspection results in project.
        /// </summary>
        public virtual bool CanFixInProject { get { return true; } }
    }
}
