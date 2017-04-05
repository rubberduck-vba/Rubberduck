//using Antlr4.Runtime;
//using Rubberduck.Parsing.Inspections.Abstract;
//using Rubberduck.VBEditor;

//namespace Rubberduck.Inspections.Abstract
//{
//    public abstract class QuickFixBase : IQuickFix
//    {
//        protected QuickFixBase(ParserRuleContext context, QualifiedSelection selection, string description)
//        {
//            Context = context;
//            Selection = selection;
//            Description = description;
//        }

//        public ParserRuleContext Context { get; }

//        public string Description { get; }
//        public QualifiedSelection Selection { get; }

//        public bool IsCancelled { get; set; }

//        public abstract void Fix();

//        /// <summary>
//        /// Indicates whether this quickfix can be applied to all inspection results in module.
//        /// </summary>
//        public virtual bool CanFixInModule => true;

//        /// <summary>
//        /// Indicates whether this quickfix can be applied to all inspection results in procedure.
//        /// </summary>
//        public virtual bool CanFixInProcedure => true;

//        /// <summary>
//        /// Indicates whether this quickfix can be applied to all inspection results in project.
//        /// </summary>
//        public virtual bool CanFixInProject => true;
//    }
//}
