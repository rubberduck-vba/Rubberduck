using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA
{
    /// <summary>
    /// Extension methods for <c>IParseTree</c> and <c>ParserRuleContext</c>.
    /// </summary>
    public static class AntlrExtensions
    {
        /// <summary>
        /// Finds all public procedures in specified parse tree.
        /// </summary>
        public static IEnumerable<VisualBasic6Parser.SubStmtContext> GetPublicProcedures(this IParseTree parseTree)
        {
            var walker = new ParseTreeWalker();

            var listener = new PublicSubListener();
            walker.Walk(listener, parseTree);

            return listener.Members;
        }

        /// <summary>
        /// Gets the text of the specified line of code (first line if unspecified).
        /// </summary>
        /// <param name="context"></param>
        /// <param name="index">The line index to get value from.</param>
        /// <returns></returns>
        public static string GetLine(this ParserRuleContext context, int index = 0)
        {
            return context.GetText().Split('\n')[index];
        }

        /// <summary>
        /// Finds all procedures in specified parse tree.
        /// </summary>
        public static IEnumerable<ParserRuleContext> GetProcedures(this IParseTree parseTree)
        {
            return parseTree.GetContexts<ProcedureListener, ParserRuleContext>(new ProcedureListener());
        }

        public static IEnumerable<ParserRuleContext> GetProcedure(this IParseTree parseTree, string name)
        {
            return parseTree.GetContexts<ProcedureNameListener, ParserRuleContext>(new ProcedureNameListener(name));
        }

        public static IEnumerable<VisualBasic6Parser.ModuleOptionContext> GetModuleOptions(this IParseTree parseTree)
        {
            return parseTree.GetContexts<ModuleOptionsListener, VisualBasic6Parser.ModuleOptionContext>(new ModuleOptionsListener());
        }

        /// <summary>
        /// Finds all declarations in specified parse tree.
        /// </summary>
        public static IEnumerable<ParserRuleContext> GetDeclarations(this IParseTree parseTree)
        {
            return parseTree.GetContexts<DeclarationListener, ParserRuleContext>(new DeclarationListener());
        }
        
        /// <summary>
        /// Finds all variable references in specified parse tree.
        /// </summary>
        public static IEnumerable<VisualBasic6Parser.AmbiguousIdentifierContext> GetVariableReferences(this IParseTree parseTree)
        {
            return parseTree.GetContexts<VariableReferencesListener, VisualBasic6Parser.AmbiguousIdentifierContext>(new VariableReferencesListener());
        }

        private static IEnumerable<TContext> GetContexts<TListener, TContext>(this IParseTree parseTree, TListener listener)
            where TListener : IExtensionListener<TContext>, IParseTreeListener
            where TContext : ParserRuleContext
        {
            var walker = new ParseTreeWalker();
            walker.Walk(listener, parseTree);

            return listener.Members;
        }

        private interface IExtensionListener<out TContext> 
            where TContext : ParserRuleContext
        {
            IEnumerable<TContext> Members { get; }
        }

        private class VariableReferencesListener : VisualBasic6BaseListener,
            IExtensionListener<VisualBasic6Parser.AmbiguousIdentifierContext>
        {
            private readonly IList<VisualBasic6Parser.AmbiguousIdentifierContext> _members = new List<VisualBasic6Parser.AmbiguousIdentifierContext>(); 

            public IEnumerable<VisualBasic6Parser.AmbiguousIdentifierContext> Members { get { return _members; } }

            public override void EnterAmbiguousIdentifier(VisualBasic6Parser.AmbiguousIdentifierContext context)
            {
                _members.Add(context);
            }
        }

        private class DeclarationListener : VisualBasic6BaseListener, IExtensionListener<ParserRuleContext>
        {
            private readonly IList<ParserRuleContext> _members = new List<ParserRuleContext>();
            public IEnumerable<ParserRuleContext> Members { get { return _members; } }

            public override void EnterVariableStmt(VisualBasic6Parser.VariableStmtContext context)
            {
                _members.Add(context);
                foreach (var child in context.variableListStmt().variableSubStmt())
                {
                    _members.Add(child);
                }
            }

            public override void EnterEnumerationStmt(VisualBasic6Parser.EnumerationStmtContext context)
            {
                _members.Add(context);
            }

            public override void EnterConstStmt(VisualBasic6Parser.ConstStmtContext context)
            {
                _members.Add(context);
                foreach (var child in context.constSubStmt())
                {
                    _members.Add(child);
                }
            }

            public override void EnterTypeStmt(VisualBasic6Parser.TypeStmtContext context)
            {
                _members.Add(context);
            }

            public override void EnterDeclareStmt(VisualBasic6Parser.DeclareStmtContext context)
            {
                _members.Add(context);
            }

            public override void EnterEventStmt(VisualBasic6Parser.EventStmtContext context)
            {
                _members.Add(context);
            }

            public override void EnterArg(VisualBasic6Parser.ArgContext context)
            {
                _members.Add(context);
            }
        }

        private class ModuleOptionsListener : VisualBasic6BaseListener, IExtensionListener<VisualBasic6Parser.ModuleOptionContext>
        {
            private readonly IList<VisualBasic6Parser.ModuleOptionContext> _members = new List<VisualBasic6Parser.ModuleOptionContext>();
            public IEnumerable<VisualBasic6Parser.ModuleOptionContext> Members { get { return _members; } }

            public override void EnterModuleOptions(VisualBasic6Parser.ModuleOptionsContext context)
            {
                foreach (var option in context.moduleOption())
                {
                    _members.Add(option);
                }
            }
        }

        private class PublicSubListener : VisualBasic6BaseListener, IExtensionListener<VisualBasic6Parser.SubStmtContext>
        {
            private readonly IList<VisualBasic6Parser.SubStmtContext> _members = new List<VisualBasic6Parser.SubStmtContext>();
            public IEnumerable<VisualBasic6Parser.SubStmtContext> Members { get { return _members; } }

            public override void EnterSubStmt(VisualBasic6Parser.SubStmtContext context)
            {
                var visibility = context.visibility();
                if (visibility == null || visibility.PUBLIC() != null)
                {
                    _members.Add(context);
                }
            }
        }

        private class ProcedureListener : VisualBasic6BaseListener, IExtensionListener<ParserRuleContext>
        {
            private readonly IList<ParserRuleContext> _members = new List<ParserRuleContext>();
            public IEnumerable<ParserRuleContext> Members { get { return _members; } }

            public override void EnterSubStmt(VisualBasic6Parser.SubStmtContext context)
            {
                _members.Add(context);
            }

            public override void EnterFunctionStmt(VisualBasic6Parser.FunctionStmtContext context)
            {
                _members.Add(context);
            }

            public override void EnterPropertyGetStmt(VisualBasic6Parser.PropertyGetStmtContext context)
            {
                _members.Add(context);
            }

            public override void EnterPropertyLetStmt(VisualBasic6Parser.PropertyLetStmtContext context)
            {
                _members.Add(context);
            }

            public override void EnterPropertySetStmt(VisualBasic6Parser.PropertySetStmtContext context)
            {
                _members.Add(context);
            }
        }

        private class ProcedureNameListener : ProcedureListener
        {
            private readonly string _name;

            public ProcedureNameListener(string name)
            {
                _name = name;
            }

            public override void EnterFunctionStmt(VisualBasic6Parser.FunctionStmtContext context)
            {
                if (context.ambiguousIdentifier().GetText() == _name)
                {
                    base.EnterFunctionStmt(context);
                }
            }

            public override void EnterSubStmt(VisualBasic6Parser.SubStmtContext context)
            {
                if (context.ambiguousIdentifier().GetText() == _name)
                {
                    base.EnterSubStmt(context);
                }
            }

            public override void EnterPropertyGetStmt(VisualBasic6Parser.PropertyGetStmtContext context)
            {
                if (context.ambiguousIdentifier().GetText() == _name)
                {
                    base.EnterPropertyGetStmt(context);
                }
            }

            public override void EnterPropertyLetStmt(VisualBasic6Parser.PropertyLetStmtContext context)
            {
                if (context.ambiguousIdentifier().GetText() == _name)
                {
                    base.EnterPropertyLetStmt(context);
                }
            }

            public override void EnterPropertySetStmt(VisualBasic6Parser.PropertySetStmtContext context)
            {
                if (context.ambiguousIdentifier().GetText() == _name)
                {
                    base.EnterPropertySetStmt(context);
                }
            }
        }
    }
}