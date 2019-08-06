using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Experimentals;
using static Rubberduck.Parsing.Grammar.VBAParser;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Concrete.Extensions;
using Antlr4.Runtime.Tree;
using System;
using static Rubberduck.Inspections.Concrete.Extensions.EmptyMethodInspectionMeasures;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies empty module member blocks.
    /// </summary>
    /// <why>
    /// Methods containing no executable statements are misleading as they appear to be doing something which they actually don't.
    /// This might be the result of delaying the actual implementation for a later stage of development, and then forgetting all about that.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Sub Foo()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Sub Foo()
    ///     MsgBox "?"
    /// End Sub
    /// ]]>
    /// </example>
    internal class EmptyMethodInspection : InspectionBase
    {
        public EmptyMethodInspection(RubberduckParserState state, bool measure)
            :base(state)
        {

        }

        public EmptyMethodInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<Declaration> UserDeclarations => base.UserDeclarations;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            // Exclude empty members in user interfaces, as long as all members of the interface are empty,
            // since some VB users might use concrete user defined classes as interfaces,
            // while RD marks them as interfaces all the same.

            //var firstImplementation = UserDeclarations.OfType<ModuleBodyElementDeclaration>()
            //    .Where(bodyElement => !BlockContainsExecutableStatements(bodyElement.Block))
            //    .GroupBy(bodyElement => bodyElement.QualifiedModuleName);

            var allInterfaces = new HashSet<ClassModuleDeclaration>(State.DeclarationFinder.FindAllUserInterfaces());

            return State.DeclarationFinder.UserDeclarations(DeclarationType.Member)
                .Where(member => !member.IsIgnoringInspectionResultFor(AnnotationName)
                                  && !BlockContainsExecutableStatements(((ModuleBodyElementDeclaration)member).Block))
                
                // Exclude results from user PURE interfaces only - → → → ↓↓↓↓
                .GroupBy(bodyElement => bodyElement.QualifiedModuleName)
                .Where(bodyElements => !allInterfaces
                                            // by filtering out any classModules as long ALL its members contain no executables
                                            // which means that it IS used as concrete and SHOULD be inspected for empty members
                                            .Any(interfaceModule => interfaceModule.QualifiedModuleName == bodyElements.Key
                                                                    && interfaceModule.Members.Count == bodyElements.Count()))
                .SelectMany(bodyElements => bodyElements)
                .Select(result => new DeclarationInspectionResult(this,
                                                                  string.Format(InspectionResults.EmptyMethodInspection,
                                                                                result.DeclarationType.ToFormatted(),
                                                                                result.IdentifierName),
                                                                  result));
        }

        private bool BlockContainsExecutableStatements(BlockContext block)
        {
            bool result1 = false;
            bool result2 = false;
            int reps = 200;
            string text = State.DeclarationFinder.UserDeclarations(DeclarationType.Member).First().Context.GetText();
            Measure(text, Syntax.Old, reps, () =>
                result1 = block?.children != null && ContainsExecutableStatements(block.children));

            Measure(text, Syntax.Linq, reps, () =>
                result2 = block?.children != null && ContainsExecutableStatementsLinq(block.children));

            System.Diagnostics.Debug.Assert(result1 == result2);
            return result1;
        }

        private bool ContainsExecutableStatements(IList<IParseTree> blockChildren)
        {
            foreach (var child in blockChildren)
            {
                if (child is BlockStmtContext blockStmt)
                {
                    var mainBlockStmt = blockStmt.mainBlockStmt();

                    if (mainBlockStmt == null)
                    {
                        continue;   //We have a lone line label, which is not executable.
                    }

                    // exclude variables and consts because they are not executable statements
                    if (mainBlockStmt.GetChild(0) is VariableStmtContext ||
                        mainBlockStmt.GetChild(0) is ConstStmtContext)
                    {
                        continue;
                    }

                    return true;
                }

                if (child is RemCommentContext ||
                    child is CommentContext ||
                    child is CommentOrAnnotationContext ||
                    child is EndOfStatementContext)
                {
                    continue;
                }

                return true;
            }

            return false;
        }

        private bool ContainsExecutableStatementsLinq(IList<IParseTree> blockChildren)
        {
            return blockChildren.Any(child => (child is BlockStmtContext blockStmt
                                               && !(blockStmt.mainBlockStmt() == null
                                                    || IsConstantOrVariable(blockStmt))));
        }

        private bool IsConstantOrVariable(BlockStmtContext blockStmt)
        {
            var context = blockStmt.mainBlockStmt().GetChild(0);
            return context is VariableStmtContext || context is ConstStmtContext;
        }
    }

    public static class DeclarationTypeFormat
    {
        public static string ToFormatted(this DeclarationType declarationType)
        {
            string result = declarationType.ToString();
            int length = result.Length;

            for (int i = 1; i < length; i++)
            {
                if (char.IsUpper(result[i]))
                {
                    result = result.Insert(i++, " ");
                    length++;
                }
            }

            return result;
        }
    }
}