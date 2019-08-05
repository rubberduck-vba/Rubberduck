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
    [Experimental(nameof(ExperimentalNames.EmptyBlockInspections))]
    internal class EmptyMethodInspection : InspectionBase// EmptyBlockInspectionBase
    {
        public EmptyMethodInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            // Exclude empty members in user interfaces, as long as all members of the interface are empty,
            // since some VB users might use concrete user defined classes as interfaces,
            // while RD marks them as interfaces all the same.

            return UserDeclarations.OfType<ModuleBodyElementDeclaration>()
                .Where(bodyElement => !BlockContainsExecutableStatements(bodyElement.Block))
                .GroupBy(bodyElement => bodyElement.ComponentName)
                // Exclude results from user interfaces
                .Where(bodyElements => !State.DeclarationFinder.FindAllUserInterfaces()
                                       // where all members of that interface contain no executables
                                       .Where(interfaceModule => interfaceModule.ComponentName == bodyElements.Key
                                                                && interfaceModule.Members.Count == bodyElements.Count())
                                       .Any())
                .SelectMany(bodyElements => bodyElements)
                .Select(result => new DeclarationInspectionResult(this,
                                                                  string.Format(InspectionResults.EmptyMethodInspection,
                                                                                result.DeclarationType.ToFormatted(),
                                                                                result.IdentifierName),
                                                                  result));
        }

        private bool BlockContainsExecutableStatements(BlockContext block)
        {
            return block?.children != null && ContainsExecutableStatements(block.children);
        }

        private bool ContainsExecutableStatements(IList<Antlr4.Runtime.Tree.IParseTree> blockChildren)
        {
            foreach (var child in blockChildren)
            {
                if (child is BlockStmtContext blockStmt)
                {
                    var mainBlockStmt = blockStmt.mainBlockStmt();

                    if (mainBlockStmt == null)
                    {
                        continue;   //We have a lone line lable, which is not executable.
                    }

                    System.Diagnostics.Debug.Assert(mainBlockStmt.ChildCount == 1);

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