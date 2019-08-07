using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Extensions;

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
        public EmptyMethodInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var allInterfaces = new HashSet<ClassModuleDeclaration>(State.DeclarationFinder.FindAllUserInterfaces());

            return State.DeclarationFinder.UserDeclarations(DeclarationType.Member)
                .Where(member => !allInterfaces.Any(userInterface => userInterface.QualifiedModuleName == member.QualifiedModuleName)
                                 && !member.IsIgnoringInspectionResultFor(AnnotationName)
                                 && !((ModuleBodyElementDeclaration)member).Block.ContainsExecutableStatements())

                .Select(result => new DeclarationInspectionResult(this,
                                                                  string.Format(InspectionResults.EmptyMethodInspection,
                                                                                result.DeclarationType.ToFormatted(),
                                                                                result.IdentifierName),
                                                                  result));
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