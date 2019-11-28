using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Common;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor;

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
            var finder = State.DeclarationFinder;

            var userInterfaces = UserInterfaces(finder);
            var emptyMethods = EmptyNonInterfaceMethods(finder, userInterfaces);

            return emptyMethods.Select(Result);
        }

        private static ICollection<QualifiedModuleName> UserInterfaces(DeclarationFinder finder)
        {
            return finder
                .FindAllUserInterfaces()
                .Select(decl => decl.QualifiedModuleName)
                .ToHashSet();
        }

        private static IEnumerable<Declaration> EmptyNonInterfaceMethods(DeclarationFinder finder, ICollection<QualifiedModuleName> userInterfaces)
        {
            return finder
                .UserDeclarations(DeclarationType.Member)
                .Where(member => !userInterfaces.Contains(member.QualifiedModuleName)
                                 && member is ModuleBodyElementDeclaration moduleBodyElement
                                 && !moduleBodyElement.Block.ContainsExecutableStatements());
        }

        private IInspectionResult Result(Declaration member)
        {
            return new DeclarationInspectionResult(
                this,
                ResultDescription(member),
                member);
        }

        private static string ResultDescription(Declaration member)
        {
            var identifierName = member.IdentifierName;
            var declarationType = member.DeclarationType.ToLocalizedString();

            return string.Format(
                InspectionResults.EmptyMethodInspection,
                declarationType,
                identifierName);
        }
    }
}