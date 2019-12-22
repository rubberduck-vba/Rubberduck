using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about member calls against an extensible interface, that cannot be validated at compile-time.
    /// </summary>
    /// <why>
    /// Extensible COM types can have members attached at run-time; VBA cannot bind these member calls at compile-time.
    /// If there is an early-bound alternative way to achieve the same result, it should be preferred.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal adoConnection As ADODB.Connection)
    ///     adoConnection.SomeStoredProcedure 42
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal adoConnection As ADODB.Connection)
    ///     Dim adoCommand As ADODB.Command
    ///     Set adoCommand.ActiveConnection = adoConnection
    ///     adoCommand.CommandText = "SomeStoredProcedure"
    ///     adoCommand.CommandType = adCmdStoredProc
    ///     adoCommand.Parameters.Append adocommand.CreateParameter(Value:=42)
    ///     adoCommand.Execute
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class MemberNotOnInterfaceInspection : InspectionBase
    {
        public MemberNotOnInterfaceInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            // prefilter to reduce searchspace
            var unresolved = State.DeclarationFinder.UnresolvedMemberDeclarations()
                .Where(decl => !decl.IsIgnoringInspectionResultFor(AnnotationName)).ToList();

            var targets = Declarations.Where(decl => decl.AsTypeDeclaration != null &&
                                                     !decl.AsTypeDeclaration.IsUserDefined &&
                                                     decl.AsTypeDeclaration.DeclarationType.HasFlag(DeclarationType.ClassModule) &&                                                    
                                                     ((ClassModuleDeclaration)decl.AsTypeDeclaration).IsExtensible)
                                       .SelectMany(decl => decl.References).ToList();
            return unresolved
                .Select(access => new
                {
                    access,
                    callingContext = targets.FirstOrDefault(usage => usage.Context.Equals(access.CallingContext)
                                                                     || (access.CallingContext is VBAParser.NewExprContext && 
                                                                         usage.Context.Parent.Parent.Equals(access.CallingContext))
                                                                     )
                })
                .Where(memberAccess => memberAccess.callingContext != null &&
                                       memberAccess.callingContext.Declaration.DeclarationType != DeclarationType.Control)    //TODO - remove this exception after resolving #2592)
                .Select(memberAccess => new DeclarationInspectionResult(this,
                    string.Format(InspectionResults.MemberNotOnInterfaceInspection, memberAccess.access.IdentifierName,
                        memberAccess.callingContext.Declaration.AsTypeDeclaration.IdentifierName), memberAccess.access));
        }
    }
}
