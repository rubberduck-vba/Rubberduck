using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Refactorings.MoveMember
{
    //public class SingleProcedureToStdModule : SingleMemberToStdModule
    //{
    //    public static bool IsApplicable(IMoveScenario scenario, IProvideMoveDeclarationGroups groups)
    //    {
    //        return IsApplicable(scenario, groups, DeclarationType.Procedure);
    //    }

    //    private static string[] ValueTypes = new string[]
    //    {
    //        Tokens.Boolean,
    //        Tokens.Currency,
    //        Tokens.Long,
    //        Tokens.LongLong,
    //        Tokens.Single,
    //        Tokens.Double,
    //        Tokens.String,
    //        Tokens.Byte
    //    };

    //    public static IMoveMemberRefactoringStrategy CreateStrategy(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
    //        => new SingleProcedureToStdModule(scenario, rewritingManager);

    //    private readonly IMoveScenario _scenario;
    //    private readonly MoveMemberStrategyCommon _helper;

    //    private SingleProcedureToStdModule(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
    //    {
    //        _scenario = scenario;
    //        _helper = new MoveMemberStrategyCommon(scenario, rewritingManager);
    //    }

    //    public override void ModifyContent() => _helper.ModifyContent(ModifySource);

    //    public override string PreviewDestination() => _helper.PreviewDestination();

    //    public override string DestinationMemberCodeBlock(Declaration member) => _helper.DestinationMemberCodeBlockDefault(member);

    //    public override string DestinationNewModuleContent => _helper.DestinationNewModuleContent;

    //    public override int DestinationNewContentLineCount => _helper.DestinationNewContentLineCount;

    //    private void ModifySource(IMoveEndpointRewriter sourceRewriter)
    //    {
    //        _helper.RemoveDeclarations(sourceRewriter);

    //        //if (_scenario.ForwardSelectedMemberCalls)
    //        //{
    //        //    var procedure = _scenario.SelectedDeclarations.Where(d => d.IsMember()).Single();
    //        //    var _PROCEDURE_BODY_FORMAT = "    {0}.{1} {2}";
    //        //    var newBody = string.Format(_PROCEDURE_BODY_FORMAT, MemberAccessLExpr, procedure.IdentifierName, _helper.CallSiteArguments(procedure));
    //        //    sourceRewriter.ReplaceDescendentContext<VBAParser.BlockContext>(procedure, $"{newBody}{Environment.NewLine}");
    //        //}

    //        //_helper.ReplaceMovedOrRenamedReferenceIdentifiers(sourceRewriter);

    //        _helper.UpdateSourceReferencesToMovedElements(sourceRewriter);

    //        //_helper.EnsureClassIsValidWhereReferenced(sourceRewriter);

    //        _helper.InsertNewSourceContent(sourceRewriter);
    //    }

    //    //private string MemberAccessLExpr
    //    //        => _scenario.MoveDefinition.IsClassModuleDestination ?
    //    //            _helper.ForwardToModuleLExpression()
    //    //                : _scenario.MoveDefinition.Destination.ModuleName;
    //}
}
