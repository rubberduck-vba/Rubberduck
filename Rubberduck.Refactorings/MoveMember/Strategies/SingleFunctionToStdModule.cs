using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    //public class SingleFunctionToStdModule : SingleMemberNonPropertyToStdModule
    //{
    //    public static bool IsApplicable(IMoveScenario scenario, IProvideMoveDeclarationGroups groups)
    //    {
    //        return IsApplicable(scenario, groups);
    //    }

    //    public static IMoveMemberRefactoringStrategy CreateStrategy(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
    //        => new SingleFunctionToStdModule(scenario, rewritingManager);

    //    private readonly IMoveScenario _scenario;
    //    private readonly MoveMemberStrategyCommon _helper;

    //    private SingleFunctionToStdModule(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
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
    //        //    var function = _scenario.SelectedDeclarations.Where(d => d.IsMember()).Single();
    //        //    var _FUNCTION_WITH_ARGS_BODY_FORMAT = "    {0} = {1}.{0}({2})";
    //        //    var newBody = string.Format(_FUNCTION_WITH_ARGS_BODY_FORMAT, function.IdentifierName, MemberAccessLExpr, _helper.CallSiteArguments(function));
    //        //    sourceRewriter.ReplaceDescendentContext<VBAParser.BlockContext>(function, $"{newBody}{Environment.NewLine}");
    //        //}

    //        //ReplaceMovedOrRenamedReferenceIdentifiers(sourceRewriter);

    //        _helper.UpdateSourceReferencesToMovedElements(sourceRewriter);

    //        _helper.InsertNewSourceContent(sourceRewriter);
    //    }

    //    //private string MemberAccessLExpr
    //    //        => _scenario.MoveDefinition.IsClassModuleDestination ?
    //    //            _helper.ForwardToModuleLExpression()
    //    //                : _scenario.MoveDefinition.Destination.ModuleName;
    //}
}
