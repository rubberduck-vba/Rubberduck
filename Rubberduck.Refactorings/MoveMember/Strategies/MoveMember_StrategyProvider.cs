using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public static class MoveMemberStrategyProvider
    {
        public delegate bool StrategyValidator(IMoveScenario scenario, IProvideMoveDeclarationGroups groups);

        public delegate IMoveMemberRefactoringStrategy StrategyConstructor(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager);

        private static Dictionary<StrategyValidator, StrategyConstructor> MoveMemberStrategies
            = new Dictionary<StrategyValidator, StrategyConstructor>()
            {
                [SingleConstantToStdModule.IsApplicable] = SingleConstantToStdModule.CreateStrategy,
                [SingleFieldToStdModule.IsApplicable] = SingleFieldToStdModule.CreateStrategy,
                [SingleMemberNonPropertyToStdModule.IsApplicable] = SingleMemberNonPropertyToStdModule.CreateStrategy,
                //[SingleProcedureToStdModule.IsApplicable] = SingleProcedureToStdModule.CreateStrategy,
                //[SingleFunctionToStdModule.IsApplicable] = SingleFunctionToStdModule.CreateStrategy,
                //TODO: Future [SingleFunctionSelected_UseStateParams.IsApplicable] = SingleFunctionSelected_UseStateParams.CreateStrategy,
                //TODO: Future [SinglePropertySelected.IsApplicable] = SinglePropertySelected.CreateStrategy,
            };

        public static IEnumerable<IMoveMemberRefactoringStrategy> FindStrategies(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
        {
            Debug.Assert(scenario != null);
            var groups = scenario as IProvideMoveDeclarationGroups;

            var applicableStrategies = MoveMemberStrategies.Keys
                .Where(key => key(scenario, scenario as IProvideMoveDeclarationGroups)).Select(k => MoveMemberStrategies[k](scenario, rewritingManager));

            return applicableStrategies; 
        }
    }
}
