using System.Collections.Generic;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class CodeInspectionConfigProvider : IConfigProvider<CodeInspectionSettings>
    {
        private readonly IPersistanceService<CodeInspectionSettings> _persister;

        public CodeInspectionConfigProvider(IPersistanceService<CodeInspectionSettings> persister)
        {
            _persister = persister;
        }

        public CodeInspectionSettings Create()
        {
            var prototype = new CodeInspectionSettings(GetDefaultCodeInspections(), new WhitelistedIdentifierSetting[] { }, true);
            return _persister.Load(prototype) ?? prototype;
        }

        public CodeInspectionSettings CreateDefaults()
        {
            return new CodeInspectionSettings(GetDefaultCodeInspections(), new WhitelistedIdentifierSetting[] {}, true);
        }

        public void Save(CodeInspectionSettings settings)
        {
            _persister.Save(settings);
        }

        public HashSet<CodeInspectionSetting> GetDefaultCodeInspections()
        {
            // https://github.com/rubberduck-vba/Rubberduck/issues/3021
            return new HashSet<CodeInspectionSetting>
            {
                new CodeInspectionSetting("ApplicationWorksheetFunctionInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("AssignedByValParameterInspection", CodeInspectionType.LanguageOpportunities),
                new CodeInspectionSetting("ConstantNotUsedInspection", CodeInspectionType.CodeQualityIssues),
                new CodeInspectionSetting("DefaultProjectNameInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("EmptyIfBlockInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("EmptyStringLiteralInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("EncapsulatePublicFieldInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("FunctionReturnValueNotUsedInspection", CodeInspectionType.CodeQualityIssues),
                new CodeInspectionSetting("HostSpecificExpressionInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("HungarianNotationInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("IllegalAnnotationInspection", CodeInspectionType.RubberduckOpportunities, CodeInspectionSeverity.Error),
                new CodeInspectionSetting("ImplicitActiveSheetReferenceInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("ImplicitActiveWorkbookReferenceInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("ImplicitByRefParameterInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("ImplicitDefaultMemberAssignmentInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("ImplicitPublicMemberInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ImplicitVariantReturnTypeInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("MemberNotOnInterfaceInspection", CodeInspectionType.CodeQualityIssues),
                new CodeInspectionSetting("MissingAnnotationArgumentInspection", CodeInspectionType.RubberduckOpportunities, CodeInspectionSeverity.Error),
                new CodeInspectionSetting("MissingAnnotationInspection", CodeInspectionType.RubberduckOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("MissingAttributeInspection", CodeInspectionType.RubberduckOpportunities),
                new CodeInspectionSetting("ModuleScopeDimKeywordInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("MoveFieldCloserToUsageInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("MultilineParameterInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("MultipleDeclarationsInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("NonReturningFunctionInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error),
                new CodeInspectionSetting("ObjectVariableNotSetInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error),
                new CodeInspectionSetting("ObsoleteCallStatementInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ObsoleteCommentSyntaxInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ObsoleteGlobalInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ObsoleteLetStatementInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ObsoleteTypeHintInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("OptionBaseInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("OptionExplicitInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error),
                new CodeInspectionSetting("ParameterCanBeByValInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ParameterNotUsedInspection", CodeInspectionType.CodeQualityIssues),
                new CodeInspectionSetting("ProcedureCanBeWrittenAsFunctionInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ProcedureNotUsedInspection", CodeInspectionType.CodeQualityIssues),
                new CodeInspectionSetting("RedundantOptionInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("SelfAssignedDeclarationInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("UnassignedVariableUsageInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error),
                new CodeInspectionSetting("UndeclaredVariableUsageInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error),
                new CodeInspectionSetting("UntypedFunctionUsageInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("UseMeaningfulNameInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("VariableNotAssignedInspection", CodeInspectionType.CodeQualityIssues),
                new CodeInspectionSetting("VariableNotUsedInspection", CodeInspectionType.CodeQualityIssues),
                new CodeInspectionSetting("VariableTypeNotDeclaredInspection", CodeInspectionType.LanguageOpportunities),
                new CodeInspectionSetting("WriteOnlyPropertyInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Suggestion),
            };
        }
    }
}
