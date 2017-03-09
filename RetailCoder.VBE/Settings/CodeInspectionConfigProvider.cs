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
            //This no longer sucks.
            return new CodeInspectionSettings(GetDefaultCodeInspections(), new WhitelistedIdentifierSetting[] {}, true);
        }

        public void Save(CodeInspectionSettings settings)
        {
            _persister.Save(settings);
        }

        public HashSet<CodeInspectionSetting> GetDefaultCodeInspections()
        {
            //*This* sucks now.
            return new HashSet<CodeInspectionSetting>
            {
                new CodeInspectionSetting("ObjectVariableNotSetInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("FunctionReturnValueNotUsedInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("SelfAssignedDeclarationInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("MoveFieldCloserToUsageInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("EncapsulatePublicFieldInspection", string.Empty, CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("EmptyStringLiteralInspection", string.Empty, CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ImplicitActiveSheetReferenceInspection", string.Empty, CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ImplicitActiveWorkbookReferenceInspection", string.Empty, CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("MultipleFolderAnnotationsInspection", string.Empty, CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Error,  CodeInspectionSeverity.Error),
                new CodeInspectionSetting("ProcedureCanBeWrittenAsFunctionInspection", string.Empty, CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("UseMeaningfulNameInspection", string.Empty, CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("WriteOnlyPropertyInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("UntypedFunctionUsageInspection", string.Empty, CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("AssignedByValParameterInspection", string.Empty, CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ConstantNotUsedInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("DefaultProjectNameInspection", string.Empty, CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ImplicitPublicMemberInspection", string.Empty, CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("MultilineParameterInspection", string.Empty, CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("NonReturningFunctionInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error,  CodeInspectionSeverity.Error),
                new CodeInspectionSetting("ObsoleteCallStatementInspection", string.Empty, CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ObsoleteGlobalInspection", string.Empty, CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ObsoleteLetStatementInspection", string.Empty, CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ObsoleteTypeHintInspection", string.Empty, CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("OptionBaseInspection", string.Empty, CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("ParameterCanBeByValInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ParameterNotUsedInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ProcedureNotUsedInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("UnassignedVariableUsageInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error,  CodeInspectionSeverity.Error),
                new CodeInspectionSetting("VariableNotUsedInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("VariableNotAssignedInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ImplicitByRefParameterInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("ImplicitVariantReturnTypeInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("MultipleDeclarationsInspection", string.Empty, CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ObsoleteCommentSyntaxInspection", string.Empty, CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("OptionExplicitInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error,  CodeInspectionSeverity.Error),
                new CodeInspectionSetting("VariableTypeNotDeclaredInspection", string.Empty, CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("MissingAnnotationArgumentInspection", string.Empty, CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error,  CodeInspectionSeverity.Error)
            };
        }
    }
}
