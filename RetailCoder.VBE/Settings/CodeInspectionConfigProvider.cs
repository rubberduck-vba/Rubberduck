using System.Collections.Generic;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
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
                new CodeInspectionSetting("ObjectVariableNotSetInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("FunctionReturnValueNotUsedInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("SelfAssignedDeclarationInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("MoveFieldCloserToUsageInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("EncapsulatePublicFieldInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("EmptyStringLiteralInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ImplicitActiveSheetReferenceInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ImplicitActiveWorkbookReferenceInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("MultipleFolderAnnotationsInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Error,  CodeInspectionSeverity.Error),
                new CodeInspectionSetting("ProcedureCanBeWrittenAsFunctionInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("UseMeaningfulNameInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("WriteOnlyPropertyInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("UntypedFunctionUsageInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("AssignedByValParameterInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ConstantNotUsedInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("DefaultProjectNameInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ImplicitPublicMemberInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("MultilineParameterInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("NonReturningFunctionInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error,  CodeInspectionSeverity.Error),
                new CodeInspectionSetting("ObsoleteCallStatementInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ObsoleteGlobalInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ObsoleteLetStatementInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ObsoleteTypeHintInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("OptionBaseInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("ParameterCanBeByValInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("ParameterNotUsedInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ProcedureNotUsedInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("UnassignedVariableUsageInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error,  CodeInspectionSeverity.Error),
                new CodeInspectionSetting("VariableNotUsedInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("VariableNotAssignedInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ImplicitByRefParameterInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("ImplicitVariantReturnTypeInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Hint,  CodeInspectionSeverity.Hint),
                new CodeInspectionSetting("MultipleDeclarationsInspection", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("ObsoleteCommentSyntaxInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion,  CodeInspectionSeverity.Suggestion),
                new CodeInspectionSetting("OptionExplicitInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error,  CodeInspectionSeverity.Error),
                new CodeInspectionSetting("VariableTypeNotDeclaredInspection", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Warning,  CodeInspectionSeverity.Warning),
                new CodeInspectionSetting("MalformedAnnotationInspection", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error,  CodeInspectionSeverity.Error)
            };
        }
    }
}
