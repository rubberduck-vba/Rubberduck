using System.Globalization;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RenameDeclarationQuickFix : RefactoringQuickFixBase
    {
        public RenameDeclarationQuickFix(RenameRefactoring refactoring)
            : base(refactoring,
                typeof(HungarianNotationInspection), 
                typeof(UseMeaningfulNameInspection),
                typeof(DefaultProjectNameInspection), 
                typeof(UnderscoreInPublicClassModuleMemberInspection),
                typeof(ExcelUdfNameIsValidCellReferenceInspection))
        {}

        protected override void Refactor(IInspectionResult result)
        {
            Refactoring.Refactor(result.Target);
        }

        public override string Description(IInspectionResult result)
        {
            return string.Format(RubberduckUI.Rename_DeclarationType,
                RubberduckUI.ResourceManager.GetString("DeclarationType_" + result.Target.DeclarationType,
                    CultureInfo.CurrentUICulture));
        }
    }
}