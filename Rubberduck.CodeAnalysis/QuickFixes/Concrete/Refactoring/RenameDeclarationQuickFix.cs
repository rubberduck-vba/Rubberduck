using System.Globalization;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Resources;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete.Refactoring
{
    /// <summary>
    /// Prompts for a new name, renames a declaration accordingly, and updates all usages.
    /// </summary>
    /// <inspections>
    /// <inspection name="HungarianNotationInspection" />
    /// <inspection name="UseMeaningfulNameInspection" />
    /// <inspection name="DefaultProjectNameInspection" />
    /// <inspection name="UnderscoreInPublicClassModuleMemberInspection" />
    /// <inspection name="ExcelUdfNameIsValidCellReferenceInspection" />
    /// </inspections>
    /// <canfix multiple="false" procedure="false" module="false" project="false" all="false" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     A1
    /// End Sub
    /// 
    /// Public Sub A1(ByVal value As Long)
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Renamed
    /// End Sub
    /// 
    /// Public Sub Renamed(ByVal value As Long)
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RenameDeclarationQuickFix : RefactoringQuickFixBase
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
            return string.Format(Resources.Inspections.QuickFixes.RenameDeclarationQuickFix,
                RubberduckUI.ResourceManager.GetString("DeclarationType_" + result.Target.DeclarationType, CultureInfo.CurrentUICulture));
        }
    }
}