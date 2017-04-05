using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ChangeDimToPrivateQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type> { typeof(ModuleScopeDimKeywordInspection) };
        public static IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public static void AddSupportedInspectionType(Type inspectionType)
        {
            if (!inspectionType.GetInterfaces().Contains(typeof(IInspection)))
            {
                throw new ArgumentException("Type must implement IInspection", nameof(inspectionType));
            }

            _supportedInspections.Add(inspectionType);
        }

        public void Fix(IInspectionResult result)
        {
            var module = result.QualifiedSelection.QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            var context = (VBAParser.VariableStmtContext)result.Target.Context.Parent.Parent;
            var newInstruction = Tokens.Private + " ";
            for (var i = 1; i < context.ChildCount; i++)
            {
                // index 0 would be the 'Dim' keyword
                newInstruction += context.GetChild(i).GetText();
            }

            var selection = context.GetSelection();
            var oldContent = module.GetLines(selection);
            var newContent = oldContent.Replace(context.GetText(), newInstruction);

            module.DeleteLines(selection);
            module.InsertLines(selection.StartLine, newContent);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ChangeDimToPrivateQuickFix;
        }

        public bool CanFixInProcedure { get; } = false;
        public bool CanFixInModule { get; } = true;
        public bool CanFixInProject { get; } = true;
    }
}