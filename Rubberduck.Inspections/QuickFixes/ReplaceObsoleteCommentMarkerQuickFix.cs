using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ReplaceObsoleteCommentMarkerQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ObsoleteCommentSyntaxInspection)
        };

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

            if (module.IsWrappingNullReference)
            {
                return;
            }
            var comment = result.Context.GetText();
            var start = result.Context.Start.Line;           
            var commentLine = module.GetLines(start, 1);
            var newComment = commentLine.Substring(0, result.Context.Start.Column) +
                             Tokens.CommentMarker +
                             comment.Substring(Tokens.Rem.Length, comment.Length - Tokens.Rem.Length);

            var lines = result.QualifiedSelection.Selection.LineCount;
            if (lines > 1)
            {
                module.DeleteLines(start + 1, lines - 1);
            }
            module.ReplaceLine(start, newComment);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveObsoleteStatementQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}