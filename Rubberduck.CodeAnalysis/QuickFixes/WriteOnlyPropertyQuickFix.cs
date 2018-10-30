using System;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class WriteOnlyPropertyQuickFix : QuickFixBase
    {
        public WriteOnlyPropertyQuickFix()
            : base(typeof(WriteOnlyPropertyInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var parameters = ((IParameterizedDeclaration) result.Target).Parameters.ToList();

            var signatureParams = parameters.Except(new[] {parameters.Last()}).Select(GetParamText);

            var propertyGet = string.Format("Public Property Get {0}({1}) As {2}{3}End Property{3}{3}",
                                            result.Target.IdentifierName,
                                            string.Join(", ", signatureParams),
                                            parameters.Last().AsTypeName,
                                            Environment.NewLine);

            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
            rewriter.InsertBefore(result.Target.Context.Start.TokenIndex, propertyGet);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.WriteOnlyPropertyQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;

        private string GetParamText(ParameterDeclaration param)
        {
            return string.Format("{0} {1} As {2}",
                ((VBAParser.ArgContext)param.Context).BYVAL() == null ? "ByRef" : "ByVal",
                param.IdentifierName,
                param.AsTypeName);
        }
    }
}
