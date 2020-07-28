using System;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Introduces a 'Property Get' member to make a write-only property read/write; Rubberduck will not infer the property's backing field, the body of the new member must be implemented manually.
    /// </summary>
    /// <inspections>
    /// <inspection name="WriteOnlyPropertyInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// Private internalValue As Long
    /// 
    /// Public Property Let Something(ByVal value As Long)
    ///     internalValue = value
    /// End Property
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// Private internalValue As Long
    /// 
    /// Public Property Get Something() As Long
    /// End Property
    /// 
    /// Public Property Let Something(ByVal value As Long)
    ///     internalValue = value
    /// End Property
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class IntroduceGetAccessorQuickFix : QuickFixBase
    {
        public IntroduceGetAccessorQuickFix()
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

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.IntroduceGetAccessorQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;

        private string GetParamText(ParameterDeclaration param)
        {
            return string.Format("{0} {1} As {2}",
                ((VBAParser.ArgContext)param.Context).BYVAL() == null ? "ByRef" : "ByVal",
                param.IdentifierName,
                param.AsTypeName);
        }
    }
}
