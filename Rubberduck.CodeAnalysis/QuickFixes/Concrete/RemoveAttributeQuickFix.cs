using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Removes a hidden attribute, in order to maintain consistency between hidden attributes and (missing) annotation comments.
    /// </summary>
    /// <inspections>
    /// <inspection name="MissingModuleAnnotationInspection" />
    /// <inspection name="MissingMemberAnnotationInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="false" project="false" all="false" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    /// Attribute VB_UserMemId = 0
    /// 
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    /// 
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal class RemoveAttributeQuickFix : QuickFixBase
    {
        private readonly IAttributesUpdater _attributesUpdater;

        public RemoveAttributeQuickFix(IAttributesUpdater attributesUpdater)
        :base(typeof(MissingModuleAnnotationInspection), typeof(MissingMemberAnnotationInspection))
        {
            _attributesUpdater = attributesUpdater;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (!(result is IWithInspectionResultProperties<(string AttributeName, IReadOnlyList<string> AttributeValues)> resultProperties))
            {
                return;
            }

            var declaration = result.Target;
            var (attributeBaseName, attributeValues) = resultProperties.Properties;

            var attributeName = declaration.DeclarationType.HasFlag(DeclarationType.Module)
                ? attributeBaseName
                : Attributes.MemberAttributeName(attributeBaseName,declaration.IdentifierName);

            _attributesUpdater.RemoveAttribute(rewriteSession, declaration, attributeName, attributeValues);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveAttributeQuickFix;

        public override CodeKind TargetCodeKind => CodeKind.AttributesCode;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
        public override bool CanFixAll => false;
    }
}