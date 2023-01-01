using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractedMethod : IExtractedMethod
    {
        public string MethodName { get => GetNewMethodName(); }
        public string NewMethodNameBase { get; set; }
        public Accessibility Accessibility { get; set; }
        public bool SetReturnValue { get; set; }
        public ExtractMethodParameter ReturnValue { get; set; }
        public IEnumerable<ExtractMethodParameter> Parameters { get; set; }
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public ExtractedMethod(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
            NewMethodNameBase = RefactoringsUI.ExtractMethod_DefaultNewMethodName;
        }
        public virtual string NewMethodCall()
        {
            //functionality inside ExtractMethodModel for now
            return string.Empty;
        }
        public string GetNewMethodName()
        {
            var newMethodName = NewMethodNameBase;

            var newMethodInc = 0;
            // iterate until we have a non-clashing method name.
            while (IsConflictingName(newMethodName))
            {
                newMethodInc++;
                newMethodName = NewMethodNameBase + newMethodInc;
            }
            return newMethodName;
        }

        public bool IsConflictingName(string methodName)
        {
            IEnumerable<Declaration> declarations = _declarationFinderProvider.DeclarationFinder.AllUserDeclarations;
            var existingName = declarations.FirstOrDefault(d =>
                        Enumerable.Contains(ProcedureTypes, d.DeclarationType)
                    && d.IdentifierName.Equals(methodName));
            return existingName != null;
        }

        public static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };
    }
}
