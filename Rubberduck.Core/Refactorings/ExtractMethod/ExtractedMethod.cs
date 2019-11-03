using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractedMethod : IExtractedMethod
    {
        private const string NEW_METHOD = "NewMethod";

        public string MethodName { get; set; }
        public Accessibility Accessibility { get; set; }
        public bool SetReturnValue { get; set; }
        public ExtractedParameter ReturnValue { get; set; }
        public IEnumerable<ExtractedParameter> Parameters { get; set; }

        public virtual string NewMethodCall()
        {
            if (String.IsNullOrWhiteSpace(MethodName))
            {
                MethodName = NEW_METHOD;
            }
            string result = "" + MethodName;
            string argList;
            if (Parameters.Any())
            {
                argList = String.Join(", ", Parameters.Select(p => p.Name));
                result += " " + argList;
            }
            return result;
        }
        public string getNewMethodName(IEnumerable<Declaration> declarations)
        {
            var newMethodName = NEW_METHOD;

            var newMethodInc = 0;
            // iterate until we have a non-clashing method name.
            while (isConflictingName(declarations, newMethodName))
            {
                newMethodInc++;
                newMethodName = NEW_METHOD + newMethodInc;
            }
            return newMethodName;
        }

        public bool isConflictingName(IEnumerable<Declaration> declarations, string methodName)
        {
            var existingName = declarations.FirstOrDefault(d =>
                        Enumerable.Contains(ProcedureTypes, d.DeclarationType)
                    && d.IdentifierName.Equals(methodName));
            return (existingName != null);
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
