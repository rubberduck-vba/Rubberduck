using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberPropertiesData
    {
        public MoveMemberPropertiesData(IEnumerable<Declaration> elements, Declaration module)
        {
            Elements = elements;
            LoadVariableToReferencingProperties();
            var allProperties = VariableToReferencingPropertiesMap.Values.SelectMany(v => v);

            var resolvedPropertyCandidates = new List<Declaration>();
            var propertiesReferencingMultipleVariables = new HashSet<Declaration>();
            foreach (var property in allProperties)
            {
                if (resolvedPropertyCandidates.Contains(property))
                {
                    propertiesReferencingMultipleVariables.Add(property);
                }
                else
                {
                    resolvedPropertyCandidates.Add(property);
                }
            }
            resolvedPropertyCandidates = resolvedPropertyCandidates.Except(propertiesReferencingMultipleVariables).ToList();


            if (resolvedPropertyCandidates.Any())
            {
                foreach( var kv in VariableToReferencingPropertiesMap)
                {
                    if (kv.Value.All( p => resolvedPropertyCandidates.Contains(p)))
                    {
                        PropertyBackingVariablesResolved.Add(kv.Key, kv.Value);
                    }
                }
            }

            if (propertiesReferencingMultipleVariables.Any())
            {
                foreach (var key in VariableToReferencingPropertiesMap.Keys)
                {
                    var properties = VariableToReferencingPropertiesMap[key];
                    if (properties.All(p => propertiesReferencingMultipleVariables.Contains(p)))
                    {
                        PropertyBackingVariablesUnresolved.Add(key, properties);
                    }
                }
            }
        }

        private IEnumerable<Declaration> Elements { set; get; }

        private Dictionary<Declaration, List<Declaration>> VariableToReferencingPropertiesMap { set; get; } = new Dictionary<Declaration, List<Declaration>>();

        private Dictionary<Declaration, List<Declaration>> PropertyBackingVariablesResolved { set; get; } = new Dictionary<Declaration, List<Declaration>>();

        private Dictionary<Declaration, List<Declaration>> PropertyBackingVariablesUnresolved { set; get; } = new Dictionary<Declaration, List<Declaration>>();

        public bool IsPropertyBackingVariable(Declaration variable)
            => PropertyBackingVariablesResolved.ContainsKey(variable);

        public bool TryGetPropertiesFromBackingVariable(Declaration backingVariable, out List<Declaration> properties)
            => PropertyBackingVariablesResolved.TryGetValue(backingVariable, out properties);

        public bool TryGetBackingVariableFromPropertyIdentifier(string propertyIdentifier, out Declaration variable)
        {
            variable = null;
            foreach (var key in PropertyBackingVariablesResolved.Keys)
            {
                if (PropertyBackingVariablesResolved[key].First().IdentifierName.Equals(propertyIdentifier))
                {
                    variable = key;
                    return true;
                }
            }
            return false;
        }

        private void LoadVariableToReferencingProperties()
        {
            foreach (var variable in Elements.Where(m => m.IsVariable()))
            {
                var properties = new List<Declaration>();
                if (VariableToReferencingPropertiesMap.ContainsKey(variable))
                {
                    continue;
                }

                //TODO: This misses case where variable is initialized within Class_Initialization (for example)
                //TODO: Consider assignments where RHS is a constant value as still qualifying as a backing variable
                //TODO: Also consider assignments of the backing variable withing Class_Initialize() as still qualifying
                var letterSetter = variable.References.SingleOrDefault(rf => rf.IsAssignment
                    && (rf.ParentScoping.DeclarationType.HasFlag(DeclarationType.PropertyLet)
                    || rf.ParentScoping.DeclarationType.HasFlag(DeclarationType.PropertySet)))?.ParentScoping ?? null;

                if (letterSetter != null)
                {
                    properties.Add(letterSetter);
                    var associatedGetter = Elements.SingleOrDefault(ame => ame.DeclarationType.HasFlag(DeclarationType.PropertyGet) && ame.IdentifierName.Equals(letterSetter.IdentifierName));
                    if (associatedGetter != null)
                    {
                        var options = new string[] { $"{Tokens.Let} {associatedGetter.IdentifierName}", $"{Tokens.Set } {associatedGetter.IdentifierName }", associatedGetter.IdentifierName };
                        if (variable.References.Any(rf => rf.ParentScoping.Equals(associatedGetter)
                            &&  variable.AsTypeName.Equals(rf.ParentScoping.AsTypeName)))
                        {
                            var assignment = associatedGetter.References.SingleOrDefault(rf => rf.IsAssignment && rf.ParentScoping.Equals(associatedGetter));
                            if (assignment.Context.Parent is VBAParser.LetStmtContext letStmt)
                            {
                                if (!options.Any(opt => letStmt.GetText().StartsWith(opt))
                                    || !letStmt.GetText().EndsWith(variable.IdentifierName))
                                {
                                    continue;
                                }
                            }
                            properties.Add(associatedGetter);
                        }
                        else
                        {
                            //If there is an associated getter, then there 
                            //needs to be a reference to the variable in both Read/Write Properties to 
                            //qualify as 'the' backing variable
                            continue;
                        }
                    }
                }
                else
                {
                    var readonlyProperty = Elements.SingleOrDefault(ame => 
                        ame.DeclarationType.HasFlag(DeclarationType.PropertyGet) 
                        && variable.References.Any(rf => rf.ParentScoping.Equals(ame) 
                        && ame.AsTypeName.Equals(rf.ParentScoping.AsTypeName)));

                    if (readonlyProperty != null)
                    {
                        var readRefs = variable.References.Where(rf => rf.ParentScoping.Equals(readonlyProperty));
                        if (readRefs.Any())
                        {
                            var letStmtCtxts = readRefs.Where(rf => rf.Context.GetAncestor<VBAParser.LetStmtContext>() != null)
                                .Select(rf => rf.Context.GetAncestor<VBAParser.LetStmtContext>());

                            if (letStmtCtxts.Any() && readRefs.Last().Context.TryGetAncestor<VBAParser.LExprContext>(out var lExpr))
                            {
                                var refReadsAsLetStmtChildren = readRefs.Where(rf => letStmtCtxts.First().children.Contains(lExpr));
                                if (refReadsAsLetStmtChildren.Any() && !VariableToReferencingPropertiesMap.Values.SelectMany(v => v).Any(v => properties.Contains(v)))
                                {
                                    properties.Add(readonlyProperty);
                                }
                            }
                        }
                    }
                }

                if (properties.Any())
                {
                    VariableToReferencingPropertiesMap.Add(variable, properties);
                }
            }
        }
    }
}
