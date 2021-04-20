using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.ReferenceManagement.CompilationPasses;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    /// <summary>
    /// An abstraction responsible for getting the SuperType names for a document module.
    /// </summary>
    public interface IDocumentModuleSuperTypeNamesProvider
    {
        IEnumerable<string> GetSuperTypeNamesFor(DocumentModuleDeclaration document);
    }

    /// <summary>
    /// Gets the SuperType names for a document module using IComObject.
    /// </summary>
    public class DocumentModuleSuperTypeNamesProvider : IDocumentModuleSuperTypeNamesProvider
    {
        private readonly IUserComProjectProvider _userComProjectProvider;

        public DocumentModuleSuperTypeNamesProvider(IUserComProjectProvider userComProjectProvider)
        {
            _userComProjectProvider = userComProjectProvider;
        }

        // skip IDispatch.. just about everything implements it and RD doesn't need to care about it; don't care about IUnknown either
        private static readonly HashSet<string> IgnoredComInterfaces = new HashSet<string>(new[] { "IDispatch", "IUnknown" });

        public IEnumerable<string> GetSuperTypeNamesFor(DocumentModuleDeclaration document)
        {
            var userComProject = _userComProjectProvider.UserProject(document.ProjectId);
            if (userComProject == null)
            {
                return Enumerable.Empty<string>();
            }

            var comModule = userComProject.Members.SingleOrDefault(m => m.Name == document.ComponentName);
            if (comModule == null)
            {
                return Enumerable.Empty<string>();
            }

            var inheritedInterfaces = comModule is ComCoClass documentCoClass
                ? documentCoClass.ImplementedInterfaces.ToList()
                : (comModule as ComInterface)?.InheritedInterfaces.ToList();

            if (inheritedInterfaces == null)
            {
                return Enumerable.Empty<string>();
            }

            var relevantInterfaces = inheritedInterfaces
                .Where(i => !i.IsRestricted && !IgnoredComInterfaces.Contains(i.Name))
                .ToList();

            //todo: Find a way to deal with the VBE's document type assignment and interface behaviour not relying on an assumption about an interface naming conventions. 

            //Some hosts like Access chose to have a separate hidden interface for each document module and only let that inherit the built-in base interface.
            //Since we do not have a declaration for the hidden interface, we have to go one more step up the hierarchy.
            var additionalInterfaces = relevantInterfaces
                .Where(i => i.Name.Equals("_" + comModule.Name))
                .SelectMany(i => i.InheritedInterfaces)
                .ToList();

            relevantInterfaces.AddRange(additionalInterfaces);

            var superTypeNames = relevantInterfaces
                .Select(i => i.Name)
                .ToList();

            //This emulates the VBE's behaviour to allow assignment to the coclass type instead on the interface.
            var additionalSuperTypeNames = superTypeNames
                .Where(name => name.StartsWith("_"))
                .Select(name => name.Substring(1))
                .Where(name => !name.Equals(comModule.Name))
                .ToList();

            superTypeNames.AddRange(additionalSuperTypeNames);
            return superTypeNames.Distinct();
        }
    }
}
