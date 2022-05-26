using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public interface IDeleteDeclarationsModel : IRefactoringModel
    {
        IEnumerable<Declaration> Targets { get; }

        /// <summary>
        /// When set to true, this flag overides all other flags and prevents changes to Annotations and Comments.  
        /// Default value is false.
        /// </summary>
        bool DeleteDeclarationsOnly { set; get; }

        /// <summary>
        /// when set to true, this flag enables edits to comments adjacent to delete targets.  Default value is true.
        /// </summary>
        bool InsertValidationTODOForRetainedComments { set;  get; }

        /// <summary>
        /// When set to true, this flag enables the deletion of comments found on the same logical line as a 
        /// deleted Declaration.  Default value is true.
        /// </summary>
        bool DeleteDeclarationLogicalLineComments { set;  get; }

        /// <summary>
        /// When set to true, this flag enables the removal of Annotations exclusively associated with a set of deleted Declarations.  
        /// Default value is true.
        /// </summary>
        bool DeleteAnnotations { set;  get; }
    }

    public class DeleteDeclarationsModel : IDeleteDeclarationsModel
    {
        private readonly HashSet<Declaration> _targets = new HashSet<Declaration>();

        public DeleteDeclarationsModel()
        {}

        public DeleteDeclarationsModel(params Declaration[] targets)
        {
            AddRangeOfDeclarationsToDelete(targets);
        }

        public DeleteDeclarationsModel(IEnumerable<Declaration> targets)
        {
            AddRangeOfDeclarationsToDelete(targets);
        }

        public void AddDeclarationsToDelete(params Declaration[] targets)
        {
            AddRangeOfDeclarationsToDelete(targets);
        }

        public void AddRangeOfDeclarationsToDelete(IEnumerable<Declaration> targets)
        {
            foreach (var t in targets)
            {
                _targets.Add(t);
            }
        }

        public IEnumerable<Declaration> Targets => new List<Declaration>(_targets);

        public bool DeleteDeclarationsOnly { set; get; } = false;

        public bool InsertValidationTODOForRetainedComments { set; get; } = true;

        public bool DeleteDeclarationLogicalLineComments { set;  get; } = true;

        public bool DeleteAnnotations { set; get; } = true;
    }
}
