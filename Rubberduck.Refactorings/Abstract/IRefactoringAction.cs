using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Refactorings
{
    /// <summary>
    /// The heart of a refactoring: this part of the refactoring performs the actual transformation of the code once all necessary information has been gathered.
    /// </summary>
    /// <typeparam name="TModel">The model used by the refactoring containing all information needed to specify what to do.</typeparam>
    public interface IRefactoringAction<in TModel> where TModel : class, IRefactoringModel
    {
        /// <summary>
        /// Performs the actual refactoring based on the parameters specified in the model.
        /// </summary>
        /// <param name="model">The model specifying all parameters of the refactoring</param>
        void Refactor(TModel model);
    }

    public interface IRefactoringPreviewProvider<in TModel>
        where TModel : class, IRefactoringModel
    {
        /// <summary>
        /// Returns some preview of the refactored code.
        /// </summary>
        /// <param name="model">The model used by the refactoring containing all information needed to specify what to do.</param>
        /// <returns>Preview of the refactored code</returns>
        string Preview(TModel model);
    }

    public interface ICodeOnlyRefactoringAction<in TModel> : IRefactoringAction<TModel>
        where TModel : class, IRefactoringModel
    {
        /// <summary>
        /// Performs the refactoring according to the model and using the provided rewrite session.
        /// </summary>
        /// <param name="model">The model specifying all parameters of the refactoring</param>
        /// <param name="rewriteSession">Rewrite session used to manipulate the code (Does not get executed.)</param>
        void Refactor(TModel model, IRewriteSession rewriteSession);
    }
}