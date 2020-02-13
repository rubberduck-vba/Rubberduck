namespace Rubberduck.Refactorings
{
    /// <summary>
    /// The heart of a refactoring: this part of the refactoring performs the actual transformation of the code once all necessary information has been gathered.
    /// </summary>
    /// <typeparam name="TModel">The model used by the refactoring containing all information needed to specify what to do.</typeparam>
    public interface IBaseRefactoring<in TModel> where TModel : class, IRefactoringModel
    {
        /// <summary>
        /// Performs the actual refactoring based on the parameters specified in the model.
        /// </summary>
        /// <param name="model">The model specifying all parameters of the refactoring</param>
        void Refactor(TModel model);
    }

    public interface IBaseRefactoringWithPreview<in TModel> : IBaseRefactoring<TModel>
        where TModel : class, IRefactoringModel
    {
        /// <summary>
        /// Returns some preview of the refactored code.
        /// </summary>
        /// <param name="model">The model used by the refactoring containing all information needed to specify what to do.</param>
        /// <returns>Preview of the refactored code</returns>
        string Preview(TModel model);
    }
}