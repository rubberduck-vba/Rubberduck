using Antlr4.Runtime;
using Microsoft.CSharp.RuntimeBinder;
using NLog;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace Rubberduck.Refactorings.MoveMember
{
    //public enum MoveEndpoints
    //{
    //    Undefined,
    //    StdToStd,
    //    ClassToStd,
    //    ClassToClass,
    //    StdToClass,
    //    FormToStd,
    //    FormToClass
    //};

    //public class MoveMemberResources
    //{
    //    public static string Class_Initialize => "Class_Initialize";
    //    public static string Class_Terminate => "Class_Terminate";
    //    public static string UserForm => "UserForm";
    //    public static string OptionExplicit => $"{Tokens.Option} {Tokens.Explicit}";

    //    public static string Caption => "MoveMember";
    //    public static string DefaultErrorMessageFormat => "Unable to Move Member: {0}";
    //    public static string InvalidMoveDefinition => "Incomplete Move definition: Code element(s) and destination module must be defined";
    //    public static string VBALanguageSpecificationViolation => "The defined Move would result in a VBA Language Specification violation and generate uncompilable code";
    //    public static string ApplicableStrategyNotFound => "Applicable move strategy not found";
    //    public static string Prefix_Variable => "xxx_";
    //    public static string Prefix_Parameter => "x";  //"arg_";
    //    public static string Prefix_ClassInstantiationProcedure => "X_"; // "Create__";
    //    public static string UnsupportedMoveExceptionFormat => "Unable to Move Member: {1}({0})";

    //    public static bool IsOrNamedLikeALifeCycleHandler(Declaration member)
    //        => member.IdentifierName.Equals(Class_Initialize) || member.IdentifierName.Equals(Class_Terminate);
    //}

    public class MoveMemberRefactoring : InteractiveRefactoringBase<IMoveMemberPresenter, MoveMemberModel>
    {
        private readonly IMessageBox _messageBox;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly RubberduckParserState _state;
        private readonly IParseManager _parseManager;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly ISelectionService _selectionService;

        private MoveMemberModel Model { set; get; } = null;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public MoveMemberRefactoring(
            IDeclarationFinderProvider declarationFinderProvider, 
            IParseManager parseManager, 
            IMessageBox messageBox, 
            IRefactoringPresenterFactory factory, 
            IRewritingManager rewritingManager, 
            ISelectionService selectionService,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IUiDispatcher uiDispatcher)
            : base(rewritingManager, selectionService, factory, uiDispatcher)
                  
        {
            _state = declarationFinderProvider as RubberduckParserState;
            _declarationFinderProvider = declarationFinderProvider;
            _parseManager = parseManager;
            _messageBox = messageBox;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _selectionService = selectionService;
        }

        protected override MoveMemberModel InitializeModel(Declaration target)
        {
            return InitializeModel(target.QualifiedSelection);
        }

        private MoveMemberModel InitializeModel(QualifiedSelection qSelection)
        {
            Model = new MoveMemberModel(_state, RewritingManager);

            var initialTargetDeclaration = AcquireSingleMemberSelection(qSelection, Model, _declarationFinderProvider);

            Model.DefineMove(initialTargetDeclaration);
            return Model;
        }

        protected override void RefactorImpl(MoveMemberModel model)
        {
            if (!model.IsValidMoveDefinition)
            {
                _messageBox?.Message(MoveMemberResources.InvalidMoveDefinition);
                return;
            }

            if (model.Strategy == null)
            {
                _messageBox?.Message(MoveMemberResources.ApplicableStrategyNotFound);
                return;
            }

            Model = model;

            if (!Model.CurrentScenario.CreatesNewModule)
            {
                SafeMoveMembers();
                return;
            }

            var suspendResult = _parseManager.OnSuspendParser(this, new[] { ParserState.Ready }, SafeMoveMembers);
            var suspendOutcome = suspendResult.Outcome;
            if (suspendOutcome != SuspensionOutcome.Completed)
            {
                _logger.Warn($"AddModule: {Model.CurrentScenario.DestinationContentProvider.ModuleName} failed.");
            }
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            return AcquireTargets(targetSelection).FirstOrDefault();
        }

        public IEnumerable<Declaration> AcquireTargets(QualifiedSelection? selection)
        {
            Model = Model ?? InitializeModel(selection.Value);
            return Model.CurrentScenario.SelectedDeclarations;
        }

        private Declaration AcquireSingleMemberSelection(QualifiedSelection? selection, MoveMemberModel model, IDeclarationFinderProvider declarationFinderProvider)
        {
            if (!selection.HasValue)
            {
                return null;
            }

            var selectionModule = declarationFinderProvider.DeclarationFinder.ModuleDeclaration(selection.Value.QualifiedName);
            if (selectionModule is null)
            {
                return null;
            }

            //var selected = declarationFinderProvider.DeclarationFinder.FindSelectedDeclaration(selection.Value);
            var selected = _selectedDeclarationProvider.SelectedDeclaration(selection.Value);
            if (selected.IsMember() 
                || selected.IsConstant() && !selected.IsLocalConstant()
                || selected.IsVariable() && !selected.IsLocalVariable())
            {
                return selected;
            }

            //if (MoveScenario.DeclarationCanBeAnalyzed(selected))
            //{
            //    return selected;
            //}

            if (selected.DeclarationType.HasFlag(DeclarationType.Parameter)
                //|| selected.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember)
                || selected.IsLocalVariable()
                || selected.IsLocalConstant())
            {
                return selected.ParentDeclaration;
            }

            selected = declarationFinderProvider.DeclarationFinder.Members(selectionModule)
                .Where(m => m.IsMember() || m.IsConstant() && !m.IsLocalConstant()) // MoveScenario.DeclarationCanBeAnalyzed(m))
                .OrderBy(m => m.Selection)
                .FirstOrDefault(member => IsCloseToSelection(member, selection.Value.Selection));

            return selected ?? model.DefaultMemberToMove(selectionModule.QualifiedModuleName);
        }

        private static bool IsCloseToSelection(Declaration member, Selection selection)
        {
            var context = member.Context;
            if (context is null)
            {
                return false;
            }

            if (context.Start.Line < selection.StartLine && context.Stop.Line > selection.EndLine)
            {
                return true;
            }

            if (member.DeclarationType.HasFlag(DeclarationType.Member))
            {
                return context.Start.Line == selection.StartLine
                     || context.Stop.Line == selection.EndLine;
            }

            //if (member.DeclarationType.HasFlag(DeclarationType.Variable))
            //{
            //    return ContainsSelection(context as VBAParser.VariableSubStmtContext, selection, context.Parent as VBAParser.VariableListStmtContext, context.Parent.Parent as VBAParser.VariableStmtContext);
            //}

            if (member.DeclarationType.HasFlag(DeclarationType.Constant))
            {
                return ContainsSelection(context as VBAParser.ConstSubStmtContext, selection, context.Parent as VBAParser.ConstStmtContext, context.Parent as VBAParser.ConstStmtContext);
            }
            return false;
        }

        private static bool ContainsSelection<TContext, TList, TListParent>(TContext context, Selection selection, TList list, TListParent listParent)
            where TContext : ParserRuleContext
            where TList : ParserRuleContext
            where TListParent : ParserRuleContext
        {
            var firstContext = list.children.FirstOrDefault(ch => ch is TContext);
            if (firstContext == context)
            {
                var firstContextSelection = new Selection(listParent.Start.Line, listParent.Start.Column + 1, context.Stop.Line, context.Stop.Column + context.Stop.Text.Length + 1);
                return firstContextSelection.Contains(selection);
            }

            var lastContext = list.children.LastOrDefault(ch => ch is TContext);
            if (lastContext == context)
            {
                return context.Stop.Line == selection.EndLine;
            }

            var varContexts = list.children.Where(ch => ch is TContext).Select(ch => ch as TContext);
            for (var idx = 1; idx < varContexts.Count() - 1; idx++)
            {
                if (varContexts.ElementAt(idx) == context && context != lastContext)
                {
                    var nextContext = varContexts.ElementAt(idx + 1);
                    var ctxtSelection = new Selection(context.Start.Line, context.Start.Column, nextContext.Start.Line, nextContext.Start.Column - 1);
                    return ctxtSelection.Contains(selection);
                }
            }
            return false;
        }

        private void SafeMoveMembers()
        {
            ICodeModule newlyCreatedCodeModule = null;
            var newModulePostMoveSelection = new Selection();
            try
            {
                if (Model.Strategy is null)
                {
                    return;
                }

                Model.Strategy.ModifyContent();

                if (!Model.MoveRewritingManager.MoveRewriteSession.TryRewrite())
                {
                    PresentMoveMemberErrorMessage(BuildDefaultErrorMessage(Model.CurrentScenario.SelectedDeclarations.FirstOrDefault()));
                    return;
                }

                if (Model.CurrentScenario.CreatesNewModule)
                {
                    //CreateNewModuleWithContent returns an ICodeModule reference to support setting the post-move Selection.
                    //Unable to use the ISelectionService after creating a module, since the
                    //new Component is apparently not available via VBComponents until after a reparse
                    newlyCreatedCodeModule = CreateNewModule(Model.Strategy.DestinationNewModuleContent, Model.CurrentScenario.MoveDefinition);
                    newModulePostMoveSelection = new Selection(newlyCreatedCodeModule.CountOfLines - Model.Strategy.DestinationNewContentLineCount + 1, 1);
                }
            }
            //TODO: Review these catches
            catch (MoveMemberUnsupportedMoveException unsupportedMove)
            {
                PresentMoveMemberErrorMessage(unsupportedMove.Message);
            }
            catch (RuntimeBinderException rbEx)
            {
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(Model.CurrentScenario.SelectedDeclarations.FirstOrDefault())}: {rbEx.Message}");
            }
            catch (COMException comEx)
            {
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(Model.CurrentScenario.SelectedDeclarations.FirstOrDefault())}: {comEx.Message}");
            }
            catch (ArgumentException argEx)
            {
                //This exception is often thrown when there is a rewrite conflict (e.g., try to insert where something's been deleted)
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(Model.CurrentScenario.SelectedDeclarations.FirstOrDefault())}: {argEx.Message}");
            }
            catch (Exception unhandledEx)
            {
                PresentMoveMemberErrorMessage($"{BuildDefaultErrorMessage(Model.CurrentScenario.SelectedDeclarations.FirstOrDefault())}: {unhandledEx.Message}");
            }
            finally
            {
                if (newlyCreatedCodeModule != null)
                {
                    using (newlyCreatedCodeModule)
                    {
                        SetPostMoveSelection(newModulePostMoveSelection, newlyCreatedCodeModule);
                    }
                }
                else
                {
                    SetPostMoveSelection();
                }
            }
        }

        private static ICodeModule CreateNewModule(string newModuleContent, MoveDefinition moveDefinition)
        {
            ICodeModule codeModule = null;
            var vbProject = moveDefinition.Source.Module.Project;
            using (var components = vbProject.VBComponents)
            {
                using (var newComponent = components.Add(moveDefinition.Destination.ComponentType))
                {
                    newComponent.Name = moveDefinition.Destination.ModuleName;
                    using (var newModule = newComponent.CodeModule)
                    {
                        //If VBE Option 'Require Variable Declaration' is set, then
                        //Option Explicit is included with a newly inserted Module...hence, the check
                        if (newModule.Content().Contains(MoveMemberResources.OptionExplicit))
                        {
                            newModule.InsertLines(newModule.CountOfLines, newModuleContent);
                        }
                        else
                        {
                            newModule.InsertLines(1, $"{MoveMemberResources.OptionExplicit}{Environment.NewLine}{Environment.NewLine}{newModuleContent}");
                        }
                        codeModule = newModule;
                    }
                }
            }
            return codeModule;
        }

        private void SetPostMoveSelection(Selection postMoveSelection = new Selection(), ICodeModule newlyCreatedCodeModule = null)
        {
            //The move/rewrite is done at this point, so do not bubble up any exceptions. 
            //If the user sees an exception, he may think that the the move failed
            try
            {
                if (newlyCreatedCodeModule != null)
                {
                    using (var codePane = newlyCreatedCodeModule.CodePane)
                    {
                        if (!codePane.IsWrappingNullReference)
                        {
                            codePane.Selection = postMoveSelection;
                        }
                    }
                    return;
                }

                var lastPreMoveDestinationMember = Model.CurrentScenario.DestinationContentProvider.ModuleDeclarations.Where(d => d.IsMember()).OrderBy(d => d.Selection).LastOrDefault();
                _selectionService.TrySetSelection(Model.CurrentScenario.MoveDefinition.Destination.Module.QualifiedModuleName, new Selection(lastPreMoveDestinationMember?.Context.Stop.Line ?? 1, 1));
            }
            catch (Exception ex)
            {
                _logger.Warn($"{ex.Message}: {nameof(SetPostMoveSelection)} threw and exception");
            }
        }

        private void PresentMoveMemberErrorMessage(string errorMsg)
        {
            _messageBox?.NotifyWarn(errorMsg, MoveMemberResources.Caption);
        }

        private string BuildDefaultErrorMessage(Declaration target)
        {
            return string.Format(MoveMemberResources.DefaultErrorMessageFormat, target?.IdentifierName ?? MoveMemberResources.InvalidMoveDefinition);
        }
    }

    //TODO: Are there any tests checking for this exception?
    [Serializable]
    class MoveMemberUnsupportedMoveException : Exception
    {
        public MoveMemberUnsupportedMoveException() { }

        public MoveMemberUnsupportedMoveException(Declaration declaration)
            : base(String.Format(MoveMemberResources.UnsupportedMoveExceptionFormat, ToLocalizedString(declaration.DeclarationType) , declaration.IdentifierName))
        { }

        private static string ToLocalizedString(DeclarationType type)
            => RubberduckUI.ResourceManager.GetString("DeclarationType_" + type, CultureInfo.CurrentUICulture);
    }



    ////These could/should(?) go into DeclarationExtensions
    //public static class MoveMemberExtensions
    //{
    //    public static bool IsVariable(this Declaration declaration)
    //        => declaration.DeclarationType.HasFlag(DeclarationType.Variable);

    //    public static bool IsLocalVariable(this Declaration declaration)
    //        => declaration.IsVariable() && declaration.ParentDeclaration.IsMember();

    //    public static bool IsLocalConstant(this Declaration declaration)
    //        => declaration.IsConstant() && declaration.ParentDeclaration.IsMember();

    //    public static bool HasPrivateAccessibility(this Declaration declaration)
    //        => declaration.Accessibility.Equals(Accessibility.Private);

    //    public static bool IsMember(this Declaration declaration)
    //        => declaration.DeclarationType.HasFlag(DeclarationType.Member);

    //    public static bool IsConstant(this Declaration declaration)
    //        => declaration.DeclarationType.HasFlag(DeclarationType.Constant);

    //    public static IEnumerable<IdentifierReference> AllReferences(this IEnumerable<Declaration> declarations)
    //    {
    //        return from dec in declarations
    //               from reference in dec.References
    //               select reference;
    //    }

        ///// <summary>
        ///// Returns a member's signature with an improved argument list.
        ///// </summary>
        //public static string ImprovedSignature(this ModuleBodyElementDeclaration declaration)
        //    => ImprovedSignature(declaration, declaration.Accessibility);

        ///// <summary>
        ///// Returns a member's signature with specified accessibility and an improved argument list.
        ///// </summary>
        //public static string ImprovedSignature(this ModuleBodyElementDeclaration declaration, Accessibility accessibility)
        //{
        //    var memberType = string.Empty;
        //    switch (declaration.Context)
        //    {
        //        case VBAParser.SubStmtContext _:
        //            memberType = Tokens.Sub;
        //            break;
        //        case VBAParser.FunctionStmtContext _:
        //            memberType = Tokens.Function;
        //            break;
        //        case VBAParser.PropertyGetStmtContext _:
        //            memberType = $"{Tokens.Property} {Tokens.Get}";
        //            break;
        //        case VBAParser.PropertyLetStmtContext _:
        //            memberType = $"{Tokens.Property} {Tokens.Let}";
        //            break;
        //        case VBAParser.PropertySetStmtContext _:
        //            memberType = $"{Tokens.Property} {Tokens.Set}";
        //            break;
        //        default:
        //            throw new ArgumentException();
        //    }

        //    var accessibilityToken = accessibility.Equals(Accessibility.Implicit) ? string.Empty : $"{accessibility.ToString()} ";

        //    var signature = $"{memberType} {declaration.IdentifierName}({declaration.ImprovedArgList()})";

        //    var fullSignature = $"{accessibilityToken}{signature}";
        //    if (declaration.AsTypeName != null)
        //    {
        //        fullSignature = $"{fullSignature} As {declaration.AsTypeName}";
        //    }
        //    return fullSignature;
        //}


        ////Modifies argument list as needed to conform with
        ////considerations identified in https://github.com/rubberduck-vba/Rubberduck/issues/3486
        ///// <summary>
        ///// Returns a member's argument list mproved with explicit ByRef/ByVal parameter modifiers
        ///// and Type identifiers.
        ///// </summary>
        //public static string ImprovedArgList(this ModuleBodyElementDeclaration declaration)
        //{
        //    var memberParams = new List<string>();
        //    if (declaration is IParameterizedDeclaration memberWithParams)
        //    {
        //        var paramAttributeTuples = memberWithParams.Parameters.TakeWhile(p => p != memberWithParams.Parameters.Last())
        //            .OrderBy(o => o.Selection.StartLine)
        //            .ThenBy(t => t.Selection.StartColumn)
        //            .Select(p => ExamineArgContext(p)).ToList();

        //        if (memberWithParams.Parameters.Any())
        //        {
        //            var lastParamRequiresByValModifier = declaration.DeclarationType.HasFlag(DeclarationType.PropertyLet)
        //                || declaration.DeclarationType.HasFlag(DeclarationType.PropertySet);
        //            paramAttributeTuples.Add(ExamineArgContext(memberWithParams.Parameters.Last(), lastParamRequiresByValModifier));
        //        }

        //        memberParams = paramAttributeTuples.Select(ps =>
        //            ParameterTupleToString(
        //                ps.Modifier,
        //                ps.Name,
        //                ps.Type)).ToList();
        //    }

        //    return string.Join(", ", memberParams);


        //    string ParameterTupleToString(string paramModifier, string paramName, string paramType)
        //        => paramModifier.Length > 0 ? $"{paramModifier} {paramName} As {paramType}" : $"{paramName} As {paramType}";

        //    (string Modifier, string Name, string Type) ExamineArgContext(ParameterDeclaration param, bool requiresByVal = false)
        //    {
        //        var arg = param.Context as VBAParser.ArgContext;
        //        var modifier = arg.BYVAL()?.ToString() ?? Tokens.ByRef;

        //        if (param.IsObject || arg.GetDescendent<VBAParser.CtLExprContext>() != null || requiresByVal)
        //        {
        //            modifier = Tokens.ByVal;
        //        }

        //        if (param.IsParamArray)
        //        {
        //            modifier = Tokens.ParamArray;
        //            return (string.Empty, arg.GetText().Replace($"{param.AsTypeName}", string.Empty).Trim(), param.AsTypeName);
        //        }
        //        return (modifier, param.IdentifierName, param.AsTypeName);
        //    }
        //}
    //}
}
