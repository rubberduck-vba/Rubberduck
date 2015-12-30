using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerDockablePresenter : DockableToolwindowPresenter
    {
        private readonly RubberduckParserState _state;
        private CodeExplorerWindow Control { get { return UserControl as CodeExplorerWindow; } }

        public CodeExplorerDockablePresenter(RubberduckParserState state, VBE vbe, AddIn addIn, CodeExplorerWindow view)
            : base(vbe, addIn, view)
        {
            _state = state;
        }

        //private string GetNodeText(Declaration declaration)
        //{
        //    if (Control.DisplayStyle == TreeViewDisplayStyle.MemberNames)
        //    {
        //        var result = declaration.IdentifierName;
        //        if (declaration.DeclarationType == DeclarationType.PropertyGet)
        //        {
        //            result += " (" + Tokens.Get + ")";
        //        }
        //        else if (declaration.DeclarationType == DeclarationType.PropertyLet)
        //        {
        //            result += " (" + Tokens.Let + ")";
        //        }
        //        else if (declaration.DeclarationType == DeclarationType.PropertySet)
        //        {
        //            result += " (" + Tokens.Set + ")";
        //        }

        //        return result;
        //    }

        //    if (declaration.DeclarationType == DeclarationType.Procedure)
        //    {
        //        return ((VBAParser.SubStmtContext) declaration.Context).Signature();
        //    }

        //    if (declaration.DeclarationType == DeclarationType.Function)
        //    {
        //        return ((VBAParser.FunctionStmtContext)declaration.Context).Signature();
        //    }

        //    if (declaration.DeclarationType == DeclarationType.PropertyGet)
        //    {
        //        return ((VBAParser.PropertyGetStmtContext)declaration.Context).Signature();
        //    }

        //    if (declaration.DeclarationType == DeclarationType.PropertyLet)
        //    {
        //        return ((VBAParser.PropertyLetStmtContext)declaration.Context).Signature();
        //    }

        //    if (declaration.DeclarationType == DeclarationType.PropertySet)
        //    {
        //        return ((VBAParser.PropertySetStmtContext)declaration.Context).Signature();
        //    }

        //    return declaration.IdentifierName;
        //}

        //private string GetImageKeyForDeclaration(Declaration declaration)
        //{
        //    var result = string.Empty;
        //    switch (declaration.DeclarationType)
        //    {
        //        case DeclarationType.Module:
        //            break;
        //        case DeclarationType.Class:
        //            break;
        //        case DeclarationType.Procedure:
        //        case DeclarationType.Function:
        //            if (declaration.Accessibility == Accessibility.Private)
        //            {
        //                result = "PrivateMethod";
        //                break;
        //            }
        //            if (declaration.Accessibility == Accessibility.Friend)
        //            {
        //                result = "FriendMethod";
        //                break;
        //            }
        //            result = "PublicMethod";
        //            break;

        //        case DeclarationType.PropertyGet:
        //        case DeclarationType.PropertyLet:
        //        case DeclarationType.PropertySet:
        //            if (declaration.Accessibility == Accessibility.Private)
        //            {
        //                result = "PrivateProperty";
        //                break;
        //            }
        //            if (declaration.Accessibility == Accessibility.Friend)
        //            {
        //                result = "FriendProperty";
        //                break;
        //            }
        //            result = "PublicProperty";
        //            break;

        //        case DeclarationType.Parameter:
        //            break;
        //        case DeclarationType.Variable:
        //            if (declaration.Accessibility == Accessibility.Private)
        //            {
        //                result = "PrivateField";
        //                break;
        //            }
        //            if (declaration.Accessibility == Accessibility.Friend)
        //            {
        //                result = "FriendField";
        //                break;
        //            }
        //            result = "PublicField";
        //            break;

        //        case DeclarationType.Constant:
        //            if (declaration.Accessibility == Accessibility.Private)
        //            {
        //                result = "PrivateConst";
        //                break;
        //            }
        //            if (declaration.Accessibility == Accessibility.Friend)
        //            {
        //                result = "FriendConst";
        //                break;
        //            }
        //            result = "PublicConst";
        //            break;

        //        case DeclarationType.Enumeration:
        //            if (declaration.Accessibility == Accessibility.Private)
        //            {
        //                result = "PrivateEnum";
        //                break;
        //            }
        //            if (declaration.Accessibility == Accessibility.Friend)
        //            {
        //                result = "FriendEnum";
        //                break;
        //            }
        //            result = "PublicEnum";
        //            break;

        //        case DeclarationType.EnumerationMember:
        //            result = "EnumItem";
        //            break;

        //        case DeclarationType.Event:
        //            if (declaration.Accessibility == Accessibility.Private)
        //            {
        //                result = "PrivateEvent";
        //                break;
        //            }
        //            if (declaration.Accessibility == Accessibility.Friend)
        //            {
        //                result = "FriendEvent";
        //                break;
        //            }
        //            result = "PublicEvent";
        //            break;

        //        case DeclarationType.UserDefinedType:
        //            if (declaration.Accessibility == Accessibility.Private)
        //            {
        //                result = "PrivateType";
        //                break;
        //            }
        //            if (declaration.Accessibility == Accessibility.Friend)
        //            {
        //                result = "FriendType";
        //                break;
        //            }
        //            result = "PublicType";
        //            break;

        //        case DeclarationType.UserDefinedTypeMember:
        //            result = "PublicField";
        //            break;
                
        //        case DeclarationType.LibraryProcedure:
        //        case DeclarationType.LibraryFunction:
        //            result = "Identifier";
        //            break;

        //        default:
        //            throw new ArgumentOutOfRangeException();
        //    }

        //    return result;
        //}
    }
}
