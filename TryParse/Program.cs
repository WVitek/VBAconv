using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace TryParse
{
    class Program
    {
        static void Main(string[] args)
        {
            var text = File.ReadAllText(args[0], Encoding.GetEncoding(1251));
            var tree = VisualBasicSyntaxTree.ParseText(text);
            var rootNode = tree.GetRoot();
            var myRw = new MyRewriter();
            var myRoot = myRw.Visit(rootNode);
            var dst = (args.Length > 1) ? args[1] : "SyntaxTree.log";
            File.WriteAllText(dst, myRoot.ToFullString(), Encoding.UTF8);
        }
    }

    class MyRewriter : VisualBasicSyntaxRewriter
    {
        enum State { SkipIntro = 0, GatherMembers, InProp }

        string className;
        List<SyntaxTrivia> gatheredTrivias = new List<SyntaxTrivia>();
        List<StatementSyntax> gatheredMembers = new List<StatementSyntax>();
        State state = State.SkipIntro;

        //public override SyntaxNode Visit(SyntaxNode node)
        //{
        //    var repl = base.Visit(node);
        //}

        public override SyntaxNode VisitCompilationUnit(CompilationUnitSyntax node)
        {
            var unit = (CompilationUnitSyntax)base.VisitCompilationUnit(node);
            if (className != null)
            {
                var cls = SyntaxFactory.ClassStatement(className);
                cls = cls.AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword));
                var blk = SyntaxFactory.ClassBlock(cls)
                    .NormalizeWhitespace("  ", false)
                    .WithMembers(unit.Members)
                    ;
                unit = unit.WithMembers(new SyntaxList<StatementSyntax>(blk));
            }
            return unit;
        }

        SyntaxNode Skip1(SyntaxNode node)
        {
            if (state != State.SkipIntro)
                return node;
            if (node.HasLeadingTrivia)
                gatheredTrivias.AddRange(node.GetLeadingTrivia());
            if (node.HasTrailingTrivia)
                gatheredTrivias.AddRange(node.GetTrailingTrivia());
            return null;
        }

        public override SyntaxNode VisitExpressionStatement(ExpressionStatementSyntax node) => Skip1(node);
        public override SyntaxNode VisitAssignmentStatement(AssignmentStatementSyntax node) => null;
        public override SyntaxNode VisitStopOrEndStatement(StopOrEndStatementSyntax node) => Skip1(node);
        public override SyntaxNode VisitOptionStatement(OptionStatementSyntax node) => Skip1(node);

        public override SyntaxNode VisitFieldDeclaration(FieldDeclarationSyntax node)
        {
            if (state == State.SkipIntro)
            {
                foreach (var decl in node.Declarators)
                {
                    if (decl.AsClause != null)
                        break;
                    if (decl.Initializer == null)
                        continue;
                    var name = Convert.ToString(decl.Names.FirstOrDefault());
                    if (name.StartsWith("VB_"))
                    {
                        if (name == "VB_Name")
                            className = Convert.ToString(decl.Initializer.Value).Trim('"');
                        return null;
                    }
                }
                state = State.GatherMembers;
            }
            else if (state == State.InProp)
                EndProp();

            var repl = base.VisitFieldDeclaration(node);
            if (gatheredTrivias.Count > 0)
            {
                if (repl.HasLeadingTrivia)
                    gatheredTrivias.AddRange(repl.GetLeadingTrivia());
                repl = repl.WithLeadingTrivia(gatheredTrivias);
                gatheredTrivias.Clear();
            }
            return repl;
        }

        void EndProp()
        {
            state = State.GatherMembers;
            //throw new NotImplementedException();
        }

        string prevPropName;
        List<SyntaxNode> gatheredPropStatements = new List<SyntaxNode>();

        public override SyntaxNode VisitPropertyStatement(PropertyStatementSyntax node)
        {
            var propAccessor = node.Identifier.Text;
            switch (node.Identifier.Text)
            {
                case "Get":
                case "Let":
                case "Set":
                    break;
                default:
                    return base.VisitPropertyStatement(node);
            }
            var propName = node.Identifier.TrailingTrivia.ElementAtOrDefault(1).ToString();
            if (string.IsNullOrWhiteSpace(propName))
                return base.VisitPropertyStatement(node);

            if (prevPropName != propName)
                EndProp();

            state = State.InProp;
            prevPropName = propName;
            gatheredPropStatements.Add(node);
            return null;
        }

        public override SyntaxNode VisitEndBlockStatement(EndBlockStatementSyntax node)
        {
            var repl = base.VisitEndBlockStatement(node);
            if (state != State.InProp)
                return repl;
            gatheredPropStatements.Add(repl);
            return null;
        }

        public override SyntaxNode VisitMethodBlock(MethodBlockSyntax node)
        {
            if (state == State.InProp)
                EndProp();
            return base.VisitMethodBlock(node);
        }

        public override SyntaxNode VisitEmptyStatement(EmptyStatementSyntax node)
        {
            var repl = base.VisitEmptyStatement(node);
            if (state != State.InProp)
                return repl;
            gatheredPropStatements.Add(repl);
            return null;
        }
    }
}
