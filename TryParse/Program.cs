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

        readonly ListVisitor _visitor = new ListVisitor();

        public override SyntaxList<SyntaxNode> VisitList<SyntaxNode>(SyntaxList<SyntaxNode> nodes)
        {
            var items = new List<SyntaxNode>(nodes.Count);
            foreach (var n in nodes)
            {
                var lst = _visitor.Visit(n);
                if (lst != null)
                    foreach (var r in lst)
                        items.Add((SyntaxNode)r);
            }
            if (items.Count == 0)
                return new SyntaxList<SyntaxNode>();
            return new SyntaxList<SyntaxNode>(items);
        }

        class ListVisitor : VisualBasicSyntaxVisitor<IEnumerable<SyntaxNode>>
        {
            enum State { SkipIntro = 0, GatherMembers, InProp, PropEnded }

            string className;
            List<SyntaxTrivia> gatheredTrivias = new List<SyntaxTrivia>();
            List<StatementSyntax> gatheredMembers = new List<StatementSyntax>();
            State state = State.SkipIntro;

            public override IEnumerable<SyntaxNode> DefaultVisit(SyntaxNode node)
            {
                var repl = base.DefaultVisit(node);
                if (state != State.InProp)
                    return repl;
                if (repl == null)
                    gatheredPropStatements.Add(node);
                else
                    gatheredPropStatements.AddRange(repl);
                return Enumerable.Empty<SyntaxNode>();
            }

            IEnumerable<SyntaxNode> Skip(SyntaxNode node)
            {
                if (node.HasLeadingTrivia)
                    gatheredTrivias.AddRange(node.GetLeadingTrivia());
                if (node.HasTrailingTrivia)
                    gatheredTrivias.AddRange(node.GetTrailingTrivia());
                return Enumerable.Empty<SyntaxNode>();
            }

            public override IEnumerable<SyntaxNode> VisitStopOrEndStatement(StopOrEndStatementSyntax node)
                => (state == State.SkipIntro) ? Skip(node) : base.VisitStopOrEndStatement(node);

            public override IEnumerable<SyntaxNode> VisitOptionStatement(OptionStatementSyntax node)
                => (state == State.SkipIntro) ? Skip(node) : base.VisitOptionStatement(node);

            public override IEnumerable<SyntaxNode> VisitExpressionStatement(ExpressionStatementSyntax node)
                => (state == State.SkipIntro) ? Skip(node) : base.VisitExpressionStatement(node);

            public override IEnumerable<SyntaxNode> VisitAssignmentStatement(AssignmentStatementSyntax node)
                => (state == State.SkipIntro) ? Skip(node) : base.VisitAssignmentStatement(node);

            enum AccessorKind { Get, Set };

            IEnumerable<SyntaxNode> EndProp()
            {
                var gps = gatheredPropStatements;
                state = State.GatherMembers;
                var prop = gps.OfType<PropertyStatementSyntax>()
                    .First(p => p.KindOfVBAPropStat() != SyntaxKind.None);
                // Property
                yield return prop
                    .WithIdentifier(SyntaxFactory.Identifier(prevPropName))
                    .WithParameterList(SyntaxFactory.ParameterList());
                // Accessors
                SyntaxKind keyword = SyntaxKind.None;
                EndBlockStatementSyntax lastEBS = null;
                for (int i = 0; i < gps.Count; i++)
                {
                    switch (keyword)
                    {
                        case SyntaxKind.None:
                            var pss = gps[i] as PropertyStatementSyntax;
                            if (pss == null)
                                break;
                            //*** BEGIN of property accessor
                            keyword = pss.KindOfVBAPropStat();
                            var paramz = (keyword == SyntaxKind.GetKeyword) ? default(ParameterListSyntax) : pss.ParameterList;
                            yield return SyntaxFactory
                                .GetAccessorStatement(pss.AttributeLists, SyntaxFactory.TokenList(), SyntaxFactory.Token(keyword), paramz)
                                .WithTrailingTrivia(pss.GetTrailingTrivia());
                            continue;
                        case SyntaxKind.GetKeyword:
                        case SyntaxKind.SetKeyword:
                            var ebs = gps[i] as EndBlockStatementSyntax;
                            if (ebs == null || ebs.BlockKeyword.Text != "Property")
                                break;
                            //*** END of property accessor
                            lastEBS = ebs;
                            //yield return ebs.WithEndKeyword(SyntaxFactory.Token(keyword));
                            var endKeyword = SyntaxFactory.Token(SyntaxKind.EndKeyword).WithTriviaFrom(ebs.EndKeyword);
                            var blockKeyword = SyntaxFactory.Token(keyword).WithTriviaFrom(ebs.BlockKeyword);
                            var res = (keyword == SyntaxKind.GetKeyword)
                                ? SyntaxFactory.EndGetStatement(endKeyword, blockKeyword)
                                : SyntaxFactory.EndSetStatement(endKeyword, blockKeyword);
                            yield return res;
                            keyword = SyntaxKind.None;
                            continue;
                    }
                    yield return gps[i];
                }
                // End Property

                yield return SyntaxFactory.EndPropertyStatement(
                    SyntaxFactory.Token(SyntaxKind.EndKeyword).WithTrailingTrivia(SyntaxFactory.Space),
                    SyntaxFactory.Token(SyntaxKind.PropertyKeyword).WithTriviaFrom(lastEBS.BlockKeyword)
                    );
                //
                gatheredPropStatements.Clear();
                prevPropName = null;
            }

            public override IEnumerable<SyntaxNode> VisitFieldDeclaration(FieldDeclarationSyntax node)
            {
                switch (state)
                {
                    case State.SkipIntro:
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
                                yield break;
                            }
                        }
                        state = State.GatherMembers;
                        break;
                    case State.InProp:
                        foreach (var n in base.VisitFieldDeclaration(node))
                            yield return n;
                        yield break;
                    case State.PropEnded:
                        foreach (var sn in EndProp())
                            yield return sn;
                        state = State.GatherMembers;
                        break;
                }
                var repl = node;
                if (gatheredTrivias.Count > 0)
                {
                    if (repl.HasLeadingTrivia)
                        gatheredTrivias.AddRange(repl.GetLeadingTrivia());
                    repl = repl.WithLeadingTrivia(gatheredTrivias);
                    gatheredTrivias.Clear();
                }
                yield return repl;
            }

            string prevPropName;
            List<SyntaxNode> gatheredPropStatements = new List<SyntaxNode>();

            public override IEnumerable<SyntaxNode> VisitPropertyStatement(PropertyStatementSyntax node)
            {
                System.Diagnostics.Trace.Assert(node.KindOfVBAPropStat() != SyntaxKind.None);

                var propName = node.Identifier.TrailingTrivia.ElementAtOrDefault(1).ToString();
                if (string.IsNullOrWhiteSpace(propName))
                    return base.VisitPropertyStatement(node);

                var fromPrevProp = (prevPropName != null && prevPropName != propName) ? EndProp().ToList() : null;

                state = State.InProp;
                prevPropName = propName;
                gatheredPropStatements.Add(node);

                return fromPrevProp ?? Enumerable.Empty<SyntaxNode>();
            }

            public override IEnumerable<SyntaxNode> VisitEndBlockStatement(EndBlockStatementSyntax node)
            {
                if (state == State.InProp)
                {
                    if (node.BlockKeyword.Text == "Property")
                        state = State.PropEnded;
                    gatheredPropStatements.Add(node);
                    return Enumerable.Empty<SyntaxNode>();
                }
                else return base.VisitEndBlockStatement(node);

            }

            public override IEnumerable<SyntaxNode> VisitMethodBlock(MethodBlockSyntax node)
            {
                if (state == State.InProp)
                    EndProp();
                return base.VisitMethodBlock(node);
            }

            public override IEnumerable<SyntaxNode> VisitInvocationExpression(InvocationExpressionSyntax node)
            {
                var args = node.ArgumentList
                    .WithOpenParenToken(SyntaxFactory.Token(SyntaxKind.OpenBraceToken))
                    .WithCloseParenToken(SyntaxFactory.Token(SyntaxKind.CloseBraceToken));
                return base.VisitInvocationExpression(node.WithArgumentList(args));
            }

        }
    }

    static class SyntaxExts
    {
        public static SyntaxKind KindOfVBAPropStat(this PropertyStatementSyntax pss)
        {
            if (pss == null)
                return SyntaxKind.None;
            switch (pss.Identifier.Text)
            {
                case "Get": return SyntaxKind.GetKeyword;
                case "Let": return SyntaxKind.SetKeyword;
                case "Set": return SyntaxKind.SetKeyword;
                default: return SyntaxKind.None;
            }
        }
    }
}
