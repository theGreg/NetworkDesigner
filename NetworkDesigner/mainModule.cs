public bool selectedShapeOK()
{
    ElementEnumerator oElEnum;
    Element oEl;
    int count;
    selectedShapeOK = false;
    ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 183
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
   Set oElEnum = ActiveModelReference.GetSelectedElements

 */
    oElEnum.Reset();
    while (oElEnum.MoveNext)
    {
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 290
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
      Set oEl = oElEnum.Current

 */
        if (matchParams(oEl, "Level 11", 7))
        {
            selectedShapeOK = true;
            break;
        }
    }
}
