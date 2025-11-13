using Antlr4.Runtime;
using System.IO;

namespace SheetsApp
{
    public class ThrowingErrorListener : BaseErrorListener, IAntlrErrorListener<int>
    {
        public static readonly ThrowingErrorListener Instance = new ThrowingErrorListener();

        public override void SyntaxError(TextWriter output, IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            throw new Exception($"Помилка синтаксису (парсер): {msg}");
        }

        public void SyntaxError(TextWriter output, IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            throw new Exception($"Помилка синтаксису (лексер): {msg}");
        }
    }
    public class CycleException : Exception
    {
        public CycleException(string message) : base(message) { }
    }
    public class CellNotFoundException : Exception
    {
        public CellNotFoundException(string message) : base(message) { }
    }
    public class SheetsAppVisitor : SheetsBaseVisitor<double>
    {
        public bool IsLastResultLogical { get; private set; }
        private List<Cell> cells;
        private Cell currentCell = null;
        public SheetsAppVisitor(Dictionary<(int, int), Cell> cells)
        {
            this.cells = new List<Cell>(cells.Values);
        }
        public double Eval(Cell cell)
        {
            currentCell = cell;
            IsLastResultLogical = false;

            var inputStream = new AntlrInputStream(cell.Expression);
            var lexer = new SheetsLexer(inputStream);
            lexer.RemoveErrorListeners();
            lexer.AddErrorListener(ThrowingErrorListener.Instance);
            var tokens = new CommonTokenStream(lexer);
            var parser = new SheetsParser(tokens);
            parser.RemoveErrorListeners();
            parser.AddErrorListener(ThrowingErrorListener.Instance);
            var tree = parser.expression();
            return Visit(tree);
        }

        public override double VisitMultiplyExpr(SheetsParser.MultiplyExprContext context)
        {
            var left = Visit(context.expression(0));
            var right = Visit(context.expression(1));
            IsLastResultLogical = false;
            return left * right;
        }

        public override double VisitDivideExpr(SheetsParser.DivideExprContext context)
        {
            var left = Visit(context.expression(0));
            var right = Visit(context.expression(1));
            IsLastResultLogical = false;
            return left / right;
        }

        public override double VisitAddExpr(SheetsParser.AddExprContext context)
        {
            var left = Visit(context.expression(0));
            var right = Visit(context.expression(1));
            IsLastResultLogical = false;
            return left + right;
        }

        public override double VisitSubtractExpr(SheetsParser.SubtractExprContext context)
        {
            var left = Visit(context.expression(0));
            var right = Visit(context.expression(1));
            IsLastResultLogical = false;
            return left - right;
        }

        public override double VisitDecrementExpr(SheetsParser.DecrementExprContext context)
        {
            var value = Visit(context.expression());
            IsLastResultLogical = false;
            return value - 1;
        }

        public override double VisitIncrementExpr(SheetsParser.IncrementExprContext context)
        {
            var value = Visit(context.expression());
            IsLastResultLogical = false;
            return value + 1;
        }

        public override double VisitNegateExpr(SheetsParser.NegateExprContext context)
        {
            var value = Visit(context.expression());
            IsLastResultLogical = false;
            return -value;
        }
        public override double VisitAffirmExpr(SheetsParser.AffirmExprContext context)
        {
            var value = Visit(context.expression());
            IsLastResultLogical = false;
            return value;
        }

        public override double VisitParenExpr(SheetsParser.ParenExprContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitEqualExpr(SheetsParser.EqualExprContext context)
        {
            var left = Visit(context.expression(0));
            var right = Visit(context.expression(1));
            IsLastResultLogical = true;
            return left == right ? 1 : 0;
        }

        public override double VisitLessExpr(SheetsParser.LessExprContext context)
        {
            var left = Visit(context.expression(0));
            var right = Visit(context.expression(1));
            IsLastResultLogical = true;
            return left < right ? 1 : 0;
        }
        public override double VisitGreaterExpr(SheetsParser.GreaterExprContext context)
        {
            var left = Visit(context.expression(0));
            var right = Visit(context.expression(1));
            IsLastResultLogical = true;
            return left > right ? 1 : 0;
        }

        public override double VisitDenialExpr(SheetsParser.DenialExprContext context)
        {
            var value = Visit(context.expression());
            IsLastResultLogical = true;
            return value!=0 ? 0 : 1;
        }

        public override double VisitEqualSignExpr(SheetsParser.EqualSignExprContext context)
        {
            return Visit(context.expression());
        }

        private Dictionary<string, double> evaluatedValues = new Dictionary<string, double>();

        private HashSet<string> visiting = new HashSet<string>();

        public override double VisitCellExpr(SheetsParser.CellExprContext context)
        {
            string cellName = context.GetText();
            var cell = cells.Find(c => c.Name == cellName);

            if (cell == null)
            {
                visiting.Remove(cellName);
                throw new CellNotFoundException($"Клітина {cellName} не знайдена");
            }

            if (!cell.Dependents.Contains(currentCell))
            {
                cell.Dependents.Add(currentCell);
            }

            if (evaluatedValues.ContainsKey(cellName))
            {
                return evaluatedValues[cellName];
            }

            if (visiting.Contains(cellName))
            {
                throw new CycleException($"Циклічне посилання на {cellName}");
            }

            visiting.Add(cellName);
            try
            {
                double result;
                try
                {
                    result = Eval(cell);
                }
                catch (Exception ex)
                {
                    if (ex is CycleException || ex is CellNotFoundException)
                    {
                        throw;
                    }

                    if (string.IsNullOrEmpty(cell.Expression))
                    {
                        result = 0.0;
                    }

                    else if (double.TryParse(cell.Expression, System.Globalization.NumberStyles.Float,
                             System.Globalization.CultureInfo.InvariantCulture, out double numericResult))
                    {
                        result = numericResult;
                    }

                    else
                    {
                        throw;
                    }
                }

                evaluatedValues[cellName] = result;
                return result;
            }
            catch
            {
                evaluatedValues[cellName] = double.NaN;
                throw;
            }
            finally
            {
                visiting.Remove(cellName);
            }
        }
        public override double VisitNumberExpr(SheetsParser.NumberExprContext context)
        {
            string text = context.GetText().Replace(',', '.');
            if (double.TryParse(text, System.Globalization.NumberStyles.Float,
                     System.Globalization.CultureInfo.InvariantCulture, out double result))
            {
                IsLastResultLogical = false;
                return result;
            }
            throw new Exception("Не вдалося розпарсити як число.");
        }

    }
}
