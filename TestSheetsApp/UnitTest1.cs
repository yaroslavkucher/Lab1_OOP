using Xunit;
using Microsoft.Maui.Controls;
using SheetsApp;

namespace SheetsApp.Tests
{
    public class MainPageTests
    {
        [Fact]
        public void AddRow_ShouldIncreaseRowCountAndAddNewRowToGrid()
        {
            var mainPage = new MainPage();
            int initialRowCount = mainPage.MainGrid.RowDefinitions.Count;

            mainPage.AddRowButton_Clicked(null, null);

            int newRowCount = mainPage.MainGrid.RowDefinitions.Count;
            Assert.Equal(initialRowCount + 1, newRowCount);

            var lastRowLabel = mainPage.MainGrid.Children
                .OfType<Label>()
                .FirstOrDefault(label => Grid.GetRow(label) == newRowCount - 1 && Grid.GetColumn(label) == 0);
            Assert.NotNull(lastRowLabel);
            Assert.Equal((newRowCount - 1).ToString(), lastRowLabel.Text);

            for (int col = 1; col < mainPage.MainGrid.ColumnDefinitions.Count; col++)
            {
                var entry = mainPage.GetEntryAt(newRowCount - 1, col);
                Assert.NotNull(entry);
                Assert.Equal(string.Empty, entry.Text);
            }
        }
    }

    public class SheetsAppVisitorTests
    {
        [Fact]
        public void CyclicReference_ShouldThrowException()
        {
            var cellA = new Cell("A1", "B1 + 1");
            var cellB = new Cell("B1", "A1 + 1");

            var cells = new Dictionary<(int, int), Cell>
            {
                { (0, 0), cellA },
                { (0, 1), cellB }
            };

            var visitor = new SheetsAppVisitor(cells);

            var exception = Assert.Throws<Exception>(() => visitor.Eval(cellA));
            Assert.Contains("Циклічне посилання", exception.Message);
        }

        [Fact]
        public void Eval_ShouldRespectOrderOfOperations()
        {
            var cell = new Cell("A1", "3 + 2 * 4 - 5 / (1 + 1) ^ 2"); // (1+1)^2 -> 5/(1+1)^2 -> 2*4 -> 3+8-1.25 = 9.75

            var cells = new Dictionary<(int, int), Cell>
            {
                { (0, 0), cell }
            };

            var visitor = new SheetsAppVisitor(cells);

            double result = visitor.Eval(cell);

            Assert.Equal(9.75, result, precision: 2);
        }
    }
}