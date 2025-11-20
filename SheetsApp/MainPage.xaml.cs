using Antlr4.Runtime;
using ClosedXML.Excel;
using CommunityToolkit.Maui.Storage;
using Microsoft.Maui;
using Microsoft.Maui.Graphics.Text;
using Grid = Microsoft.Maui.Controls.Grid;

namespace SheetsApp
{
    public partial class MainPage : ContentPage
    {
        public Grid MainGrid => grid;
        const int CountColumn = 10;
        const int CountRow = 10;

        private bool isDirty = false;
        private bool isClosing = false;


        private Dictionary<(int, int), Cell> cells = new Dictionary<(int, int), Cell>();

        private Entry lastEntryFocused;

        private HashSet<string> visiting = new HashSet<string>();

        public MainPage()
        {
            InitializeComponent();
            CreateGrid();
        }

        private void CreateGrid()
        {
            AddColumnsAndColumnLabels();
            AddRowsAndCellEntries();
            isDirty = false;
        }
        private Label CreateLabel(string text, int row, int column)
        {
            var label = new Label
            {
                Text = text,
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center,
                Padding = 5
            };
            Grid.SetRow(label, row);
            Grid.SetColumn(label, column);
            return label;
        }
        private Entry CreateCellEntry(int row, int column)
        {
            var entry = new Entry
            {
                Text = "",
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center,
                WidthRequest = 100,

            };
            entry.Focused += OnEntryFocused;
            entry.Unfocused += OnEntryUnfocused;
            entry.TextChanged += OnEntryTextChanged;
            cells[(row, column)] = new Cell(GetCellName(row, column), "", "");
            Grid.SetRow(entry, row);
            Grid.SetColumn(entry, column);
            return entry;
        }
        internal Entry GetEntryAt(int row, int column)
        {
            foreach (var child in grid.Children)
            {
                if (grid.GetRow(child) == row && grid.GetColumn(child) == column && child is Entry entry)
                {
                    return entry;
                }
            }
            return null;
        }

        private string FormatForDisplay(string value)
        {
            return value;
        }

        private string GetColumnName(int colIndex)
        {
            int dividend = colIndex;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;

            }
            return columnName;
        }
        private string GetCellName(int row, int column)
        {
            return GetColumnName(column) + row.ToString();
        }
        private void RemoveGridElements(Func<Microsoft.Maui.IView, bool> matchCondition)
        {
            var childrenToRemove = new List<Microsoft.Maui.IView>();
            foreach (var child in grid.Children)
            {
                if (matchCondition(child))
                {
                    childrenToRemove.Add(child);
                }
            }

            foreach (var child in childrenToRemove)
            {
                grid.Children.Remove(child);
            }
        }
        private void AddRow(int newRow)
        {
            foreach (var child in grid.Children)
            {
                int currentRow = grid.GetRow(child);
                if (currentRow >= newRow)
                {
                    if (child is Label label)
                    {
                        if (grid.GetColumn(label) == 0)
                        {
                            label.Text = (currentRow + 1).ToString();
                        }
                    }
                    grid.SetRow(child, currentRow + 1);
                }
            }

            grid.RowDefinitions.Insert(newRow, new RowDefinition { Height = GridLength.Auto });

            var rowLabel = CreateLabel(newRow.ToString(), newRow, 0);
            grid.Children.Add(rowLabel);

            for (int col = 1; col < grid.ColumnDefinitions.Count; col++)
            {
                var entry = CreateCellEntry(newRow, col);
                grid.Children.Add(entry);
            }
            isDirty = true;

            RecalculateAllValues();
        }
        private string ProcessCell(Cell cell, SheetsAppVisitor visitor)
        {
            try
            {
                double numericValue = visitor.Eval(cell);

                if (visitor.IsLastResultLogical)
                {
                    cell.Value = (numericValue != 0) ? "ІСТИНА" : "ХИБА";
                }
                else
                {
                    cell.Value = numericValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
                return null;
            }
            catch (CycleException ex)
            {
                cell.Value = "#CYCERROR";
                return ex.Message;
            }
            catch (CellNotFoundException ex)
            {
                cell.Value = "#REF_ERROR";
                return ex.Message;
            }
            catch (Exception)
            {
                bool handled = false;
                Antlr4.Runtime.Tree.IParseTree tree = null;

                try
                {
                    var lexer = new SheetsLexer(new AntlrInputStream(cell.Expression));
                    lexer.RemoveErrorListeners();
                    lexer.AddErrorListener(ThrowingErrorListener.Instance);

                    var tokens = new CommonTokenStream(lexer);
                    var parser = new SheetsParser(tokens);

                    parser.RemoveErrorListeners();
                    parser.AddErrorListener(ThrowingErrorListener.Instance);

                    tree = parser.expression();

                    var tempTree = tree;
                    while (tempTree is SheetsParser.EqualSignExprContext equalCtx)
                    {
                        tempTree = equalCtx.expression();
                    }

                    if (tempTree is SheetsParser.CellExprContext cellCtx)
                    {
                        string referencedCellName = cellCtx.GetText();
                        var referencedCell = cells.Values.FirstOrDefault(c => c.Name == referencedCellName);
                        if (referencedCell != null)
                        {
                            cell.Value = referencedCell.Value;
                            handled = true;
                        }
                    }
                }
                catch
                {
                    tree = null;
                }

                if (!handled)
                {
                    string expr = cell.Expression.Trim().ToLower();
                    bool startsWithKeyWord =
                        expr.StartsWith("=") ||
                        expr.StartsWith("not") ||
                        expr.StartsWith("inc") ||
                        expr.StartsWith("dec");

                    bool isComplexFormula = false;

                    if (tree != null)
                    {
                        bool isPrimitive =
                            tree is SheetsParser.CellExprContext ||
                            tree is SheetsParser.NumberExprContext;

                        if (!isPrimitive)
                        {
                            isComplexFormula = true;
                        }
                    }

                    if (isComplexFormula || startsWithKeyWord)
                    {
                        cell.Value = "#SYNERROR";
                    }
                    else
                    {
                        cell.Value = cell.Expression;
                    }
                }
                return null;
            }
        }
        private void UpdateDependants(Cell cell)
        {
            foreach (var dependant in cell.Dependents.ToList())
            {
                if (visiting.Contains(dependant.Name))
                    continue;

                visiting.Add(dependant.Name);

                var visitor = new SheetsAppVisitor(cells);
                ProcessCell(dependant, visitor);

                var dependantCoords = cells.FirstOrDefault(kv => kv.Value == dependant).Key;
                var entry = GetEntryAt(dependantCoords.Item1, dependantCoords.Item2);
                if (entry != null)
                {
                    entry.TextChanged -= OnEntryTextChanged;
                    entry.Text = FormatForDisplay(dependant.Value);
                    entry.TextChanged += OnEntryTextChanged;
                }

                UpdateDependants(dependant);
                visiting.Remove(dependant.Name);
            }
        }
        private void OnEntryFocused(object sender, FocusEventArgs e)
        {
            lastEntryFocused = (Entry)sender;
            var cell = cells[(grid.GetRow(lastEntryFocused), grid.GetColumn(lastEntryFocused))];
            textInput.TextChanged -= OnTextInputTextChanged;
            textInput.Text = cell.Expression;
            textInput.TextChanged += OnTextInputTextChanged;
        }

        private async void OnEntryUnfocused(object sender, FocusEventArgs e)
        {
            var entryLosingFocus = (Entry)sender;
            var cell = cells[(grid.GetRow(entryLosingFocus), grid.GetColumn(entryLosingFocus))];

            string oldValue = cell.Value;

            foreach (var otherCell in cells.Values)
            {
                if (otherCell.Dependents.Contains(cell))
                {
                    otherCell.Dependents.Remove(cell);
                }
            }

            var visitor = new SheetsAppVisitor(cells);
            if (cell.Expression != "")
            {
                string errorMsg = ProcessCell(cell, visitor);

                if (errorMsg != null && cell.Value != oldValue)
                {
                    await DisplayAlert("Помилка", errorMsg, "OK");
                }
            }
            else
            {
                cell.Value = "";
            }

            UpdateDependants(cell);

            entryLosingFocus.TextChanged -= OnEntryTextChanged;
            entryLosingFocus.Text = FormatForDisplay(cell.Value);
            entryLosingFocus.TextChanged += OnEntryTextChanged;
        }
        private void OnEntryTextChanged(object sender, TextChangedEventArgs e)
        {
            textInput.Text = e.NewTextValue;
            var cell = cells[(grid.GetRow(lastEntryFocused), grid.GetColumn(lastEntryFocused))];
            cell.Expression = e.NewTextValue;
            isDirty = true;
        }
        private void OnTextInputTextChanged(object sender, TextChangedEventArgs e)
        {
            lastEntryFocused.Text = e.NewTextValue;
            var cell = cells[(grid.GetRow(lastEntryFocused), grid.GetColumn(lastEntryFocused))];
            cell.Expression = e.NewTextValue;

        }
        void AddColumn(int newColumn)
        {
            foreach (var child in grid.Children)
            {
                int currentColumn = grid.GetColumn(child);
                if (currentColumn >= newColumn)
                {
                    if (child is Label label)
                    {
                        if (grid.GetRow(label) == 0)
                        {
                            label.Text = GetColumnName(currentColumn + 1);
                        }
                    }
                    grid.SetColumn(child, currentColumn + 1);
                }
            }

            grid.ColumnDefinitions.Insert(newColumn, new ColumnDefinition { Width = GridLength.Auto });

            var colLabel = CreateLabel(GetColumnName(newColumn), 0, newColumn);
            grid.Children.Add(colLabel);

            for (int row = 1; row < grid.RowDefinitions.Count; row++)
            {
                var entry = CreateCellEntry(row, newColumn);
                grid.Children.Add(entry);
            }
            isDirty = true;

            RecalculateAllValues();
        }
        private async void DeleteRow(int lastRowIndex)
        {
            if (lastRowIndex < 2)
            {
                if (lastRowIndex < 2)
                {
                    await DisplayAlert("⚠️Увага!", "Не можна видалити останній рядок.", "OK");
                    return;
                }
            }

            foreach (var key in cells.Keys.Where(k => k.Item1 == lastRowIndex).ToList())
            {
                cells.Remove(key);
            }

            var updatedCells = new Dictionary<(int, int), Cell>();
            foreach (var kvp in cells)
            {
                var (row, col) = kvp.Key;
                if (row > lastRowIndex)
                {
                    updatedCells[(row - 1, col)] = kvp.Value;
                    kvp.Value.Name = GetCellName(row - 1, col);
                }
                else
                {
                    updatedCells[(row, col)] = kvp.Value;
                }
            }
            cells = updatedCells;

            RemoveGridElements(child => grid.GetRow(child) == lastRowIndex);
            grid.RowDefinitions.RemoveAt(lastRowIndex);
            isDirty = true;

            RecalculateAllValues();
        }


        private async void DeleteColumn(int lastColumnIndex)
        {
            if (lastColumnIndex < 2)
            {
                await DisplayAlert("⚠️Увага!", "Не можна видалити останній стовпчик.", "OK");
                return;
            }

            foreach (var key in cells.Keys.Where(k => k.Item2 == lastColumnIndex).ToList())
            {
                cells.Remove(key);
            }

            var updatedCells = new Dictionary<(int, int), Cell>();
            foreach (var kvp in cells)
            {
                var (row, col) = kvp.Key;
                if (col > lastColumnIndex)
                {
                    updatedCells[(row, col - 1)] = kvp.Value;
                    kvp.Value.Name = GetCellName(row, col - 1);
                }
                else
                {
                    updatedCells[(row, col)] = kvp.Value;
                }
            }
            cells = updatedCells;

            RemoveGridElements(child => grid.GetColumn(child) == lastColumnIndex);
            grid.ColumnDefinitions.RemoveAt(lastColumnIndex);
            isDirty = true;

            RecalculateAllValues();
        }


        private void AddColumnsAndColumnLabels()
        {
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            for (int col = 1; col < CountColumn + 1; col++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                var colLabel = CreateLabel(GetColumnName(col), 0, col);
                grid.Children.Add(colLabel);
            }
        }

        private void AddRowsAndCellEntries()
        {
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            for (int row = 1; row < CountRow + 1; row++)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var rowLabel = CreateLabel(row.ToString(), row, 0);
                grid.Children.Add(rowLabel);

                for (int col = 1; col < grid.ColumnDefinitions.Count; col++)
                {
                    var entry = CreateCellEntry(row, col);
                    grid.Children.Add(entry);
                }
            }
        }

        private async Task<bool> SaveFile(CancellationToken cancellationToken)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    foreach (var cellEntry in cells)
                    {
                        int row = cellEntry.Key.Item1;
                        int col = cellEntry.Key.Item2;
                        worksheet.Cell(row, col).Value = cellEntry.Value.Expression;
                    }

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;

                        var fileSaverResult = await FileSaver.Default.SaveAsync("sheet.xlsx", stream, cancellationToken);
                        isDirty = false;
                        return fileSaverResult != null;
                    }
                }
            }
            catch (Exception ex)
            {
                await DisplayAlert("Помилка", "При збереженні файлу виникла помилка.", "OK");
                return false;
            }
        }

        private async Task AskToSave()
        {
            if (!isDirty)
                return;

            bool answer = await DisplayAlert("Підтвердження", "Зберегти поточну таблицю?", "Так", "Ні");
            if (answer)
            {
                bool saved = await SaveFile(CancellationToken.None);
                if (!saved)
                {
                    await DisplayAlert("Помилка", "При збереженні файлу виникла помилка.", "OK");
                    return;
                }
            }
        }
        internal void AddRowButton_Clicked(object sender, EventArgs e)
        {
            int newRow = grid.RowDefinitions.Count;
            AddRow(newRow);
        }
        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            int newColumn = grid.ColumnDefinitions.Count;
            AddColumn(newColumn);
        }
        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            if (grid.RowDefinitions.Count > 1)
            {
                int lastRowIndex = grid.RowDefinitions.Count - 1;

                DeleteRow(lastRowIndex);
            }
        }
        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            if (grid.ColumnDefinitions.Count > 1)
            {
                int lastColumnIndex = grid.ColumnDefinitions.Count - 1;
                DeleteColumn(lastColumnIndex);
            }
        }
        private void SaveButton_Clicked(object sender, EventArgs e)
        {
            SaveFile(CancellationToken.None);
        }
        private async void CreateButton_Clicked(object sender, EventArgs e)
        {
            await AskToSave();
            grid.Children.Clear();
            grid.RowDefinitions.Clear();
            grid.ColumnDefinitions.Clear();
            CreateGrid();
        }


        private async void ReadButton_Clicked(object sender, EventArgs e)
        {
            await AskToSave();

            var result = await FilePicker.Default.PickAsync(new PickOptions
            {
                PickerTitle = "Виберіть файл типу .xlsx",
                FileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
        {
            { DevicePlatform.WinUI, new[] { ".xlsx", ".xls" } },
            { DevicePlatform.Android, new[] { ".xlsx", ".xls" } },
            { DevicePlatform.iOS, new[] { ".xlsx", ".xls" } },
            { DevicePlatform.MacCatalyst, new[] { ".xlsx", ".xls" } }
        })
            });

            if (result != null)
            {
                try
                {
                    using (var stream = await result.OpenReadAsync())
                    {
                        using (var workbook = new XLWorkbook(stream))
                        {
                            var worksheet = workbook.Worksheet(1);

                            grid.Children.Clear();
                            grid.RowDefinitions.Clear();
                            grid.ColumnDefinitions.Clear();
                            cells.Clear();

                            CreateGrid();

                            for (int row = 1; row <= worksheet.LastRowUsed().RowNumber(); row++)
                            {
                                if (row >= grid.RowDefinitions.Count)
                                {
                                    AddRow(grid.RowDefinitions.Count);
                                }

                                for (int col = 1; col <= worksheet.LastColumnUsed().ColumnNumber(); col++)
                                {
                                    if (col >= grid.ColumnDefinitions.Count)
                                    {
                                        AddColumn(grid.ColumnDefinitions.Count);
                                    }

                                    var cellExpression = worksheet.Cell(row, col).GetString();
                                    var cellKey = (row, col);

                                    if (!cells.ContainsKey(cellKey))
                                    {
                                        cells[cellKey] = new Cell(GetCellName(row, col), cellExpression, "");
                                    }
                                    else
                                    {
                                        cells[cellKey].Expression = cellExpression;
                                        cells[cellKey].Value = "";
                                    }

                                    var entry = GetEntryAt(row, col);
                                    if (entry != null)
                                    {
                                        entry.TextChanged -= OnEntryTextChanged;
                                        entry.Text = cellExpression;
                                        entry.TextChanged += OnEntryTextChanged;
                                    }
                                }
                            }

                            RecalculateAllValues();
                            isDirty = false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    await DisplayAlert("Помилка", "При відкритті файла виникла помилка", "OK");
                }
            }
        }

        private void RecalculateAllValues()
        {
            var visitor = new SheetsAppVisitor(cells);
            foreach (var cell in cells.Values)
            {
                if (!string.IsNullOrEmpty(cell.Expression))
                {
                    ProcessCell(cell, visitor);
                }
                else
                {
                    cell.Value = "";
                }

                var cellCoords = cells.FirstOrDefault(kv => kv.Value == cell).Key;
                var entry = GetEntryAt(cellCoords.Item1, cellCoords.Item2);
                if (entry != null)
                {
                    entry.TextChanged -= OnEntryTextChanged;
                    entry.Text = FormatForDisplay(cell.Value);
                    entry.TextChanged += OnEntryTextChanged;
                }
            }
        }

        private async void HelpButton_Clicked(object sender, EventArgs e)
        {
            await DisplayAlert("Довідка",
        "Лабораторна робота з ООП №1\n" +
        "Студента групи К27 Кучера Ярослава\n" +
        "Варіант 35. (1,3,5,8,10)\n\n" +
        "Операції:\n" +
        "+, -, *, / (бінарні)\n" +
        "+, - (унарні)\n" +
        "inc(), dec()\n" +
        "=, <, >\n" +
        "not()\n\n" +
        "⚠️ВАЖЛИВО: назви функцій (inc, dec, not) вводяться виключно малими літерами!",
        "OK");
        }

        private async void ExitButton_Clicked(object sender, EventArgs e)
        {
            await AttemptExitApp();
        }
        private async Task AttemptExitApp()
        {
            if (isClosing)
                return;

            bool answer = await DisplayAlert("Підтвердження", "Ви дійсно хочете вийти?", "Так", "Ні");
            if (answer)
            {
                await AskToSave();

                isClosing = true;
                Application.Current.Quit();
            }
        }
    }
}