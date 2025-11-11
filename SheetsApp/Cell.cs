using System;
using System.Collections.Generic;
using System.Linq;
namespace SheetsApp
{
    public class Cell
    {
        public string Name { get; set; }
        public string Expression { get; set; }
        public string Value { get; set; }
        public List<Cell> Dependents { get; } = new List<Cell>();

        public Cell(string name, string expression = "", string value = "")
        {
            Name = name;
            Expression = expression;
            Value = value;
        }
    }
}
