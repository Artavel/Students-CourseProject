using System.Collections.Generic;

namespace Reports
{
    // Строки отчета
    public class ReportRows :List<ReportRow>
    {
        // добавить строку отчета
        public void Add(int level = 0, params string[] args)
        {
            var row = new ReportRow() { Level = level };
            row.Add(level, args);
            base.Add(row);
        }
    }

    // Строка отчета
    public class ReportRow
    {
        public List<string> Items { get; set; } = new List<string>(); // значения в колонках
        public int Level { get; set; } // уровень вложенности строки

        // Добавить значения для колонок строки
        public void Add(int level = 0, params string[] args)
        {
            foreach (var item in args)
                Items.Add(item);
        }
    }
}
