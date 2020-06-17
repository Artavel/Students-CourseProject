using System.Collections.Generic;

namespace Reports
{
    // Колонки (заголовки) отчета
    public class ReportColumns : List<ReportColumn>
    {
        // Добавить заголовки
        public void Add(params ReportColumn[] args)
        {
            foreach (var item in args)
                base.Add(item);
        }
    }

    // Ширины колонок (заголовков) отчета
    public class ReportColumn
    {
        public ReportColumn(string text = "", int width = 100)
        {
            Text = text;
            Width = width;
        }

        public string Text { get; set; }
        public int Width { get; set; } = 100;
    }

}
