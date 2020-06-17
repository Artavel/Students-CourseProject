using System;
using System.ComponentModel;

namespace AccountingPerformanceModel
{
    [Serializable]
    public enum Grade
    {
        [Description("1")]
        Один,
        [Description("2")]
        Два,
        [Description("3")]
        Три,
        [Description("4")]
        Четыре,
        [Description("5")]
        Пять,
        [Description("6")]
        Шесть,
        [Description("7")]
        Семь,
        [Description("8")]
        Восемь,
        [Description("9")]
        Девять,
        [Description("зачет")]
        Зачёт,
        [Description("незачет")]
        Незачёт,
        [Description("(не сдавал)")]
        Нет,
    }

}
