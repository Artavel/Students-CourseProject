using Reports;
using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace AccountingPerformanceView
{
    // Класс формы для показа отчётов (предварительный просмотр перед печатью)
    public partial class ReportsForm : Form
    {
        private readonly Report report;

        // Объект отчета передается через параметр конструктора
        public ReportsForm(Report report)
        {
            InitializeComponent();
            this.report = report;
            Text = report.Caption;
            // имя документа берется из заголовка отчета
            printDoc.DocumentName = report.Caption;
        }

        // Обработчик события начала печати страницы
        private void printDoc_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            RectangleF prnrect = e.PageBounds;
            prnrect.Inflate(new Size(-30, -50));
            var printdate = "Отпечатано: " + DateTime.Now;
            var printpagenumber = "Страница 1 из 1";
            SizeF strSize;
            PointF strPoint;
            // печать подвалов
            using (var printfont = new Font("Arial", 8, FontStyle.Italic))
            {
                strPoint = prnrect.Location;
                strPoint.Y += prnrect.Height;
                e.Graphics.DrawString(printdate, printfont, Brushes.Black, strPoint);
                //------------------
                strSize = e.Graphics.MeasureString(printpagenumber, printfont);
                strPoint = prnrect.Location;
                strPoint.X += prnrect.Width - strSize.Width;
                strPoint.Y += prnrect.Height;
                e.Graphics.DrawString(printpagenumber, printfont, Brushes.Black, strPoint);
            }
            // печать заголовка отчёта
            strPoint.Y = prnrect.Y;
            strPoint.X = prnrect.X;
            using (var printfont = new Font("Arial", 16, FontStyle.Bold))
            using (var sf = new StringFormat())
            {
                sf.Alignment = StringAlignment.Center;
                e.Graphics.DrawString(report.Caption, printfont, Brushes.DarkRed,
                    new RectangleF(strPoint, new SizeF(prnrect.Width, 50)), sf);
                strSize = e.Graphics.MeasureString(report.Caption, printfont, (int)prnrect.Width, sf);
            }
            strPoint.X = prnrect.X;
            strPoint.Y += strSize.Height * 1.5f;
            if (report.PrintPage == null) return;
            // печать собственно отчета
            report.PrintPage(sender, e, prnrect, strPoint);
        }

        // Обработчик нажатия на кнопку "Печать"
        private void tsbPrint_Click(object sender, EventArgs e)
        {
            printDoc.Print();
            Close();
        }

        // Обработчик нажатия на кнопку "Выбор принтера и настройка печати"
        private void tsbPrintDialog_Click(object sender, EventArgs e)
        {
            if (printDialog.ShowDialog(this) == DialogResult.OK)
            {
                printDoc.Print();
            }
        }

        // Обработчик кнопки увеличения масштаба
        private void tssbZoom_ButtonClick(object sender, EventArgs e)
        {
            var zooms = new[] { 0.25, 0.5, 0.75, 1, 1.5, 2 };
            foreach (var zoom in zooms.Where(zoom => printPreview.Zoom < zoom))
            {
                printPreview.Zoom = zoom;
                break;
            }
        }

        // Обработчик выбора различных масштабов из меню
        private void tsmZoomAuto_Click(object sender, EventArgs e)
        {
            var tsmi = (ToolStripMenuItem)sender;
            if (tsmi.Tag == null) return;
            double k;
            if (!double.TryParse(tsmi.Tag.ToString(), NumberStyles.AllowDecimalPoint, 
                CultureInfo.InvariantCulture, out k)) return;
            if (k > 0)
                printPreview.Zoom = k / 100.0;
            else
                printPreview.AutoZoom = true;
        }

        // Обработчик кнопки уменьшения масштаба
        private void tsbZoomOut_Click(object sender, EventArgs e)
        {
            var zooms = new[] { 0.25, 0.5, 0.75, 1, 1.5, 2 };
            Array.Reverse(zooms);
            foreach (var zoom in zooms.Where(zoom => printPreview.Zoom > zoom))
            {
                printPreview.Zoom = zoom;
                break;
            }
        }
    }
}
