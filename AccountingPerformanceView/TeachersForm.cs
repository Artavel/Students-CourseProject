using AccountingPerformanceModel;
using System.Windows.Forms;
using ViewGenerator;

namespace AccountingPerformanceView
{
    public partial class TeachersForm : Form
    {
        Root _root;

        public TeachersForm(Root root)
        {
            InitializeComponent();
            _root = root;
            // содаем панель с таблицей автоматически по классу и списку
            panel1.Controls.Add(GridPanelBuilder.BuildPropertyPanel(root, new Teacher(), root.Teachers));
        }

        // Прячем форму при закрытии
        private void TeachersForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }
    }
}
