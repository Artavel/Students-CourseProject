using AccountingPerformanceModel;
using System;
using System.Linq;
using System.Windows.Forms;
using ViewGenerator;

namespace AccountingPerformanceView
{
    // Класс формы редактирования курсов
    public partial class MattersCoursesForm : Form
    {
        Root _root;

        // Конструктор формы
        public MattersCoursesForm(Root root)
        {
            InitializeComponent();
            _root = root;
            // создаем панели с таблицами автоматически по классу и списку из модели
            EmptyPanel();
        }

        // Заполняем селектор специальностей
        private void cbSpecialities_DropDown(object sender, System.EventArgs e)
        {
            var speciality = (Speciality)cbSpecialities.SelectedItem;
            cbSpecialities.Items.Clear();
            foreach (var item in _root.Specialities)
                cbSpecialities.Items.Add(item);
            cbSpecialities.SelectedItem = speciality;
        }

        // При выборе специальности очищаем селектор специализаций и запрещаем таблицу курсов
        private void cbSpecialities_SelectionChangeCommitted(object sender, System.EventArgs e)
        {
            cbSpecializations.SelectedItem = null;
            EmptyPanel();
        }

        private void EmptyPanel()
        {
            panel1.Controls.Add(GridPanelBuilder.BuildPropertyPanel(_root, new MattersCourse(),
                _root.MattersCourses.FilteredBySpecialityAndSpecialization(Guid.Empty, Guid.Empty)));
            if (panel1.Controls.Count > 1) panel1.Controls.RemoveAt(0);
            panel1.Enabled = false;
        }

        // Заполняем селектор специализаций
        private void cbSpecializations_DropDown(object sender, System.EventArgs e)
        {
            // получаем выбранную специальность
            var speciality = (Speciality)cbSpecialities.SelectedItem;
            // получаем выбранную специализацию
            var specialization = (Specialization)cbSpecializations.SelectedItem;
            // очищаем список специализаций
            cbSpecializations.Items.Clear();
            if (speciality == null) return;
            // заполняем список специализаций только для выбранной специальности
            foreach (var item in _root.Specializations.Where(x => x.IdSpeciality == speciality.IdSpeciality))
                cbSpecializations.Items.Add(item);
            cbSpecializations.SelectedItem = specialization;
        }

        // При выборе специализации заполняем таблицу курсов, применяя два фильтра-селектора
        private void cbSpecializations_SelectionChangeCommitted(object sender, System.EventArgs e)
        {
            // получаем выбранную специальность
            var speciality = (Speciality)cbSpecialities.SelectedItem;
            // получаем выбранную специализацию
            var specialization = (Specialization)cbSpecializations.SelectedItem;
            // если один из фильтров не выбран, выходим
            if (speciality == null || specialization == null) return;
            // заполняем таблицу фильтрованными значениями
            panel1.Controls.Add(GridPanelBuilder.BuildPropertyPanel(_root, new MattersCourse(), _root.MattersCourses,
                  _root.MattersCourses.FilteredBySpecialityAndSpecialization(speciality.IdSpeciality, specialization.IdSpecialization), 
                               new[] { "IdSpeciality", "IdSpecialization" }, 
                               new[] { speciality.IdSpeciality, specialization.IdSpecialization }));
            // предыдущую панель убираем
            panel1.Controls.RemoveAt(0);
            panel1.Enabled = true;

        }

        // Прячем форму при закрытии
        private void MattersCoursesForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }
    }
}
