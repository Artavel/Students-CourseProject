﻿using AccountingPerformanceModel;
using Microsoft.Office.Interop.Excel;
using Reports;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ViewGenerator;

namespace AccountingPerformanceView
{
    public partial class MainForm : Form, UpdateData
    {
        private Root _root = new Root();                    // ссылка на корневой класс модели
        public static MattersForm MattersForm;              // ссылка на форму справочника предметов
        public static SpecialitiesForm SpecialitiesForm;    // ссылка на форму справочника специальностей
        public static MattersCoursesForm MattersCoursesForm;    // ссылка на форму справочника курсов
        public static StudentForm StudentForm;              // ссылка на форму справочника курсов
        public static SemestersForm SemestersForm;          // ссылка на форму справочника семестров
        public static ListOfDebtorsForm ListOfDebtorsForm;  // ссылка на форму должников
        public static TeachersForm TeachersForm;            // ссылка на форму преподавателей
        private Student _student;
        private StudyGroup _group;
        private bool _loggedIn;

        // Конструктор главной формы
        public MainForm()
        {
            InitializeComponent();
            // передаем ссылку на корневой класс модели в класс-помощник
            Helper.DefineRoot(_root);
            // прицепляем обработчик события при ошибке ввода для панелей ввода данных
            GridPanelBuilder.Error += GridPanelBuilder_Error;
        }

        // Обработчик события первой загрузки главной формы
        private void MainForm_Load(object sender, EventArgs e)
        {
            string fileName;
            fileName = Path.ChangeExtension(System.Windows.Forms.Application.ExecutablePath, ".mdb");
            if (File.Exists(fileName))
            {
                Helper.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=AccountingPerformanceView.mdb;";
                _root = SaverLoader.LoadFromBase(Helper.ConnectionString);
                if (!string.IsNullOrWhiteSpace(SaverLoader.OperationResult))
                    Console.WriteLine("Ошибка загрузки: " + SaverLoader.OperationResult);
                // при загрузке из файла корневой объект вновь создается, поэтому снова передаем ссылку на него
                Helper.DefineRoot(_root);
                // регистрируем таблицы сущностей после загрузки из файла
                _root.RegistryTables();
            }

            UpdateInterface();

            panel2.Controls.Add(GridPanelBuilder.BuildPropertyPanel(_root, new StudyGroup(),
                 _root.StudyGroups.FilteredBySpecialityAndSpecialization(Guid.Empty, Guid.Empty)));
            panel2.Enabled = false;
            var panel = GridPanelBuilder.BuildPropertyPanel(_root, new Student(),
                 _root.Students.FilteredBySpecialityAndSpecialization(Guid.Empty, Guid.Empty));
            GridPanelBuilder.HideButtonsPanel(panel);
            panel3.Controls.Add(panel);
        }

        // Обработчик события ошибок ввода данных с панелей свойств
        private void GridPanelBuilder_Error(string message, string caption)
        {
            MessageBox.Show(this, message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // Обработчик меню входа
        private void tsmiLogin_Click(object sender, EventArgs e)
        {
            if (_root.Teachers.Count > 0)
            {
                var frm = new LoginForm(_root);
                var lastLogin = _loggedIn;
                _loggedIn = frm.ShowDialog(this) == DialogResult.OK;
                if (lastLogin) _loggedIn = lastLogin;
            }
            else
            {
                MessageBox.Show("Нет ни одной записи в таблице пользователей");
                _loggedIn = true;
            }
            UpdateInterface();
        }

        private void UpdateInterface()
        {
            tableLayoutPanel1.Enabled = tsmiOperations.Enabled = tsmiReports.Enabled = _loggedIn;
        }

        private void tsmiFile_DropDownOpening(object sender, EventArgs e)
        {
            tsmiMatters.Enabled = tsmiSpecialities.Enabled = tsmiTeachers.Enabled = 
                tsmiMatterCourses.Enabled = tsmiSemesters.Enabled = _loggedIn;
        }

        // Метод для показа дочених форм
        public static void ShowForm(Form form)
        {
            // показывае форму
            form.Show();
            // если была свернута, то разворачиваем
            if (form.WindowState == FormWindowState.Minimized)
                form.WindowState = FormWindowState.Normal;
            // перемещаем наверх
            form.BringToFront();
        }

        // Обработчик меню выхода из приложения
        private void tsmiExit_Click(object sender, EventArgs e)
        {
            // закрыть приложение
            Close();
        }

        // Обработчик меню редактирования справочника предметов
        private void tsmiMatters_Click(object sender, EventArgs e)
        {
            // если ранее не показывался, то создаем новое окно
            if (MattersForm == null)
                MattersForm = new MattersForm(_root);
            // показываем окно 
            ShowForm(MattersForm);
        }

        // Обработчик меню редактирования специальностей
        private void tsmiSpecialities_Click(object sender, EventArgs e)
        {
            // если ранее не показывался, то создаем новое окно
            if (SpecialitiesForm == null)
                SpecialitiesForm = new SpecialitiesForm(_root);
            // показываем окно 
            ShowForm(SpecialitiesForm);
        }

        // Обработчик меню редактирования справочника курсов
        private void tsmiMatterCourses_Click(object sender, EventArgs e)
        {
            // если ранее не показывался, то создаем новое окно
            if (MattersCoursesForm == null)
                MattersCoursesForm = new MattersCoursesForm(_root);
            // показываем окно 
            ShowForm(MattersCoursesForm);
        }

        // Обработчик меню редактирования справочника семестров
        private void tsmiSemesters_Click(object sender, EventArgs e)
        {
            // если ранее не показывался, то создаем новое окно
            if (SemestersForm == null)
                SemestersForm = new SemestersForm(_root);
            // показываем окно 
            ShowForm(SemestersForm);
        }

        // Заполняем селектор специальностей
        private void cbSpecialities_DropDown(object sender, EventArgs e)
        {
            var speciality = (Speciality)cbSpecialities.SelectedItem;
            cbSpecialities.Items.Clear();
            foreach (var item in _root.Specialities)
                cbSpecialities.Items.Add(item);
            cbSpecialities.SelectedItem = speciality;
        }

        // При выборе специальности очищаем селектор специализаций и запрещаем таблицу курсов
        private void cbSpecialities_SelectionChangeCommitted(object sender, EventArgs e)
        {
            cbSpecializations.SelectedItem = null;
            panel2.Enabled = false;
        }

        // Заполняем селектор специализаций
        private void cbSpecializations_DropDown(object sender, EventArgs e)
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
        private void cbSpecializations_SelectionChangeCommitted(object sender, EventArgs e)
        {
            // получаем выбранную специальность
            var speciality = (Speciality)cbSpecialities.SelectedItem;
            // получаем выбранную специализацию
            var specialization = (Specialization)cbSpecializations.SelectedItem;
            // если один из фильтров не выбран, выходим
            if (speciality == null || specialization == null) return;

            // заполняем таблицу фильтрованными значениями
            var panel = GridPanelBuilder.BuildPropertyPanel(_root, new StudyGroup(), _root.StudyGroups,
                 _root.StudyGroups.FilteredBySpecialityAndSpecialization(speciality.IdSpeciality, specialization.IdSpecialization),
                               new[] { "IdSpeciality", "IdSpecialization" },
                               new[] { speciality.IdSpeciality, specialization.IdSpecialization });
            panel.GridSelectedChanged += Panel_GroupSelectedChanged;
            panel2.Controls.Add(panel);
            // предыдущую панель убираем
            if (panel2.Controls.Count > 1)
                panel2.Controls.RemoveAt(0);
            panel2.Enabled = true;

            // заполняем таблицу фильтрованными значениями
            if (_group == null)
                panel = GridPanelBuilder.BuildPropertyPanel(_root, new Student(), _root.Students,
                      _root.Students.FilteredBySpecialityAndSpecialization(speciality.IdSpeciality, specialization.IdSpecialization),
                                   new[] { "IdSpeciality", "IdSpecialization" },
                                   new[] { speciality.IdSpeciality, specialization.IdSpecialization });
            else
                panel = GridPanelBuilder.BuildPropertyPanel(_root, new Student(), _root.Students,
                      _root.Students.FilteredBySpecialityAndSpecialization(speciality.IdSpeciality, specialization.IdSpecialization, _group.IdStudyGroup),
                                   new[] { "IdSpeciality", "IdSpecialization", "IdStudyGroup" },
                                   new[] { speciality.IdSpeciality, specialization.IdSpecialization, _group.IdStudyGroup });

            panel.GridSelectedChanged += Panel_GridSelectedChanged;
            panel3.Controls.Add(panel);
            // предыдущую панель убираем
            if (panel3.Controls.Count > 1)
                panel3.Controls.RemoveAt(0);
            GridPanelBuilder.HideButtonsPanel(panel);
            panel3.Enabled = true;
            btnStudent.Enabled = btnMoveToGroup.Enabled = false;
            _student = null;
            _group = null;
        }

        // Выбор строки в таблице групп изменился
        private void Panel_GroupSelectedChanged(object obj)
        {
            _group = (StudyGroup)obj;
            if (_group == null) return;
            var number = _group.Number;
            UpdateStudents();
            FingGroupText(number);
        }

        // Выбор строки в таблице студентов изменился
        private void Panel_GridSelectedChanged(object obj)
        {
            _student = (Student)obj;
            btnStudent.Enabled = btnMoveToGroup.Enabled = _student != null;
            // обновление данных в форме студента
            if (StudentForm != null && StudentForm.Visible)
            {
                // получаем выбранную специальность
                var speciality = (Speciality)cbSpecialities.SelectedItem;
                // получаем выбранную специализацию
                var specialization = (Specialization)cbSpecializations.SelectedItem;
                // если один из фильтров не выбран, выходим
                if (speciality == null || specialization == null) return;
                if (_student != null)
                    StudentForm.Build(speciality, specialization, _student);
            }
            // обновление выбора в таблице групп
            FingGroupText(_student != null ? Helper.StudyGroupById(_student.IdStudyGroup) : "");
        }

        private void FingGroupText(string text)
        {
            if (panel2.Controls.Count > 0)
            {
                var panel = (GridPanel)panel2.Controls[0];
                panel.GridSelectedChanged -= Panel_GroupSelectedChanged;
                try
                {
                    GridPanelBuilder.FindText(panel, text);
                }
                finally
                {
                    panel.GridSelectedChanged += Panel_GroupSelectedChanged;
                }
            }
        }

        // Меню операций открывается, проверяем доступность элементов меню
        private void tsmiOperations_DropDownOpening(object sender, EventArgs e)
        {
            tsmiAddStudent.Enabled = cbSpecialities.SelectedItem != null && cbSpecializations.SelectedItem != null;
            tsmiViewEditStudent.Enabled = tsmiRecordStudentResults.Enabled = 
                tsmiMoveStudentToGroup.Enabled = _student != null;
        }

        // Меню отчетов открывается, проверяем доступность элементов меню
        private void tsmiReports_DropDownOpening(object sender, EventArgs e)
        {
            tsmiSemesterStudentProgress.Enabled = tsmiSemesterGroupProgress.Enabled =
                tsmiSummaryStudentProgress.Enabled = _student != null;
        }

        // Делаем форму о студенте видимой
        private void EnshureStudentFormVisible()
        {
            // если ранее не показывался, то создаем новое окно
            if (StudentForm == null)
                StudentForm = new StudentForm(this, _root);
            // показываем окно 
            ShowForm(StudentForm);
        }

        // Добавить нового студента
        private void tsmiAddStudent_Click(object sender, EventArgs e)
        {
            EnshureStudentFormVisible();
            // получаем выбранную специальность
            var speciality = (Speciality)cbSpecialities.SelectedItem;
            // получаем выбранную специализацию
            var specialization = (Specialization)cbSpecializations.SelectedItem;
            // если один из фильтров не выбран, выходим
            if (speciality == null || specialization == null) return;
            StudentForm.Build(speciality, specialization, null);
        }

        // Изменить/просмотреть данные о студенте
        private void tsmiViewEditStudent_Click(object sender, EventArgs e)
        {
            EnshureStudentFormVisible();
            // получаем выбранную специальность
            var speciality = (Speciality)cbSpecialities.SelectedItem;
            // получаем выбранную специализацию
            var specialization = (Specialization)cbSpecializations.SelectedItem;
            // если один из фильтров не выбран, выходим
            if (speciality == null || specialization == null || _student == null) return;
            StudentForm.Build(speciality, specialization, _student);
        }

        // Реализация интерфейса UpdateData
        public void UpdateStudents()
        {
            // обновляем данные в сетке
            cbSpecializations_SelectionChangeCommitted(cbSpecializations, new EventArgs());
        }

        // Обработчик нажатия кнопки "Перевод студента"
        private void btnMoveToGroup_Click(object sender, EventArgs e)
        {
            var frm = new MoveStudentToGroupForm(_root, _student);
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                var std = _student.DeepClone();
                frm.Update(std); // применяем выбранные настройки
                _root.Students.ChangeTo(_student, std);
                UpdateStudents();
            }
        }

        // Обработчик нажатия кнопки "Список должников"
        private void btnListOfDebtors_Click(object sender, EventArgs e)
        {
            
            // если ранее не показывался, то создаем новое окно
            if (ListOfDebtorsForm == null)
                ListOfDebtorsForm = new ListOfDebtorsForm(this, _root);
            // показываем окно 
            ShowForm(ListOfDebtorsForm);
        }

        // Обработчик записи результатов группы
        private void tsmiRecordGroupResults_Click(object sender, EventArgs e)
        {
            new GroupPerformanceForm(_root).ShowDialog(this);
        }

        // Отчёт об успеваемости студента за семестр
        private void tsmiSemesterStudentProgress_Click(object sender, EventArgs e)
        {
            if (_student == null) return;
            new SelectSemesterForm(_root, _student).ShowDialog(this);
        }

        // Отчёт об успеваемости группы за семестр
        private void tsmiSemesterGroupProgress_Click(object sender, EventArgs e)
        {
            if (_student == null) return;
            new SelectSemesterForm(_root, Helper.GetStudyGroupById(_student.IdStudyGroup)).ShowDialog(this);
        }

        private void tsmiSummaryStudentProgress_Click(object sender, EventArgs e)
        {
            if (_student == null) return;
            var t = Type.Missing;
            var xl1 = new Microsoft.Office.Interop.Excel.Application();
            xl1.Visible = true;
            var book = xl1.Workbooks.Add(t);
            var lists = book.Worksheets;
            Worksheet list = lists.Item[1];
            var cell = xl1.Selection.Cells;
            list.Range["C3", t].Value2 = _student.ToString();
            list.Range["C4", t].Value2 = _student.BirthDay.ToShortDateString();
            list.Range["H5", t].Value2 = "Минский радиотехнический колледж";
            list.Range["H6", t].Value2 = "г.Минск";
            list.Range["C7", t].Value2 = _student.EducationCertificate;
            list.Range["C8", t].Value2 = Helper.GetStudyGroupById(_student.IdStudyGroup).TrainingPeriod + " года";
            list.Range["C9", t].Value2 = Helper.SpecialityById(_student.IdSpeciality);
            list.Range["C10", t].Value2 = Helper.SpecializationById(_student.IdSpecialization);
            var row = 12;
            foreach (var semester in _root.Semesters)
            {
                var report = ReportsBuilder.GetStudentPerformance(_root, semester, _student);
                if (report.ReportRows.Count == 0) continue;
                list.Range[$"A{row}", t].Value2 = $"Семестр {semester}";
                row++;
                for (var i = 0; i < report.ReportRows.Count; i++)
                {
                    list.Range[$"B{row}", t].Value2 = report.ReportRows[i].Items[0];
                    list.Range[$"H{row}", t].Value2 = report.ReportRows[i].Items[1];
                    row++;
                }
            }
        }

        private void tbFind_TextChanged(object sender, EventArgs e)
        {
            foreach (var student in _root.Students.Where(x => x.FullName.StartsWith(tbFind.Text, 
                                                         StringComparison.CurrentCultureIgnoreCase)))
            {
                var speciality = _root.Specialities.FirstOrDefault(x => x.IdSpeciality == student.IdSpeciality);
                cbSpecialities.Items.Clear();
                foreach (var item in _root.Specialities)
                    cbSpecialities.Items.Add(item);
                cbSpecialities.SelectedItem = speciality;
                // получаем выбранную специальность
                speciality = (Speciality)cbSpecialities.SelectedItem;
                // получаем выбранную специализацию
                var specialization = _root.Specializations.FirstOrDefault(x => x.IdSpecialization == student.IdSpecialization);
                // очищаем список специализаций
                cbSpecializations.Items.Clear();
                if (speciality == null) return;
                // заполняем список специализаций только для выбранной специальности
                foreach (var item in _root.Specializations.Where(x => x.IdSpeciality == speciality.IdSpeciality))
                    cbSpecializations.Items.Add(item);
                cbSpecializations.SelectedItem = specialization;
                cbSpecializations_SelectionChangeCommitted(cbSpecializations, new EventArgs());
                if (panel3.Controls.Count > 0)
                    GridPanelBuilder.FindText((GridPanel)panel3.Controls[0], tbFind.Text);
                break;
            }
        }

        private void tsmiTeachers_Click(object sender, EventArgs e)
        {
            // если ранее не показывался, то создаем новое окно
            if (TeachersForm == null)
                TeachersForm = new TeachersForm(_root);
            // показываем окно 
            ShowForm(TeachersForm);
        }
    }

    // Интерфейс для обновления сетки данных основной формы
    public interface UpdateData
    {
        void UpdateStudents();
    }
}
