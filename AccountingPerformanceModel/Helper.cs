using System;
using System.Drawing;
using System.IO;
using System.Linq;

namespace AccountingPerformanceModel
{
    // Класс-помощник
    public static class Helper
    {
        private static Root _root;

        // Запоминаем ссылку на корневой объект модели
        public static void DefineRoot(Root root)
        {
            _root = root;
        }

        // Строка подключения к базе данных
        public static string ConnectionString { get; set; }

        // Создание картинки из массива байт
        public static Image CreateImage(byte[] imageData)
        {
            Image image;
            using (MemoryStream inStream = new MemoryStream())
            {
                inStream.Write(imageData, 0, imageData.Length);
                image = Bitmap.FromStream(inStream);
            }
            return image;
        }

        // Создание массива байт из картинки
        public static byte[] CreateByteArray(Image image)
        {
            using (MemoryStream outStream = new MemoryStream())
            {
                image.Save(outStream, System.Drawing.Imaging.ImageFormat.Png);
                return outStream.ToArray();
            }
        }

        // Получаем название семестра по его Id
        public static string SemesterById(Guid idSemester)
        {
            var semester = _root.Semesters.FirstOrDefault(item => item.IdSemester == idSemester);
            return semester != null ? semester.ToString() : idSemester.ToString();
        }

        // Определяем, что семестр используется в других таблицах
        public static bool SemesterUsed(Guid idSemester)
        {
            return _root.Performances.Any(item => item.IdSemester == idSemester);
        }

        // Определяем, что успеваемость используется в других таблицах
        public static bool PerformanceUsed(Guid idPerformance)
        {
            return false;
        }

        // Получаем название группы по её Id
        public static string StudyGroupById(Guid idStudyGroup)
        {
            var agroup = _root.StudyGroups.FirstOrDefault(item => item.IdStudyGroup == idStudyGroup);
            return agroup != null ? agroup.ToString() : idStudyGroup.ToString();
        }

        // Получаем название предмета по его Id
        public static string MatterById(Guid idMatter)
        {
            var matter = _root.Matters.FirstOrDefault(item => item.IdMatter == idMatter);
            return matter != null ? matter.ToString() : idMatter.ToString();
        }

        // Определяем, что предмет используется в других таблицах
        public static bool MatterUsed(Guid idMatter)
        {
            return _root.MattersCourses.Any(item => item.IdMatter == idMatter) ||
                   _root.Performances.Any(item => item.IdMatter == idMatter) ||
                   _root.Teachers.Any(item => item.IdMatter == idMatter) ;
        }

        // Получаем имя специальности по его Id
        public static string SpecialityById(Guid idSpeciality)
        {
            var speciality = _root.Specialities.FirstOrDefault(item => item.IdSpeciality == idSpeciality);
            return speciality != null ? speciality.ToString() : idSpeciality.ToString();
        }

        // Определяем, что специальность используется в других таблицах
        public static bool SpecialityUsed(Guid idSpeciality)
        {
            return _root.Specializations.Any(item => item.IdSpeciality == idSpeciality) ||
                   _root.StudyGroups.Any(item => item.IdSpeciality == idSpeciality) ||
                   _root.Students.Any(item => item.IdSpeciality == idSpeciality) ||
                   _root.MattersCourses.Any(item => item.IdSpeciality == idSpeciality);
        }

        // Получаем имя специальности по его Id
        public static string SpecializationById(Guid idSpecialization)
        {
            var specialization = _root.Specializations.FirstOrDefault(item => item.IdSpecialization == idSpecialization);
            return specialization != null ? specialization.ToString() : idSpecialization.ToString();
        }

        // Определяем, что специализация используется в других таблицах
        public static bool SpecializationUsed(Guid idSpecialization)
        {
            return _root.StudyGroups.Any(item => item.IdSpecialization == idSpecialization) ||
                   _root.MattersCourses.Any(item => item.IdSpecialization == idSpecialization) ||
                   _root.Students.Any(item => item.IdSpecialization == idSpecialization);
        }

        // Определяем, что курс используется в других таблицах
        public static bool MattersCourseUsed(Guid idMattersCourse)
        {
            return false;
        }

        // Получаем название студента по его Id
        public static string StudentById(Guid idStudent)
        {
            var student = _root.Students.FirstOrDefault(item => item.IdStudent == idStudent);
            return student != null ? student.ToString() : idStudent.ToString();
        }

        // Получаем студента по его Id
        public static Student GetStudentById(Guid idStudent)
        {
            return _root.Students.FirstOrDefault(item => item.IdStudent == idStudent);
        }

        // Получаем Id группы студента по его Id
        public static Guid GetStudentGroupId(Guid idStudent)
        {
            var student = _root.Students.FirstOrDefault(item => item.IdStudent == idStudent);
            return student != null ? student.IdStudyGroup : Guid.Empty;
        }

        // Определяем, что данные студента используются в других таблицах
        public static bool StudentUsed(Guid idStudent)
        {
            return _root.Performances.Any(item => item.IdStudent == idStudent);
        }

        // Определяем, что группа используется в других таблицах
        public static bool StudyGroupUsed(Guid idStudyGroup)
        {
            return _root.Students.Any(item => item.IdStudyGroup == idStudyGroup);
        }

        // Получаем группу по её Id
        public static StudyGroup GetStudyGroupById(Guid idStudyGroup)
        {
            return _root.StudyGroups.FirstOrDefault(item => item.IdStudyGroup == idStudyGroup);
        }
    }
}
