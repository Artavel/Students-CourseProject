using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;

namespace AccountingPerformanceModel
{
    // Класс поддержки чтения/записи конфигурации в файл на локальном диске
    public static class SaverLoader
    {
        // Метод загрузки сохранённой ранее конфигурации из локального файла
        public static Root LoadFromFile(string fileName)
        {
            using (var fs = File.OpenRead(fileName))
            using (var zip = new GZipStream(fs, CompressionMode.Decompress))
            {
                var formatter = new BinaryFormatter();
                return (Root)formatter.Deserialize(zip);
            }
        }

        // Метод сохранения конфигурации в файл на локальном диске
        public static void SaveToFile(string fileName, Root root)
        {
            using (var fs = File.Create(fileName))
            using (var zip = new GZipStream(fs, CompressionMode.Compress))
            {
                var formatter = new BinaryFormatter();
                formatter.Serialize(zip, root);
            }
        }

        // Метод загрузки сохранённой ранее конфигурации из базы данных
        public static Root LoadFromBase(string connection)
        {
            var root = new Root();
            var server = new OleDbServer { Connection = connection };

            // предметы
            var dataSet = server.GetRows("Matters");
            if (dataSet.Tables.Count > 0)
                foreach (var row in dataSet.Tables[0].Rows.Cast<DataRow>())
                {
                    if (row.ItemArray.Length != 2) continue;
                    root.Matters.Add(new Matter
                    {
                        IdMatter = Guid.Parse(row.ItemArray[0].ToString().Substring(1)),
                        Name = row.ItemArray[1].ToString()
                    });
                }

            root.Matters.Loaded = true;
            OperationResult = server.LastError;
            if (!string.IsNullOrWhiteSpace(OperationResult)) return root;
            // специальности
            dataSet = server.GetRows("Specialities");
            if (dataSet.Tables.Count > 0)
                foreach (var row in dataSet.Tables[0].Rows.Cast<DataRow>())
                {
                    if (row.ItemArray.Length != 3) continue;
                    root.Specialities.Add(new Speciality
                    {
                        IdSpeciality = Guid.Parse(row.ItemArray[0].ToString().Substring(1)),
                        Name = row.ItemArray[1].ToString(),
                        Number = int.Parse(row.ItemArray[2].ToString())
                    });
                }

            root.Specialities.Loaded = true;
            OperationResult = server.LastError;
            if (!string.IsNullOrWhiteSpace(OperationResult)) return root;

            // специализации
            dataSet = server.GetRows("Specializations");
            if (dataSet.Tables.Count > 0)
                foreach (var row in dataSet.Tables[0].Rows.Cast<DataRow>())
                {
                    if (row.ItemArray.Length != 3) continue;
                    root.Specializations.Add(new Specialization
                    {
                        IdSpecialization = Guid.Parse(row.ItemArray[0].ToString().Substring(1)),
                        Name = row.ItemArray[1].ToString(),
                        IdSpeciality = Guid.Parse(row.ItemArray[2].ToString().Substring(1))
                    });
                }
            
            root.Specializations.Loaded = true;
            OperationResult = server.LastError;
            if (!string.IsNullOrWhiteSpace(OperationResult)) return root;

            // курсы предметов
            dataSet = server.GetRows("MattersCourses");
            if (dataSet.Tables.Count > 0)
                foreach (var row in dataSet.Tables[0].Rows.Cast<DataRow>())
                {
                    if (row.ItemArray.Length != 7) continue;
                    root.MattersCourses.Add(new MattersCourse
                    {
                        IdMattersCourse = Guid.Parse(row.ItemArray[0].ToString().Substring(1)),
                        IdSpeciality = Guid.Parse(row.ItemArray[1].ToString().Substring(1)),
                        IdSpecialization = Guid.Parse(row.ItemArray[2].ToString().Substring(1)),
                        IdMatter = Guid.Parse(row.ItemArray[3].ToString().Substring(1)),
                        CourseType = (CourseType)Enum.Parse(typeof(CourseType), row.ItemArray[4].ToString()),
                        RatingSystem = (RatingSystem)Enum.Parse(typeof(RatingSystem), row.ItemArray[5].ToString()),
                        HoursCount = float.Parse(row.ItemArray[6].ToString())
                    });
                }
            
            root.MattersCourses.Loaded = true;
            OperationResult = server.LastError;
            if (!string.IsNullOrWhiteSpace(OperationResult)) return root;

            // семестры
            dataSet = server.GetRows("Semesters");
            if (dataSet.Tables.Count > 0)
                foreach (var row in dataSet.Tables[0].Rows.Cast<DataRow>())
                {
                    if (row.ItemArray.Length != 2) continue;
                    root.Semesters.Add(new Semester
                    {
                        IdSemester = Guid.Parse(row.ItemArray[0].ToString().Substring(1)),
                        Number = int.Parse(row.ItemArray[1].ToString())
                    });
                }
            
            root.Semesters.Loaded = true;
            OperationResult = server.LastError;
            if (!string.IsNullOrWhiteSpace(OperationResult)) return root;

            // учебные группы
            dataSet = server.GetRows("StudyGroups");
            if (dataSet.Tables.Count > 0)
                foreach (var row in dataSet.Tables[0].Rows.Cast<DataRow>())
                {
                    if (row.ItemArray.Length != 5) continue;
                    root.StudyGroups.Add(new StudyGroup
                    {
                        IdStudyGroup = Guid.Parse(row.ItemArray[0].ToString().Substring(1)),
                        Number = row.ItemArray[1].ToString(),
                        TrainingPeriod = float.Parse(row.ItemArray[2].ToString()),
                        IdSpeciality = Guid.Parse(row.ItemArray[3].ToString().Substring(1)),
                        IdSpecialization = Guid.Parse(row.ItemArray[4].ToString().Substring(1))
                    });
                }
            
            root.StudyGroups.Loaded = true;
            OperationResult = server.LastError;
            if (!string.IsNullOrWhiteSpace(OperationResult)) return root;

            // успеваемости
            dataSet = server.GetRows("Performances");
            if (dataSet.Tables.Count > 0)
                foreach (var row in dataSet.Tables[0].Rows.Cast<DataRow>())
                {
                    if (row.ItemArray.Length != 5) continue;
                    root.Performances.Add(new Performance
                    {
                        IdPerformance = Guid.Parse(row.ItemArray[0].ToString().Substring(1)),
                        IdSemester = Guid.Parse(row.ItemArray[1].ToString().Substring(1)),
                        IdMatter = Guid.Parse(row.ItemArray[2].ToString().Substring(1)),
                        Grade = (Grade)Enum.Parse(typeof(Grade), row.ItemArray[3].ToString()),
                        IdStudent = Guid.Parse(row.ItemArray[4].ToString().Substring(1))
                    });
                }
            
            root.Performances.Loaded = true;
            OperationResult = server.LastError;
            if (!string.IsNullOrWhiteSpace(OperationResult)) return root;

            // студенты
            dataSet = server.GetRows("Students");
            if (dataSet.Tables.Count > 0)
                foreach (var row in dataSet.Tables[0].Rows.Cast<DataRow>())
                {
                    if (row.ItemArray.Length != 13) continue;
                    var buff = (byte[])row.ItemArray[9];
                    root.Students.Add(new Student
                    {
                        IdStudent = Guid.Parse(row.ItemArray[0].ToString().Substring(1)),
                        FullName = row.ItemArray[1].ToString(),
                        BirthDay = DateTime.Parse(row.ItemArray[2].ToString()),
                        EducationCertificate = row.ItemArray[3].ToString(),
                        ReceiptDate = DateTime.Parse(row.ItemArray[4].ToString()),
                        Address = row.ItemArray[5].ToString(),
                        PhoneNumber = row.ItemArray[6].ToString(),
                        SocialStatus = row.ItemArray[7].ToString(),
                        Notes = row.ItemArray[8].ToString(),
                        Photo = buff,
                        IdSpeciality = Guid.Parse(row.ItemArray[10].ToString().Substring(1)),
                        IdSpecialization = Guid.Parse(row.ItemArray[11].ToString().Substring(1)),
                        IdStudyGroup = Guid.Parse(row.ItemArray[12].ToString().Substring(1))
                    });
                }
            
            root.Students.Loaded = true;
            OperationResult = server.LastError;
            if (!string.IsNullOrWhiteSpace(OperationResult)) return root;

            // преподаватели
            dataSet = server.GetRows("Teachers");
            if (dataSet.Tables.Count > 0)
                foreach (var row in dataSet.Tables[0].Rows.Cast<DataRow>())
                {
                    if (row.ItemArray.Length != 5) continue;
                    root.Teachers.Add(new Teacher
                    {
                        IdTeacher = Guid.Parse(row.ItemArray[0].ToString().Substring(1)),
                        FullName = row.ItemArray[1].ToString(),
                        IdMatter = Guid.Parse(row.ItemArray[2].ToString().Substring(1)),
                        Login = row.ItemArray[3].ToString(),
                        Password = row.ItemArray[4].ToString()
                    });
                }
            
            root.Teachers.Loaded = true;
            OperationResult = server.LastError;
            if (!string.IsNullOrWhiteSpace(OperationResult)) return root;

            return root;
        }

        // Свойство для хранения результата (текста ошибки) последней операции
        public static string OperationResult { get; private set; } = string.Empty;
    }
}
