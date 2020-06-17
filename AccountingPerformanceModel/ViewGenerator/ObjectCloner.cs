﻿using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace ViewGenerator
{
    // Глубокое клонирование для [Serializable] объектов
    public static class ObjectCloner
    {
        public static T DeepClone<T>(this T obj) where T : class
        {
            if (obj == null)
                return null;

            var bf = new BinaryFormatter();
            using (var stream = new MemoryStream())
            {
                bf.Serialize(stream, obj);
                stream.Position = 0;
                return (T)bf.Deserialize(stream);
            }
        }
    }
}