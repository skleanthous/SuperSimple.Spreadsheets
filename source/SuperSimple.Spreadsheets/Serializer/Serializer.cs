using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace SSExcel.Serializer
{
    internal class SerializerToExcelRow : ISerializerToExcelRow
    {
        public IEnumerable<ExcelRow> Serialize<T>(IEnumerable<T> itemsToSerialize, bool getHeaders = true)
        {
            var type = typeof(T);

            var properties = GetProperties(type);
            var fields = GetFields(type);

            if(getHeaders)
            {
                yield return new ExcelRow(GetTitles(properties, fields).ToArray());
            }

            foreach(var itemToSerialize in itemsToSerialize)
            {
                yield return new ExcelRow(GetValues(properties, fields, itemsToSerialize));
            }
        }

        private static IEnumerable<string> GetTitles(PropertyInfo[] properties, FieldInfo[] fields)
        {
            foreach (var property in properties)
            {
                yield return property.Name;
            }

            foreach (var field in fields)
            {
                yield return field.Name;
            }
        }

        private static FieldInfo[] GetFields(Type type)
        {
            return type.GetRuntimeFields()
                .ToArray();
        }

        private static PropertyInfo[] GetProperties(Type type)
        {
            return type.GetRuntimeProperties()//.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance)
                .Where(p => p.CanRead)
                .ToArray();
        }

        private static IEnumerable<object> GetValues(PropertyInfo[] properties, FieldInfo[] fields, object itemToSerialize)
        {
            for(int i = 0;i<properties.Length; i++)
            {
                yield return properties[i].GetValue(itemToSerialize, null) ?? "";
            }

            for(int i = 0;i<fields.Length; i++)
            {
                yield return fields[i].GetValue(itemToSerialize) ?? "";
            }
        }
    }
}
