namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Reflection;

    /// <summary>
    /// Supply help functions for getting FieldInfos.
    /// </summary>
    public static class FieldHelper
    {
        /// <summary>
        /// Get FieldInfos from an object.
        /// </summary>
        /// <param name="type">The type of the object.</param>
        /// <returns>A list contains the object's FieldInfos.</returns>
        public static List<FieldInfo> FilterFields(Type type)
        {
            List<FieldInfo> list = new List<FieldInfo>();
            while (type != typeof(SerializableBase))
            {
                FieldInfo[] fields = type.GetFields(
                        BindingFlags.Instance
                        | BindingFlags.NonPublic
                        | BindingFlags.Public
                        | BindingFlags.DeclaredOnly);
                List<FieldInfo> filtedFields = FilterFields(fields);
                for (int i = filtedFields.Count - 1; i >= 0; i--)
                {
                    list.Insert(0, filtedFields[i]);
                }

                type = type.BaseType;
            }

            return list;
        }

        /// <summary>
        /// Sort fields.
        /// </summary>
        /// <param name="fields">An array of FieldInfo.</param>
        /// <returns>A sorted fields list.</returns>
        public static List<FieldInfo> FilterFields(FieldInfo[] fields)
        {
            int i;
            List<FieldInfo> fieldList = new List<FieldInfo>();
            for (i = 0; i < fields.Length; i++)
            {
                object[] attr = fields[i].GetCustomAttributes(typeof(SerializableFieldAttribute), false);
                if (attr.Length > 0)
                {
                    Debug.Assert(attr.Length == 1, "The value must be 1.");
                    fieldList.Add(fields[i]);
                }
            }

            Comparison<FieldInfo> comp = new Comparison<FieldInfo>(FieldComparison);
            fieldList.Sort(comp);
            return fieldList;
        }

        /// <summary>
        /// Get a FieldInfo's SerializableField Attribute.
        /// </summary>
        /// <param name="fi">A FieldInfo.</param>
        /// <returns>The first SerializableField attribute of the FieldInfo.</returns>
        public static SerializableFieldAttribute GetSerializableField(FieldInfo fi)
        {
            return fi.GetCustomAttributes(typeof(SerializableFieldAttribute), false)[0]
                as SerializableFieldAttribute;
        }

        /// <summary>
        /// Compare two FieldInfos.
        /// </summary>
        /// <param name="f1">The 1st FieldInfo.</param>
        /// <param name="f2">The 2nd FieldInfo.</param>
        /// <returns>If the 1st FieldInfo's SerializableField' order is greater 
        /// than the 2nd FieldInfos's, return value greater than 0.
        /// If they are equal, return 0.
        /// Else return value less than 0.
        /// </returns>
        public static int FieldComparison(FieldInfo f1, FieldInfo f2)
        {
            SerializableFieldAttribute o1 = f1.GetCustomAttributes(typeof(SerializableFieldAttribute), false)[0]
                as SerializableFieldAttribute;
            SerializableFieldAttribute o2 = f2.GetCustomAttributes(typeof(SerializableFieldAttribute), false)[0]
                as SerializableFieldAttribute;

            return o1.Order - o2.Order;
        }

        /// <summary>
        /// Get the first Field Attribute of type T.
        /// </summary>
        /// <typeparam name="T">An Attribute type.</typeparam>
        /// <param name="fi">A FieldInfo.</param>
        /// <returns>The FieldInfo's first T Attribute.</returns>
        public static T GetFieldAttribute<T>(FieldInfo fi)
            where T : Attribute
        {
            return fi.GetCustomAttributes(typeof(T), false)[0]
                as T;
        }

        /// <summary>
        /// Get the First Custom Attribute.
        /// </summary>
        /// <typeparam name="T">An Attribute type.</typeparam>
        /// <param name="obj">The object.</param>
        /// <param name="inherit">look up inherited Custom attributes.</param>
        /// <returns>The FieldInfo's first T Attribute.</returns>
        public static T GetFirstCustomAttribute<T>(object obj, bool inherit)
            where T : Attribute
        {
            Type type = obj.GetType();
            object[] atts = type.GetCustomAttributes(typeof(T), inherit);
            if (atts.Length > 0)
            {
                return atts[0] as T;
            }
            else
            {
                return null;
            }
        }
    }
}