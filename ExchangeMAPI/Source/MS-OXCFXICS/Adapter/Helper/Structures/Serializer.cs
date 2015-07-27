//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Contains method to serialize a serializable object.
    /// </summary>
    public class Serializer
    {
        /// <summary>
        /// Serialize a serializable object
        /// </summary>
        /// <param name="obj">An object instance must have a SerializableObject 
        /// attribute.
        /// </param>
        /// <param name="stream">The Stream.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public static int Serialize(object obj, Stream stream)
        {
            if (obj == null)
            {
                return -1;
            }

            if (SerializableObjectAttribute.IsSerializableObject(obj))
            { 
                SerializableObjectAttribute att = SerializableObjectAttribute.GetSerializableObject(obj);
                if (att.UseSelfSerialize)
                {
                    AdapterHelper.Site.Assert.Fail("The object 'obj' is not supported.");
                }
            }

            return Serialize(obj, -1, -1, null, stream);
        }

        /// <summary>
        /// Fill the stream.
        /// </summary>
        /// <param name="val">The parameter to fill.</param>
        /// <param name="min">The min must be -1.</param>
        /// <param name="max">The max must be -1.</param>
        /// <param name="filledValue">This parameter must be null.</param>
        /// <param name="stream">The stream to be filled.</param>
        /// <returns>The size have been filled.</returns>
        private static int FillStream(
            byte[] val,
            int min,
            int max,
            byte[] filledValue,
            Stream stream)
        {
            if (min < 0 && max < 0 && filledValue == null)
            {
                if (val != null)
                {
                    stream.Write(val, 0, val.Length);
                    int size = val.Length;
                    return size;
                }
            }

            AdapterHelper.Site.Assert.Fail("Method is not implemented.");
            return -1;
        }

        /// <summary>
        /// Serialize the object to the stream.
        /// </summary>
        /// <param name="obj">The object to be Serialize</param>
        /// <param name="min">The min must be -1.</param>
        /// <param name="max">The max must be -1.</param>
        /// <param name="fillValue">This parameter must be null.</param>
        /// <param name="stream">The stream to serialize.</param>
        /// <returns>The size have been Serialize.</returns>
        private static int Serialize(object obj, int min, int max, byte[] fillValue, Stream stream)
        {
            int size = 0;
            int i;
            if (obj == null)
            {
                size += FillStream(null, min, max, fillValue, stream);
            }
            else if (SerializableObjectAttribute.IsSerializableObject(obj))
            {
                SerializableObjectAttribute att = SerializableObjectAttribute.GetSerializableObject(obj);
                if (att.UseSelfSerialize)
                {
                    IStructSerializable serial = obj as IStructSerializable;
                    size += serial.Serialize(stream);
                }
                else
                {
                    List<FieldInfo> fields = FieldHelper.FilterFields(obj.GetType());
                    for (i = 0; i < fields.Count; i++)
                    {
                        SerializableFieldAttribute sf = FieldHelper.GetSerializableField(fields[i]);
                        size += Serialize(
                            fields[i].GetValue(obj),
                            sf.MinAllocSize,
                            sf.MaxAllocSize,
                            null,
                            stream);
                    }
                }
            }
            else if (obj is IList)
            {
                IList lst = obj as IList;
                for (i = 0; i < lst.Count; i++)
                {
                    size += Serialize(lst[i], min, max, fillValue, stream);
                }
            }
            else if (obj.GetType().IsValueType)
            {
                int bufferSize = Marshal.SizeOf(obj);
                byte[] buffer = new byte[bufferSize];
                IntPtr p = Marshal.AllocHGlobal(bufferSize);
                Marshal.StructureToPtr(obj, p, false);
                Marshal.Copy(p, buffer, 0, bufferSize);
                size += FillStream(buffer, min, max, fillValue, stream);
                Marshal.FreeHGlobal(p);
            }
            else if (obj is string)
            {
                AdapterHelper.Site.Assert.Fail("Method is not implemented.");
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("This type of 'obj' is not supported, its type is {0}.", obj.GetType().Name);
            }

            return size;
        }
    }
}