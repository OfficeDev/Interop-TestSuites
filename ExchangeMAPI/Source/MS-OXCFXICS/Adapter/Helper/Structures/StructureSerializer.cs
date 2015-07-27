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
    using System.IO;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Serialize structures.
    /// </summary>
    public static class StructureSerializer
    {
        /// <summary>
        /// Deserialize bytes to structure.
        /// </summary>
        /// <typeparam name="T">Structure type.</typeparam>
        /// <param name="bytes">The bytes to be deserialized.</param>
        /// <returns>Structure deserialized from bytes.</returns>
        public static T Deserialize<T>(byte[] bytes)
            where T : struct
        {
            Type type = typeof(T);
            int size = Marshal.SizeOf(type);
            if (size > bytes.Length)
            {
                AdapterHelper.Site.Assert.Fail(string.Format("The length of bytes to be deserialized doesn't match the size of an instance of type {0}.", type.FullName));
            }

            IntPtr p = Marshal.AllocHGlobal(size);
            Marshal.Copy(bytes, 0, p, size);
            T obj = (T)Marshal.PtrToStructure(p, type);
            Marshal.FreeHGlobal(p);
            return obj;
        }

        /// <summary>
        /// Serialize an object to bytes.
        /// </summary>
        /// <param name="obj">The object to be serialized.</param>
        /// <returns>Serialized object.</returns>
        public static byte[] Serialize(object obj)
        { 
            Type type = obj.GetType();
            int size = Marshal.SizeOf(type);   
            byte[] buffer = new byte[size];
            IntPtr tmp = Marshal.AllocHGlobal(size);
            Marshal.StructureToPtr(obj, tmp, false);
            Marshal.Copy(tmp, buffer, 0, size);
            Marshal.FreeHGlobal(tmp);
            return buffer;
        }

        /// <summary>
        /// Deserialize bytes to structure.
        /// </summary>
        /// <typeparam name="T">Structure type.</typeparam>
        /// <param name="stream">The stream.</param>
        /// <returns>A structure deserialized from bytes.</returns>
        public static T Deserialize<T>(Stream stream)
            where T : struct
        {
            Type type = typeof(T);
            int size = Marshal.SizeOf(type);
            byte[] buffer = new byte[size];
            stream.Read(buffer, 0, size);
            return Deserialize<T>(buffer);
        }

        /// <summary>
        /// Serialize structures to the data stream.
        /// </summary>
        /// <param name="obj">Object to be serialized.</param>
        /// <param name="stream">Data stream contains serialized object.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public static int Serialize(object obj, Stream stream)
        {
            byte[] buffer = Serialize(obj);
            stream.Write(buffer, 0, buffer.Length);
            return buffer.Length;
        }
    }
}