//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to deserialize input bytes into an ROP object
    /// </summary>
    public sealed class RopDeserializer
    {
        /// <summary>
        /// A dictionary used to store the ROP's Id and related Deserializer
        /// </summary>
        private static Dictionary<int, IDeserializable> maps;

        /// <summary>
        /// Prevents a default instance of the <see cref="RopDeserializer" /> class from being created.
        /// </summary>
        private RopDeserializer()
        {
        }

        /// <summary>
        /// Initialize and allocate memory for the internal dictionary
        /// </summary>
        public static void Init()
        {
            maps = new Dictionary<int, IDeserializable>();
        }

        /// <summary>
        /// Register a ROP's Id together with a Deserializer
        /// </summary>
        /// <param name="ropId">The ROP's Id</param>
        /// <param name="iropDeserializer">The interface define the methods that is needed to deserialize a bytes array into an ROP object</param>
        public static void Register(int ropId, IDeserializable iropDeserializer)
        {
            maps[ropId] = iropDeserializer;
        }

        /// <summary>
        /// Deserialize input bytes indicated by ropBytes into a list of ROPs indicated by ropList
        /// </summary>
        /// <param name="ropBytes">The bytes need to be deserialized</param>
        /// <param name="ropList">The ROPs list deserialized into</param>
        public static void Deserialize(byte[] ropBytes, ref List<IDeserializable> ropList)
        {
            int index = 0;
            while (index < ropBytes.Length)
            {
                int ropId = ropBytes[index];
                bool isRopId = false;
                foreach (int id in maps.Keys)
                {
                    if (ropId == id)
                    {
                        isRopId = true;
                    }
                }

                if (isRopId)
                {
                    // Deserialize to a rop
                    // Renew every response before deserialize.
                    Type ropType = maps[ropId].GetType();
                    maps[ropId] = (IDeserializable)ropType.InvokeMember("new", System.Reflection.BindingFlags.CreateInstance, null, null, null);
                    int desBytes = maps[ropId].Deserialize(ropBytes, index);

                    // Add deserialized rop into rop list
                    ropList.Add(maps[ropId]);
                    index += desBytes;
                }
                else
                {
                    throw new Exception("Unknown response rop id: " + ropId.ToString());
                }
            }
        }

        /// <summary>
        /// Deserialize input bytes indicated by ropBytes into a list of ROPs indicated by ropList
        /// </summary>
        /// <param name="ropBytes">The bytes need to be deserialized</param>
        /// <param name="ropList">The ROPs list deserialized into</param>
        public static void Deserialize(byte[] ropBytes, ref IDeserializable ropList)
        {
            int index = 0;
            if (index < ropBytes.Length)
            {
                int ropId = ropBytes[index];

                // Deserialize to a rop
                // Renew every response before deserialize.
                Type type  = maps[ropId].GetType();
                maps[ropId] = (IDeserializable)type.InvokeMember("new", System.Reflection.BindingFlags.CreateInstance, null, null, null);

                int desBytes = maps[ropId].Deserialize(ropBytes, index);

                // Add deserialized rop into rop list
                ropList = maps[ropId];
                index += desBytes;
            }
        }
    }
}