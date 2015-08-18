namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Deal with the memory management.
    /// </summary>
    public static class NspiHelper
    {
        /// <summary>
        /// Allocate memory for the specific property values.
        /// </summary>
        /// <param name="propertyValue">PropertyValue_r instance.</param>
        /// <returns>A pointer points to memory allocated.</returns>
        public static IntPtr AllocPropertyValue_r(PropertyValue_r propertyValue)
        {
            IntPtr ptr = Marshal.AllocHGlobal(16);
            int offset = 0;
            Marshal.WriteInt32(ptr, (int)propertyValue.PropTag);
            offset += sizeof(uint);
            Marshal.WriteInt32(ptr, offset, (int)propertyValue.Reserved);
            offset += sizeof(uint);

            PropertyType proptype = (PropertyType)(0x0000FFFF & propertyValue.PropTag);
            switch (proptype)
            {
                case PropertyType.PtypInteger16:
                    {
                        Marshal.WriteInt16(ptr, offset, propertyValue.Value.I);
                        break;
                    }

                case PropertyType.PtypInteger32:
                    {
                        Marshal.WriteInt32(ptr, offset, propertyValue.Value.L);
                        break;
                    }

                case PropertyType.PtypBoolean:
                    {
                        Marshal.WriteInt16(ptr, offset, (short)propertyValue.Value.B);
                        break;
                    }

                case PropertyType.PtypString8:
                    {
                        IntPtr strA = Marshal.StringToHGlobalAnsi(propertyValue.Value.LpszA);
                        Marshal.WriteInt32(ptr, offset, strA.ToInt32());
                        break;
                    }

                case PropertyType.PtypBinary:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.Bin.Cb);
                        offset += sizeof(uint);
                        IntPtr lpb = Marshal.AllocHGlobal((int)propertyValue.Value.Bin.Cb);
                        for (int i = 0; i < propertyValue.Value.Bin.Cb; i++)
                        {
                            Marshal.WriteByte(lpb, i, propertyValue.Value.Bin.Lpb[i]);
                        }
                        
                        Marshal.WriteInt32(ptr, offset, lpb.ToInt32());
                        break;
                    }

                case PropertyType.PtypString:
                    {
                        IntPtr strW = Marshal.StringToHGlobalAnsi(propertyValue.Value.LpszW);
                        Marshal.WriteInt32(ptr, offset, strW.ToInt32());
                        break;
                    }

                case PropertyType.PtypGuid:
                    {
                        IntPtr guid = Marshal.AllocHGlobal(16);
                        for (int i = 0; i < 16; i++)
                        {
                            Marshal.WriteByte(guid, i, propertyValue.Value.Guid[0].Ab[i]);
                        }
                        
                        Marshal.WriteInt32(ptr, offset, guid.ToInt32());
                        break;
                    }

                case PropertyType.PtypTime:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.FileTime.LowDateTime);
                        offset += sizeof(uint);
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.FileTime.HighDateTime);
                        break;
                    }

                case PropertyType.PtypErrorCode:
                    {
                        Marshal.WriteInt32(ptr, offset, propertyValue.Value.ErrorCode);
                        break;
                    }

                case PropertyType.PtypMultipleInteger16:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVi.Values);
                        offset += sizeof(uint);
                        IntPtr lpi = Marshal.AllocHGlobal((int)(propertyValue.Value.MVi.Values * sizeof(short)));
                        for (int i = 0; i < propertyValue.Value.MVi.Values; i++)
                        {
                            Marshal.WriteInt16(lpi, i * sizeof(short), propertyValue.Value.MVi.Lpi[i]);
                        }
                        
                        Marshal.WriteInt32(ptr, offset, lpi.ToInt32());
                        break;
                    }

                case PropertyType.PtypMultipleInteger32:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVl.Values);
                        offset += sizeof(uint);
                        IntPtr lpl = Marshal.AllocHGlobal((int)(propertyValue.Value.MVl.Values * sizeof(int)));
                        for (int i = 0; i < propertyValue.Value.MVl.Values; i++)
                        {
                            Marshal.WriteInt32(lpl, i * sizeof(int), propertyValue.Value.MVl.Lpl[i]);
                        }
                        
                        Marshal.WriteInt32(ptr, offset, lpl.ToInt32());
                        break;
                    }

                case PropertyType.PtypMultipleString8:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVszA.Values);
                        offset += sizeof(uint);
                        IntPtr lppszA = Marshal.AllocHGlobal((int)(propertyValue.Value.MVszA.Values * 4));
                        for (int i = 0; i < propertyValue.Value.MVszA.Values; i++)
                        {
                            Marshal.WriteInt32(lppszA, 4 * i, Marshal.StringToHGlobalAnsi(propertyValue.Value.MVszA.LppszA[i]).ToInt32());
                        }
                        
                        Marshal.WriteInt32(ptr, offset, lppszA.ToInt32());
                        break;
                    }

                case PropertyType.PtypMultipleBinary:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVbin.Values);
                        offset += sizeof(uint);
                        IntPtr mvbin = Marshal.AllocHGlobal((int)(propertyValue.Value.MVbin.Values * 8));
                        for (int i = 0; i < propertyValue.Value.MVbin.Values; i++)
                        {
                            Marshal.WriteInt32(mvbin, (int)propertyValue.Value.MVbin.Lpbin[0].Cb);
                            IntPtr lpb = Marshal.AllocHGlobal((int)propertyValue.Value.MVbin.Lpbin[0].Cb);
                            for (int j = 0; j < propertyValue.Value.MVbin.Lpbin[0].Cb; j++)
                            {
                                Marshal.WriteByte(lpb, j, propertyValue.Value.MVbin.Lpbin[0].Lpb[j]);
                            }
                            
                            Marshal.WriteInt32(mvbin, sizeof(uint), lpb.ToInt32());
                            mvbin = new IntPtr(mvbin.ToInt32() + 8);
                        }
                        
                        Marshal.WriteInt32(ptr, offset, mvbin.ToInt32());
                        break;
                    }

                case PropertyType.PtypMultipleString:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVszW.Values);
                        offset += sizeof(uint);
                        IntPtr lppszW = Marshal.AllocHGlobal((int)(propertyValue.Value.MVszW.Values * 4));
                        for (int i = 0; i < propertyValue.Value.MVszW.Values; i++)
                        {
                            Marshal.WriteInt32(lppszW, 4 * i, Marshal.StringToHGlobalUni(propertyValue.Value.MVszW.LppszW[i]).ToInt32());
                        }
                        
                        Marshal.WriteInt32(ptr, offset, lppszW.ToInt32());
                        break;
                    }

                case PropertyType.PtypMultipleGuid:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVguid.Values);
                        offset += sizeof(uint);
                        IntPtr lpguid = Marshal.AllocHGlobal((int)(propertyValue.Value.MVguid.Values * 4));
                        for (int i = 0; i < propertyValue.Value.MVguid.Values; i++)
                        {
                            IntPtr guid = Marshal.AllocHGlobal(16);
                            for (int j = 0; j < 16; j++)
                            {
                                Marshal.WriteByte(guid, j, propertyValue.Value.MVguid.Guid[i].Ab[j]);
                            }
                            
                            Marshal.WriteInt32(lpguid, 4 * i, guid.ToInt32());
                        }
                        
                        Marshal.WriteInt32(ptr, offset, lpguid.ToInt32());
                        break;
                    }

                case PropertyType.PtypNull:
                case PropertyType.PtypEmbeddedTable:
                    {
                        Marshal.WriteInt32(ptr, offset, propertyValue.Value.Reserved);
                        break;
                    }
            }

            return ptr;
        }

        /// <summary>
        /// Free memory previously allocated for the specific property values.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        public static void FreePropertyValue_r(IntPtr ptr)
        {
            int offset = 0;

            uint propTag = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(uint) * 2;

            PropertyType proptype = (PropertyType)(0x0000FFFF & propTag);

            switch (proptype)
            {
                case PropertyType.PtypString8:
                case PropertyType.PtypGuid:
                    {
                        Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(ptr, offset)));
                        break;
                    }

                case PropertyType.PtypString:
                    {
                        Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(ptr, offset)));
                        break;
                    }

                case PropertyType.PtypBinary:
                case PropertyType.PtypMultipleInteger16:
                case PropertyType.PtypMultipleInteger32:
                    {
                        offset += sizeof(uint);
                        Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(ptr, offset)));
                        break;
                    }

                case PropertyType.PtypMultipleString8:
                case PropertyType.PtypMultipleString:
                    {
                        int values = Marshal.ReadInt32(ptr, offset);
                        offset += sizeof(uint);
                        IntPtr lppsz = new IntPtr(Marshal.ReadInt32(ptr, offset));
                        int offset2 = 0;
                        for (int i = 0; i < values; i++)
                        {
                            Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(lppsz, offset2)));
                            offset2 += 4;
                        }
                        
                        Marshal.FreeHGlobal(lppsz);
                        break;
                    }

                case PropertyType.PtypMultipleBinary:
                    {
                        int values = Marshal.ReadInt32(ptr, offset);
                        offset += sizeof(uint);
                        IntPtr mvbin = new IntPtr(Marshal.ReadInt32(ptr, offset));
                        int offset2 = 4;
                        for (int i = 0; i < values; i++)
                        {
                            Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(mvbin, offset2)));
                            offset2 += 8;
                        }
                        
                        Marshal.FreeHGlobal(mvbin);
                        break;
                    }

                case PropertyType.PtypMultipleGuid:
                    {
                        int values = Marshal.ReadInt32(ptr, offset);
                        offset += 4;
                        IntPtr mvguid = new IntPtr(Marshal.ReadInt32(ptr, offset));
                        int offset2 = 0;
                        for (int i = 0; i < values; i++)
                        {
                            Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(mvguid, offset2)));
                            offset2 += 4;
                        }
                        
                        Marshal.FreeHGlobal(mvguid);
                        break;
                    }
            }

            Marshal.FreeHGlobal(ptr);
        }

        /// <summary>
        /// Allocate memory for the specific property values.
        /// </summary>
        /// <param name="pta_r">PropertyTagArray_r instance.</param>
        /// <returns>A pointer points to memory allocated.</returns>        
        public static IntPtr AllocPropertyTagArray_r(PropertyTagArray_r pta_r)
        {
            int offset = 0;
            int cb = (int)(sizeof(uint) + (pta_r.Values * sizeof(uint)));

            IntPtr ptr = Marshal.AllocHGlobal(cb);

            Marshal.WriteInt32(ptr, offset, (int)pta_r.Values);
            offset += sizeof(uint);

            for (int i = 0; i < pta_r.Values; i++)
            {
                Marshal.WriteInt32(ptr, offset, (int)pta_r.AulPropTag[i]);
                offset += sizeof(uint);
            }

            return ptr;
        }

        /// <summary>
        /// Free memory previously allocated for the specific property values.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        public static void FreePropertyTagArray_r(IntPtr ptr)
        {
            Marshal.FreeHGlobal(ptr);
        }

        /// <summary>
        /// Allocate memory for the specific property values.
        /// </summary>
        /// <param name="propertyValues">Array of PropertyValue_r instances.</param>
        /// <returns>A pointer points to memory allocated.</returns>
        public static IntPtr AllocPropertyValue_rs(PropertyValue_r[] propertyValues)
        {
            IntPtr ptr = Marshal.AllocHGlobal(16 * propertyValues.Length);

            for (int k = 0; k < propertyValues.Length; k++)
            {
                int offset = 0;
                Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].PropTag);
                offset += sizeof(uint);
                Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Reserved);
                offset += sizeof(uint);

                PropertyType proptype = (PropertyType)(0x0000FFFF & propertyValues[k].PropTag);
                switch (proptype)
                {
                    case PropertyType.PtypInteger16:
                        {
                            Marshal.WriteInt16(ptr, (16 * k) + offset, propertyValues[k].Value.I);
                            break;
                        }

                    case PropertyType.PtypInteger32:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, propertyValues[k].Value.L);
                            break;
                        }

                    case PropertyType.PtypBoolean:
                        {
                            Marshal.WriteInt16(ptr, (16 * k) + offset, (short)propertyValues[k].Value.B);
                            break;
                        }

                    case PropertyType.PtypString8:
                        {
                            IntPtr strA = Marshal.StringToHGlobalAnsi(propertyValues[k].Value.LpszA);
                            Marshal.WriteInt32(ptr, (16 * k) + offset, strA.ToInt32());
                            break;
                        }

                    case PropertyType.PtypBinary:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.Bin.Cb);
                            offset += sizeof(uint);
                            IntPtr lpb = Marshal.AllocHGlobal((int)propertyValues[k].Value.Bin.Cb);
                            for (int i = 0; i < propertyValues[k].Value.Bin.Cb; i++)
                            {
                                Marshal.WriteByte(lpb, i, propertyValues[k].Value.Bin.Lpb[i]);
                            }
                            
                            Marshal.WriteInt32(ptr, (16 * k) + offset, lpb.ToInt32());
                            break;
                        }

                    case PropertyType.PtypString:
                        {
                            IntPtr strW = Marshal.StringToHGlobalAnsi(propertyValues[k].Value.LpszW);
                            Marshal.WriteInt32(ptr, (16 * k) + offset, strW.ToInt32());
                            break;
                        }

                    case PropertyType.PtypGuid:
                        {
                            IntPtr guid = Marshal.AllocHGlobal(16);
                            for (int i = 0; i < 16; i++)
                            {
                                Marshal.WriteByte(guid, i, propertyValues[k].Value.Guid[0].Ab[i]);
                            }
                            
                            Marshal.WriteInt32(ptr, (16 * k) + offset, guid.ToInt32());
                            break;
                        }

                    case PropertyType.PtypTime:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.FileTime.LowDateTime);
                            offset += sizeof(uint);
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.FileTime.HighDateTime);
                            break;
                        }

                    case PropertyType.PtypErrorCode:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, propertyValues[k].Value.ErrorCode);
                            break;
                        }

                    case PropertyType.PtypMultipleInteger16:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVi.Values);
                            offset += sizeof(uint);
                            IntPtr lpi = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVi.Values * sizeof(short)));
                            for (int i = 0; i < propertyValues[k].Value.MVi.Values; i++)
                            {
                                Marshal.WriteInt16(lpi, i * sizeof(short), propertyValues[k].Value.MVi.Lpi[i]);
                            }
                            
                            Marshal.WriteInt32(ptr, (16 * k) + offset, lpi.ToInt32());
                            break;
                        }

                    case PropertyType.PtypMultipleInteger32:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVl.Values);
                            offset += sizeof(uint);
                            IntPtr lpl = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVl.Values * sizeof(int)));
                            for (int i = 0; i < propertyValues[k].Value.MVl.Values; i++)
                            {
                                Marshal.WriteInt32(lpl, i * sizeof(int), propertyValues[k].Value.MVl.Lpl[i]);
                            }
                            
                            Marshal.WriteInt32(ptr, (16 * k) + offset, lpl.ToInt32());
                            break;
                        }

                    case PropertyType.PtypMultipleString8:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVszA.Values);
                            offset += sizeof(uint);
                            IntPtr lppszA = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVszA.Values * 4));
                            for (int i = 0; i < propertyValues[k].Value.MVszA.Values; i++)
                            {
                                Marshal.WriteInt32(lppszA, 4 * i, Marshal.StringToHGlobalAnsi(propertyValues[k].Value.MVszA.LppszA[i]).ToInt32());
                            }
                            
                            Marshal.WriteInt32(ptr, (16 * k) + offset, lppszA.ToInt32());
                            break;
                        }

                    case PropertyType.PtypMultipleBinary:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVbin.Values);
                            offset += sizeof(uint);
                            IntPtr mvbin = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVbin.Values * 8));
                            for (int i = 0; i < propertyValues[k].Value.MVbin.Values; i++)
                            {
                                Marshal.WriteInt32(mvbin, (int)propertyValues[k].Value.MVbin.Lpbin[0].Cb);
                                IntPtr lpb = Marshal.AllocHGlobal((int)propertyValues[k].Value.MVbin.Lpbin[0].Cb);
                                for (int j = 0; j < propertyValues[k].Value.MVbin.Lpbin[0].Cb; j++)
                                {
                                    Marshal.WriteByte(lpb, j, propertyValues[k].Value.MVbin.Lpbin[0].Lpb[j]);
                                }
                                
                                Marshal.WriteInt32(mvbin, sizeof(uint), lpb.ToInt32());
                                mvbin = new IntPtr(mvbin.ToInt32() + 8);
                            }
                            
                            Marshal.WriteInt32(ptr, (16 * k) + offset, mvbin.ToInt32());
                            break;
                        }

                    case PropertyType.PtypMultipleString:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVszW.Values);
                            offset += sizeof(uint);
                            IntPtr lppszW = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVszW.Values * 4));
                            for (int i = 0; i < propertyValues[k].Value.MVszW.Values; i++)
                            {
                                Marshal.WriteInt32(lppszW, 4 * i, Marshal.StringToHGlobalUni(propertyValues[k].Value.MVszW.LppszW[i]).ToInt32());
                            }
                            
                            Marshal.WriteInt32(ptr, (16 * k) + offset, lppszW.ToInt32());
                            break;
                        }

                    case PropertyType.PtypMultipleGuid:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVguid.Values);
                            offset += sizeof(uint);
                            IntPtr lpguid = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVguid.Values * 4));
                            for (int i = 0; i < propertyValues[k].Value.MVguid.Values; i++)
                            {
                                IntPtr guid = Marshal.AllocHGlobal(16);
                                for (int j = 0; j < 16; j++)
                                {
                                    Marshal.WriteByte(guid, j, propertyValues[k].Value.MVguid.Guid[i].Ab[j]);
                                }
                                
                                Marshal.WriteInt32(lpguid, 4 * i, guid.ToInt32());
                            }
                            
                            Marshal.WriteInt32(ptr, (16 * k) + offset, lpguid.ToInt32());
                            break;
                        }

                    case PropertyType.PtypNull:
                    case PropertyType.PtypEmbeddedTable:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, propertyValues[k].Value.Reserved);
                            break;
                        }
                }
            }

            return ptr;
        }

        /// <summary>
        /// Free memory previously allocated for the specific property values.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        /// <param name="count">The number of PropertyValue_r.</param>
        public static void FreePropertyValue_rs(IntPtr ptr, int count)
        {
            const int PropertyValueLengthInBytes = 16;

            for (int k = 0; k < count; k++)
            {
                int offset = 0;

                uint propTag = (uint)Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset);
                offset += sizeof(uint) * 2;

                PropertyType proptype = (PropertyType)(0x0000FFFF & propTag);

                switch (proptype)
                {
                    case PropertyType.PtypString8:
                    case PropertyType.PtypGuid:
                        {
                            Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset)));
                            break;
                        }

                    case PropertyType.PtypString:
                        {
                            Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset)));
                            break;
                        }

                    case PropertyType.PtypBinary:
                    case PropertyType.PtypMultipleInteger16:
                    case PropertyType.PtypMultipleInteger32:
                        {
                            offset += sizeof(uint);
                            Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset)));
                            break;
                        }

                    case PropertyType.PtypMultipleString8:
                    case PropertyType.PtypMultipleString:
                        {
                            int values = Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset);
                            offset += sizeof(uint);
                            IntPtr lppsz = new IntPtr(Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset));
                            int offset2 = 0;
                            for (int i = 0; i < values; i++)
                            {
                                Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(lppsz, offset2)));
                                offset2 += 4;
                            }
                            
                            Marshal.FreeHGlobal(lppsz);
                            break;
                        }

                    case PropertyType.PtypMultipleBinary:
                        {
                            int values = Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset);
                            offset += sizeof(uint);
                            IntPtr mvbin = new IntPtr(Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset));
                            int offset2 = 4;
                            for (int i = 0; i < values; i++)
                            {
                                Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(mvbin, offset2)));
                                offset2 += 8;
                            }
                            
                            Marshal.FreeHGlobal(mvbin);
                            break;
                        }

                    case PropertyType.PtypMultipleGuid:
                        {
                            int values = Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset);
                            offset += 4;
                            IntPtr mvguid = new IntPtr(Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset));
                            int offset2 = 0;
                            for (int i = 0; i < values; i++)
                            {
                                Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(mvguid, offset2)));
                                offset2 += 4;
                            }
                            
                            Marshal.FreeHGlobal(mvguid);
                            break;
                        }
                }
            }

            Marshal.FreeHGlobal(ptr);
        }

        /// <summary>
        /// Allocate memory for the Table row.
        /// </summary>
        /// <param name="propertyRow">Instance of table row structure.</param>
        /// <returns>A pointer points to memory allocated.</returns>
        public static IntPtr AllocPropertyRow_r(PropertyRow_r propertyRow)
        {
            IntPtr ptr = Marshal.AllocHGlobal(12);
            int offset = 0;
            Marshal.WriteInt32(ptr, offset, (int)propertyRow.Reserved);
            offset += sizeof(uint);
            Marshal.WriteInt32(ptr, offset, (int)propertyRow.Values);
            offset += sizeof(uint);

            IntPtr pv = AllocPropertyValue_rs(propertyRow.Props);

            Marshal.WriteInt32(ptr, offset, pv.ToInt32());

            return ptr;
        }

        /// <summary>
        /// Free memory previously allocated for the Table row.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        public static void FreePropertyRow_r(IntPtr ptr)
        {
            int offset = sizeof(uint);
            int count = Marshal.ReadInt32(ptr, offset);
            offset += sizeof(uint);
            IntPtr propsHandle = new IntPtr(Marshal.ReadInt32(ptr, offset));

            FreePropertyValue_rs(propsHandle, count);

            Marshal.FreeHGlobal(ptr);
        }

        /// <summary>
        /// Allocate memory for GUIDs.
        /// </summary>
        /// <param name="fuid_r">Instance of GUIDs.</param>
        /// <returns>A pointer points to memory allocated.</returns>
        public static IntPtr AllocFlatUID_r(FlatUID_r fuid_r)
        {
            IntPtr ptr = Marshal.AllocHGlobal(16);

            for (int i = 0; i < 16; i++)
            {
                Marshal.WriteByte(ptr, i, fuid_r.Ab[i]);
            }

            return ptr;
        }

        /// <summary>
        /// Free memory previously allocated for GUIDs.
        /// </summary>
        /// <param name="ptr">A pointer points to memory allocated.</param>
        public static void FreeFlatUID_r(IntPtr ptr)
        {
            Marshal.FreeHGlobal(ptr);
        }
    }
}