namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Net;
    using System.Runtime.InteropServices;
    using System.Text.RegularExpressions;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to assist MS-OXNSPIAdapter.
    /// </summary>
    public static class AdapterHelper
    {
        #region Variables

        /// <summary>
        /// The current session context cookies for the request.
        /// </summary>
        private static CookieCollection sessionContextCookies;

        /// <summary>
        /// The site which is used to print log information.
        /// </summary>
        private static ITestSite site;

        /// <summary>
        /// The transport used by the test suite.
        /// </summary>
        private static string transport;

        /// <summary>
        /// Gets or sets the instance of site.
        /// </summary>
        public static ITestSite Site
        {
            get { return AdapterHelper.site; }
            set { AdapterHelper.site = value; }
        }

        /// <summary>
        /// Gets or sets the transport used by the test suite.
        /// </summary>
        public static string Transport
        {
            get { return AdapterHelper.transport; }
            set { AdapterHelper.transport = value; }
        }

        /// <summary>
        /// Gets or sets the current session context cookies for the request.
        /// </summary>
        public static CookieCollection SessionContextCookies
        {
            get
            {
                if (sessionContextCookies == null)
                {
                    sessionContextCookies = new CookieCollection();
                }

                return sessionContextCookies;
            }

            set
            {
                sessionContextCookies = value;
            }
        }

        #endregion Variables

        #region Adapter Help Methods
        #region Marshal Methods
        /// <summary>
        /// Allocate memory for Restriction_r array instances.
        /// </summary>
        /// <param name="res_r">The BinaryArray_r instance.</param>
        /// <returns>The pointer points to the allocated memory.</returns>
        public static IntPtr AllocRestriction_rs(Restriction_r[] res_r)
        {
            // The size of Restriction_r is 16.
            const int Size = 16;
            IntPtr ptr = Marshal.AllocHGlobal(Size * res_r.Length);

            for (int k = 0; k < res_r.Length; k++)
            {
                int offset = 0;

                // Write rt field value.
                Marshal.WriteInt32(ptr, (Size * k) + offset, (int)res_r[k].Rt);
                offset += sizeof(int);

                switch (res_r[k].Rt)
                {
                    // AndRestriction_r
                    case 0x00000000:
                        Marshal.WriteInt32(ptr, (Size * k) + offset, (int)res_r[k].Res.ResAnd.CRes);
                        offset += sizeof(int);

                        IntPtr ptrAndRes = AllocRestriction_rs(res_r[k].Res.ResAnd.LpRes);
                        Marshal.WriteInt32(ptr, (Size * k) + offset, ptrAndRes.ToInt32());
                        break;

                    // OrRestriction_r
                    case 0x00000001:
                        Marshal.WriteInt32(ptr, (Size * k) + offset, (int)res_r[k].Res.ResOr.CRes);
                        offset += sizeof(int);

                        IntPtr ptrOrRes = AllocRestriction_rs(res_r[k].Res.ResOr.LpRes);
                        Marshal.WriteInt32(ptr, (Size * k) + offset, ptrOrRes.ToInt32());
                        break;

                    // NotRestriction_r
                    case 0x00000002:
                        IntPtr ptrNotRes = AllocRestriction_rs(res_r[k].Res.ResNot.LpRes);
                        Marshal.WriteInt32(ptr, (Size * k) + offset, ptrNotRes.ToInt32());
                        break;

                    // ContentRestriction_r
                    case 0x00000003:
                        Marshal.WriteInt32(ptr, (Size * k) + offset, (int)res_r[k].Res.ResContent.FuzzyLevel);
                        offset += sizeof(int);

                        Marshal.WriteInt32(ptr, (Size * k) + offset, (int)res_r[k].Res.ResContent.PropTag);
                        offset += sizeof(int);

                        IntPtr ptrContRes = AllocPropertyValue_rs(res_r[k].Res.ResContent.Prop);
                        Marshal.WriteInt32(ptr, (Size * k) + offset, ptrContRes.ToInt32());
                        break;

                    // PropertyRestriction_r
                    case 0x00000004:
                        Marshal.WriteInt32(ptr, (Size * k) + offset, (int)res_r[k].Res.ResProperty.Relop);
                        offset += sizeof(int);

                        Marshal.WriteInt32(ptr, (Size * k) + offset, (int)res_r[k].Res.ResProperty.PropTag);
                        offset += sizeof(int);

                        IntPtr ptrPropRes = AllocPropertyValue_rs(res_r[k].Res.ResProperty.Prop);
                        Marshal.WriteInt32(ptr, (Size * k) + offset, ptrPropRes.ToInt32());
                        break;

                    // ExistRestriction_r
                    case 0x00000008:
                        Marshal.WriteInt32(ptr, (Size * k) + offset, (int)res_r[k].Res.ResExist.Reserved1);
                        offset += sizeof(int);

                        Marshal.WriteInt32(ptr, (Size * k) + offset, (int)res_r[k].Res.ResExist.PropTag);
                        offset += sizeof(int);

                        Marshal.WriteInt32(ptr, (Size * k) + offset, (int)res_r[k].Res.ResExist.Reserved2);
                        offset += sizeof(int);
                        break;

                    default:
                        throw new ArgumentException("Unknown type of RestrictionUnion_r.");
                }
            }

            return ptr;
        }

        /// <summary>
        /// Free the allocated memory for Restriction_r array instances.
        /// </summary>
        /// <param name="ptr">The pointer points to the memory.</param>
        /// <param name="count">The count of Restriction_r instances to be free.</param>
        public static void FreeRestriction_rs(IntPtr ptr, int count)
        {
            // The size of Restriction_r is 16.
            const int Size = 16;

            for (int k = 0; k < count; k++)
            {
                int offset = 0;
                int rt = Marshal.ReadInt32(ptr, (Size * k) + offset);
                offset += sizeof(int);

                switch (rt)
                {
                    // AndRestriction_r
                    case 0x00000000:
                        int cresOfAndRes = Marshal.ReadInt32(ptr, (Size * k) + offset);
                        offset += sizeof(int);

                        IntPtr ptrAndRes = new IntPtr(Marshal.ReadInt32(ptr, (Size * k) + offset));
                        FreeRestriction_rs(ptrAndRes, cresOfAndRes);
                        break;

                    // OrRestriction_r
                    case 0x00000001:
                        int cresOfOrRes = Marshal.ReadInt32(ptr, (Size * k) + offset);
                        offset += sizeof(int);

                        IntPtr ptrOrRes = new IntPtr(Marshal.ReadInt32(ptr, (Size * k) + offset));
                        FreeRestriction_rs(ptrOrRes, cresOfOrRes);
                        break;

                    // NotRestriction_r
                    case 0x00000002:
                        IntPtr ptrNotRes = new IntPtr(Marshal.ReadInt32(ptr, (Size * k) + offset));

                        // NotRestriction has a Restriction structure.
                        FreeRestriction_rs(ptrNotRes, 1);
                        break;

                    // ContentRestriction_r
                    case 0x00000003:
                        Marshal.ReadInt32(ptr, (Size * k) + offset);
                        offset += sizeof(int);

                        Marshal.ReadInt32(ptr, (Size * k) + offset);
                        offset += sizeof(int);

                        IntPtr ptrContRes = new IntPtr(Marshal.ReadInt32(ptr, (Size * k) + offset));
                        FreePropertyValue_rs(ptrContRes, 1);
                        break;

                    // Propertyrestriction_r
                    case 0x00000004:
                        Marshal.ReadInt32(ptr, (Size * k) + offset);
                        offset += sizeof(int);

                        Marshal.ReadInt32(ptr, (Size * k) + offset);
                        offset += sizeof(int);

                        IntPtr ptrPropRes = new IntPtr(Marshal.ReadInt32(ptr, (Size * k) + offset));
                        FreePropertyValue_rs(ptrPropRes, 1);
                        break;

                    // ExistRestriction_r
                    case 0x00000008:
                        Marshal.ReadInt32(ptr, (Size * k) + offset);
                        offset += sizeof(int);

                        Marshal.ReadInt32(ptr, (Size * k) + offset);
                        offset += sizeof(int);

                        Marshal.ReadInt32(ptr, (Size * k) + offset);
                        offset += sizeof(int);
                        break;

                    default:
                        throw new ArgumentException("Unknown type of RestrictionUnion_r.");
                }
            }

            Marshal.FreeHGlobal(ptr);
        }

        /// <summary>
        /// Allocate memory for BinaryArray_r instance.
        /// </summary>
        /// <param name="binArr_r">The BinaryArray_r instance.</param>
        /// <returns>The pointer points to the allocated memory.</returns>
        public static IntPtr AllocBinaryArray_r(BinaryArray_r binArr_r)
        {
            IntPtr ptr = Marshal.AllocHGlobal(8);

            IntPtr brptr = Marshal.AllocHGlobal((int)binArr_r.CValues * 8);
            int offset = 0;

            for (uint i = 0; i < binArr_r.CValues; i++)
            {
                Marshal.WriteInt32(brptr, offset, (int)binArr_r.Lpbin[i].Cb);
                offset += 4;

                IntPtr tmp = Marshal.AllocHGlobal((int)binArr_r.Lpbin[i].Cb);

                for (uint j = 0; j < binArr_r.Lpbin[i].Cb; j++)
                {
                    Marshal.WriteByte(tmp, (int)j, binArr_r.Lpbin[i].Lpb[j]);
                }

                Marshal.WriteInt32(brptr, offset, tmp.ToInt32());
                offset += 4;
            }

            Marshal.WriteInt32(ptr, (int)binArr_r.CValues);
            Marshal.WriteInt32(ptr, 4, brptr.ToInt32());

            return ptr;
        }

        /// <summary>
        /// Free the allocated memory for BinaryArray_r instance.
        /// </summary>
        /// <param name="ptr">The pointer points to the allocated memory.</param>
        public static void FreeBinaryArray_r(IntPtr ptr)
        {
            int cvalues = Marshal.ReadInt32(ptr);
            IntPtr brptr = new IntPtr(Marshal.ReadInt32(ptr, 4));

            int offset = 4;

            for (int i = 0; i < cvalues; i++)
            {
                IntPtr tmp = new IntPtr(Marshal.ReadInt32(brptr, offset));
                Marshal.FreeHGlobal(tmp);
                offset += 8;
            }

            Marshal.FreeHGlobal(brptr);
            Marshal.FreeHGlobal(ptr);
        }

        /// <summary>
        /// Allocate memory for StringsArray_r instance.
        /// </summary>
        /// <param name="strArr_r">The StringsArray_r instance.</param>
        /// <returns>The pointer points to the allocated memory.</returns>
        public static IntPtr AllocStringsArray_r(StringsArray_r strArr_r)
        {
            int size = sizeof(uint) + (4 * (int)strArr_r.CValues);
            IntPtr ptr = Marshal.AllocHGlobal(size);
            int offset = 0;

            Marshal.WriteInt32(ptr, (int)strArr_r.CValues);
            offset += 4;

            for (uint i = 0; i < strArr_r.CValues; i++)
            {
                Marshal.WriteInt32(ptr, offset, Marshal.StringToHGlobalAnsi(strArr_r.LppszA[i]).ToInt32());
                offset += 4;
            }

            return ptr;
        }

        /// <summary>
        /// Free the allocated memory for StringsArray_r instance.
        /// </summary>
        /// <param name="ptr">The pointer points to the allocated memory.</param>
        public static void FreeStringsArray_r(IntPtr ptr)
        {
            int count = Marshal.ReadInt32(ptr);
            int offset = 4;

            for (int i = 0; i < count; i++)
            {
                Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(ptr, offset)));
                offset += 4;
            }

            Marshal.FreeHGlobal(ptr);
        }

        /// <summary>
        /// Allocate memory for WStringsArray_r instance.
        /// </summary>
        /// <param name="wsa_r">The WStringsArray_r instance.</param>
        /// <returns>The pointer points to the allocated memory.</returns>
        public static IntPtr AllocWStringsArray_r(WStringsArray_r wsa_r)
        {
            int size = sizeof(uint) + (4 * (int)wsa_r.CValues);
            IntPtr ptr = Marshal.AllocHGlobal(size);
            int offset = 0;

            Marshal.WriteInt32(ptr, (int)wsa_r.CValues);
            offset += 4;

            for (uint i = 0; i < wsa_r.CValues; i++, offset += 4)
            {
                Marshal.WriteInt32(ptr, offset, Marshal.StringToHGlobalUni(wsa_r.LppszW[i]).ToInt32());
            }

            return ptr;
        }

        /// <summary>
        /// Allocate memory for stat instance.
        /// </summary>
        /// <param name="stat">The stat instance.</param>
        /// <returns>The pointer points to the allocated memory.</returns>
        public static IntPtr AllocStat(STAT stat)
        {
            IntPtr ptr = Marshal.AllocHGlobal(Marshal.SizeOf(stat));
            int offset = 0;
            Marshal.WriteInt32(ptr, offset, (int)stat.SortType);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.ContainerID);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.CurrentRec);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, stat.Delta);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.NumPos);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.TotalRecs);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.CodePage);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.TemplateLocale);
            offset += sizeof(int);
            Marshal.WriteInt32(ptr, offset, (int)stat.SortLocale);

            return ptr;
        }

        /// <summary>
        /// Allocate memory for the specific property values.
        /// </summary>
        /// <param name="propertyValue">PropertyValue_r instance.</param>
        /// <returns>A pointer points to the allocated memory.</returns>
        public static IntPtr AllocPropertyValue_r(PropertyValue_r propertyValue)
        {
            IntPtr ptr = Marshal.AllocHGlobal(16);
            int offset = 0;
            Marshal.WriteInt32(ptr, (int)propertyValue.PropTag);
            offset += sizeof(uint);
            Marshal.WriteInt32(ptr, offset, (int)propertyValue.Reserved);
            offset += sizeof(uint);

            PropertyTypeValue proptype = (PropertyTypeValue)(0x0000FFFF & propertyValue.PropTag);
            switch (proptype)
            {
                case PropertyTypeValue.PtypInteger16:
                    {
                        Marshal.WriteInt16(ptr, offset, propertyValue.Value.I);
                        break;
                    }

                case PropertyTypeValue.PtypInteger32:
                    {
                        Marshal.WriteInt32(ptr, offset, propertyValue.Value.L);
                        break;
                    }

                case PropertyTypeValue.PtypBoolean:
                    {
                        Marshal.WriteInt16(ptr, offset, (short)propertyValue.Value.B);
                        break;
                    }

                case PropertyTypeValue.PtypString8:
                    {
                        IntPtr strA = IntPtr.Zero;

                        if (propertyValue.Value.LpszA != null)
                        {
                            string str = System.Text.Encoding.Default.GetString(propertyValue.Value.LpszA);
                            strA = Marshal.StringToHGlobalAnsi(str);
                        }

                        Marshal.WriteInt32(ptr, offset, strA.ToInt32());
                        break;
                    }

                case PropertyTypeValue.PtypBinary:
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

                case PropertyTypeValue.PtypString:
                    {
                        IntPtr strW = IntPtr.Zero;

                        if (propertyValue.Value.LpszW != null)
                        {
                            string str = System.Text.Encoding.Unicode.GetString(propertyValue.Value.LpszW);
                            strW = Marshal.StringToHGlobalUni(str);
                        }

                        Marshal.WriteInt32(ptr, offset, strW.ToInt32());
                        break;
                    }

                case PropertyTypeValue.PtypGuid:
                    {
                        IntPtr guid = Marshal.AllocHGlobal(16);

                        for (int i = 0; i < 16; i++)
                        {
                            Marshal.WriteByte(guid, i, propertyValue.Value.Lpguid[0].Ab[i]);
                        }

                        Marshal.WriteInt32(ptr, offset, guid.ToInt32());
                        break;
                    }

                case PropertyTypeValue.PtypTime:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.Ft.LowDateTime);
                        offset += sizeof(uint);
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.Ft.HighDateTime);
                        break;
                    }

                case PropertyTypeValue.PtypErrorCode:
                    {
                        Marshal.WriteInt32(ptr, offset, propertyValue.Value.Err);
                        break;
                    }

                case PropertyTypeValue.PtypMultipleInteger16:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVi.CValues);
                        offset += sizeof(uint);
                        IntPtr lpi = Marshal.AllocHGlobal((int)(propertyValue.Value.MVi.CValues * sizeof(short)));

                        for (int i = 0; i < propertyValue.Value.MVi.CValues; i++)
                        {
                            Marshal.WriteInt16(lpi, i * sizeof(short), propertyValue.Value.MVi.Lpi[i]);
                        }

                        Marshal.WriteInt32(ptr, offset, lpi.ToInt32());
                        break;
                    }

                case PropertyTypeValue.PtypMultipleInteger32:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVl.CValues);
                        offset += sizeof(uint);
                        IntPtr lpl = Marshal.AllocHGlobal((int)(propertyValue.Value.MVl.CValues * sizeof(int)));

                        for (int i = 0; i < propertyValue.Value.MVl.CValues; i++)
                        {
                            Marshal.WriteInt32(lpl, i * sizeof(int), propertyValue.Value.MVl.Lpl[i]);
                        }

                        Marshal.WriteInt32(ptr, offset, lpl.ToInt32());
                        break;
                    }

                case PropertyTypeValue.PtypMultipleString8:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVszA.CValues);
                        offset += sizeof(uint);
                        IntPtr lppszA = Marshal.AllocHGlobal((int)(propertyValue.Value.MVszA.CValues * 4));

                        for (int i = 0; i < propertyValue.Value.MVszA.CValues; i++)
                        {
                            Marshal.WriteInt32(lppszA, 4 * i, Marshal.StringToHGlobalAnsi(propertyValue.Value.MVszA.LppszA[i]).ToInt32());
                        }

                        Marshal.WriteInt32(ptr, offset, lppszA.ToInt32());
                        break;
                    }

                case PropertyTypeValue.PtypMultipleBinary:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVbin.CValues);
                        offset += sizeof(uint);
                        IntPtr mvbin = Marshal.AllocHGlobal((int)(propertyValue.Value.MVbin.CValues * 8));

                        for (int i = 0; i < propertyValue.Value.MVbin.CValues; i++)
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

                case PropertyTypeValue.PtypMultipleString:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVszW.CValues);
                        offset += sizeof(uint);
                        IntPtr lppszW = Marshal.AllocHGlobal((int)(propertyValue.Value.MVszW.CValues * 4));

                        for (int i = 0; i < propertyValue.Value.MVszW.CValues; i++)
                        {
                            Marshal.WriteInt32(lppszW, 4 * i, Marshal.StringToHGlobalUni(propertyValue.Value.MVszW.LppszW[i]).ToInt32());
                        }

                        Marshal.WriteInt32(ptr, offset, lppszW.ToInt32());
                        break;
                    }

                case PropertyTypeValue.PtypMultipleGuid:
                    {
                        Marshal.WriteInt32(ptr, offset, (int)propertyValue.Value.MVguid.CValues);
                        offset += sizeof(uint);
                        IntPtr lpguid = Marshal.AllocHGlobal((int)(propertyValue.Value.MVguid.CValues * 4));

                        for (int i = 0; i < propertyValue.Value.MVguid.CValues; i++)
                        {
                            IntPtr guid = Marshal.AllocHGlobal(16);
                            for (int j = 0; j < 16; j++)
                            {
                                Marshal.WriteByte(guid, j, propertyValue.Value.MVguid.Lpguid[i].Ab[j]);
                            }

                            Marshal.WriteInt32(lpguid, 4 * i, guid.ToInt32());
                        }

                        Marshal.WriteInt32(ptr, offset, lpguid.ToInt32());
                        break;
                    }

                case PropertyTypeValue.PtypNull:
                case PropertyTypeValue.PtypEmbeddedTable:
                    {
                        Marshal.WriteInt32(ptr, offset, propertyValue.Value.Reserved);
                        break;
                    }
            }

            return ptr;
        }

        /// <summary>
        /// Free the memory previously allocated for the specific property values.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        public static void FreePropertyValue_r(IntPtr ptr)
        {
            int offset = 0;

            uint propTag = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(uint) * 2;

            PropertyTypeValue proptype = (PropertyTypeValue)(0x0000FFFF & propTag);

            switch (proptype)
            {
                case PropertyTypeValue.PtypString8:
                case PropertyTypeValue.PtypString:
                case PropertyTypeValue.PtypGuid:
                    {
                        Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(ptr, offset)));
                        break;
                    }

                case PropertyTypeValue.PtypBinary:
                case PropertyTypeValue.PtypMultipleInteger16:
                case PropertyTypeValue.PtypMultipleInteger32:
                    {
                        offset += sizeof(uint);
                        Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(ptr, offset)));
                        break;
                    }

                case PropertyTypeValue.PtypMultipleString8:
                case PropertyTypeValue.PtypMultipleString:
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

                case PropertyTypeValue.PtypMultipleBinary:
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

                case PropertyTypeValue.PtypMultipleGuid:
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

                default:
                    Site.Log.Add(LogEntryKind.Debug, "Property type {0} is returned by the server but is not covered by the current test suite.", proptype);
                    break;
            }

            Marshal.FreeHGlobal(ptr);
        }

        /// <summary>
        /// Allocate memory for the specific property values.
        /// </summary>
        /// <param name="pta_r">PropertyTagArray_r instance.</param>
        /// <returns>A pointer points to the allocated memory.</returns>
        public static IntPtr AllocPropertyTagArray_r(PropertyTagArray_r pta_r)
        {
            int offset = 0;
            int cb = (int)(sizeof(uint) + (pta_r.CValues * sizeof(uint)));

            IntPtr ptr = Marshal.AllocHGlobal(cb);

            Marshal.WriteInt32(ptr, offset, (int)pta_r.CValues);
            offset += sizeof(uint);

            for (int i = 0; i < pta_r.CValues; i++)
            {
                Marshal.WriteInt32(ptr, offset, (int)pta_r.AulPropTag[i]);
                offset += sizeof(uint);
            }

            return ptr;
        }

        /// <summary>
        /// Allocate memory for the specific property values.
        /// </summary>
        /// <param name="propertyValues">Array of PropertyValue_r instances.</param>
        /// <returns>A pointer points to the allocated memory.</returns>
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

                PropertyTypeValue proptype = (PropertyTypeValue)(0x0000FFFF & propertyValues[k].PropTag);
                switch (proptype)
                {
                    case PropertyTypeValue.PtypInteger16:
                        {
                            Marshal.WriteInt16(ptr, (16 * k) + offset, propertyValues[k].Value.I);
                            break;
                        }

                    case PropertyTypeValue.PtypInteger32:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, propertyValues[k].Value.L);
                            break;
                        }

                    case PropertyTypeValue.PtypBoolean:
                        {
                            Marshal.WriteInt16(ptr, (16 * k) + offset, (short)propertyValues[k].Value.B);
                            break;
                        }

                    case PropertyTypeValue.PtypString8:
                        {
                            IntPtr strA = IntPtr.Zero;

                            if (propertyValues[k].Value.LpszA != null)
                            {
                                string str = System.Text.Encoding.Default.GetString(propertyValues[k].Value.LpszA);
                                strA = Marshal.StringToHGlobalAnsi(str);
                            }

                            Marshal.WriteInt32(ptr, (16 * k) + offset, strA.ToInt32());
                            break;
                        }

                    case PropertyTypeValue.PtypBinary:
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

                    case PropertyTypeValue.PtypString:
                        {
                            IntPtr strW = IntPtr.Zero;

                            if (propertyValues[k].Value.LpszW != null)
                            {
                                string str = System.Text.Encoding.Unicode.GetString(propertyValues[k].Value.LpszW);
                                strW = Marshal.StringToHGlobalUni(str);
                            }

                            Marshal.WriteInt32(ptr, (16 * k) + offset, strW.ToInt32());
                            break;
                        }

                    case PropertyTypeValue.PtypGuid:
                        {
                            IntPtr guid = Marshal.AllocHGlobal(16);

                            for (int i = 0; i < 16; i++)
                            {
                                Marshal.WriteByte(guid, i, propertyValues[k].Value.Lpguid[0].Ab[i]);
                            }

                            Marshal.WriteInt32(ptr, (16 * k) + offset, guid.ToInt32());
                            break;
                        }

                    case PropertyTypeValue.PtypTime:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.Ft.LowDateTime);
                            offset += sizeof(uint);
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.Ft.HighDateTime);
                            break;
                        }

                    case PropertyTypeValue.PtypErrorCode:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, propertyValues[k].Value.Err);
                            break;
                        }

                    case PropertyTypeValue.PtypMultipleInteger16:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVi.CValues);
                            offset += sizeof(uint);
                            IntPtr lpi = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVi.CValues * sizeof(short)));

                            for (int i = 0; i < propertyValues[k].Value.MVi.CValues; i++)
                            {
                                Marshal.WriteInt16(lpi, i * sizeof(short), propertyValues[k].Value.MVi.Lpi[i]);
                            }

                            Marshal.WriteInt32(ptr, (16 * k) + offset, lpi.ToInt32());
                            break;
                        }

                    case PropertyTypeValue.PtypMultipleInteger32:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVl.CValues);
                            offset += sizeof(uint);
                            IntPtr lpl = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVl.CValues * sizeof(int)));

                            for (int i = 0; i < propertyValues[k].Value.MVl.CValues; i++)
                            {
                                Marshal.WriteInt32(lpl, i * sizeof(int), propertyValues[k].Value.MVl.Lpl[i]);
                            }

                            Marshal.WriteInt32(ptr, (16 * k) + offset, lpl.ToInt32());
                            break;
                        }

                    case PropertyTypeValue.PtypMultipleString8:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVszA.CValues);
                            offset += sizeof(uint);
                            IntPtr lppszA = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVszA.CValues * 4));

                            for (int i = 0; i < propertyValues[k].Value.MVszA.CValues; i++)
                            {
                                Marshal.WriteInt32(lppszA, 4 * i, Marshal.StringToHGlobalAnsi(propertyValues[k].Value.MVszA.LppszA[i]).ToInt32());
                            }

                            Marshal.WriteInt32(ptr, (16 * k) + offset, lppszA.ToInt32());
                            break;
                        }

                    case PropertyTypeValue.PtypMultipleBinary:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVbin.CValues);
                            offset += sizeof(uint);
                            IntPtr mvbin = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVbin.CValues * 8));

                            for (int i = 0; i < propertyValues[k].Value.MVbin.CValues; i++)
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

                    case PropertyTypeValue.PtypMultipleString:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVszW.CValues);
                            offset += sizeof(uint);
                            IntPtr lppszW = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVszW.CValues * 4));

                            for (int i = 0; i < propertyValues[k].Value.MVszW.CValues; i++)
                            {
                                Marshal.WriteInt32(lppszW, 4 * i, Marshal.StringToHGlobalUni(propertyValues[k].Value.MVszW.LppszW[i]).ToInt32());
                            }

                            Marshal.WriteInt32(ptr, (16 * k) + offset, lppszW.ToInt32());
                            break;
                        }

                    case PropertyTypeValue.PtypMultipleGuid:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, (int)propertyValues[k].Value.MVguid.CValues);
                            offset += sizeof(uint);
                            IntPtr lpguid = Marshal.AllocHGlobal((int)(propertyValues[k].Value.MVguid.CValues * 4));

                            for (int i = 0; i < propertyValues[k].Value.MVguid.CValues; i++)
                            {
                                IntPtr guid = Marshal.AllocHGlobal(16);

                                for (int j = 0; j < 16; j++)
                                {
                                    Marshal.WriteByte(guid, j, propertyValues[k].Value.MVguid.Lpguid[i].Ab[j]);
                                }

                                Marshal.WriteInt32(lpguid, 4 * i, guid.ToInt32());
                            }

                            Marshal.WriteInt32(ptr, (16 * k) + offset, lpguid.ToInt32());
                            break;
                        }

                    case PropertyTypeValue.PtypNull:
                    case PropertyTypeValue.PtypEmbeddedTable:
                        {
                            Marshal.WriteInt32(ptr, (16 * k) + offset, propertyValues[k].Value.Reserved);
                            break;
                        }

                    default:
                        Site.Log.Add(LogEntryKind.Debug, "Type {0} is returned by the server but is not covered by the current test suite.", proptype);
                        break;
                }
            }

            return ptr;
        }

        /// <summary>
        /// Free the memory previously allocated for the specific property values.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <param name="count">The number of PropertyValue_r.</param>
        public static void FreePropertyValue_rs(IntPtr ptr, int count)
        {
            const int PropertyValueLengthInBytes = 16;

            for (int k = 0; k < count; k++)
            {
                int offset = 0;

                uint propTag = (uint)Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset);
                offset += sizeof(uint) * 2;

                PropertyTypeValue proptype = (PropertyTypeValue)(0x0000FFFF & propTag);

                switch (proptype)
                {
                    case PropertyTypeValue.PtypString8:
                    case PropertyTypeValue.PtypString:
                    case PropertyTypeValue.PtypGuid:
                        {
                            Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset)));
                            break;
                        }

                    case PropertyTypeValue.PtypBinary:
                    case PropertyTypeValue.PtypMultipleInteger16:
                    case PropertyTypeValue.PtypMultipleInteger32:
                        {
                            offset += sizeof(uint);
                            Marshal.FreeHGlobal(new IntPtr(Marshal.ReadInt32(ptr, (PropertyValueLengthInBytes * k) + offset)));
                            break;
                        }

                    case PropertyTypeValue.PtypMultipleString8:
                    case PropertyTypeValue.PtypMultipleString:
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

                    case PropertyTypeValue.PtypMultipleBinary:
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

                    case PropertyTypeValue.PtypMultipleGuid:
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

                    default:
                        Site.Log.Add(LogEntryKind.Debug, "Property type {0} is returned by the server but is not covered by the current test suite.", proptype);
                        break;
                }
            }

            Marshal.FreeHGlobal(ptr);
        }

        /// <summary>
        /// Allocate memory for the table row.
        /// </summary>
        /// <param name="propertyRow">Instance of table row structure.</param>
        /// <returns>A pointer points to the allocated memory.</returns>
        public static IntPtr AllocPropertyRow_r(PropertyRow_r propertyRow)
        {
            IntPtr ptr = Marshal.AllocHGlobal(12);
            int offset = 0;
            Marshal.WriteInt32(ptr, offset, (int)propertyRow.Reserved);
            offset += sizeof(uint);
            Marshal.WriteInt32(ptr, offset, (int)propertyRow.CValues);
            offset += sizeof(uint);

            IntPtr pv = AllocPropertyValue_rs(propertyRow.LpProps);
            Marshal.WriteInt32(ptr, offset, pv.ToInt32());

            return ptr;
        }

        /// <summary>
        /// Free the memory previously allocated for the table row.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
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
        /// <returns>A pointer points to the allocated memory.</returns>
        public static IntPtr AllocFlatUID_r(FlatUID_r fuid_r)
        {
            IntPtr ptr = Marshal.AllocHGlobal(Constants.FlatUIDByteSize);

            for (int i = 0; i < Constants.FlatUIDByteSize; i++)
            {
                Marshal.WriteByte(ptr, i, fuid_r.Ab[i]);
            }

            return ptr;
        }

        /// <summary>
        /// Allocate memory for the specific PropertyName_r values.
        /// </summary>
        /// <param name="propName">The PropertyName_r instance.</param>
        /// <returns>A pointer points to the allocated memory.</returns>
        public static IntPtr AllocPropertyName_r(PropertyName_r propName)
        {
            IntPtr ptr = Marshal.AllocHGlobal(12);
            IntPtr uid = Marshal.AllocHGlobal(Constants.FlatUIDByteSize);
            int offset = 0;

            // Allocate memory for FlatUID_r structure of propName.
            for (int i = 0; i < Constants.FlatUIDByteSize; i++)
            {
                Marshal.WriteByte(uid, i, propName.Guid[0].Ab[i]);
            }

            // Write address of FlatUid_r structure.
            Marshal.WriteInt32(ptr, offset, uid.ToInt32());
            offset += sizeof(int);

            // Write ulReserved value.
            Marshal.WriteInt32(ptr, offset, (int)propName.Reserved);
            offset += sizeof(int);

            // Write lID value.
            Marshal.WriteInt32(ptr, offset, propName.ID);
            offset += sizeof(int);

            return ptr;
        }

        /// <summary>
        /// Free the memory allocated for the specific PropertyName_r instance.
        /// </summary>
        /// <param name="ptr">The pointer points to the allocated memory.</param>
        public static void FreePropertyName_r(IntPtr ptr)
        {
            // Free FlatUID_r structure.
            IntPtr uid = new IntPtr(Marshal.ReadInt32(ptr));
            Marshal.FreeHGlobal(uid);

            Marshal.FreeHGlobal(ptr);
        }
        #endregion

        #region Comparator Methods
        /// <summary>
        /// Compare whether the two PropertyRowSet_r structures are equal.
        /// </summary>
        /// <param name="propertyRowSet1">The first element to be compared.</param>
        /// <param name="propertyRowSet2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoPropertyRowSetEqual(PropertyRowSet_r? propertyRowSet1, PropertyRowSet_r? propertyRowSet2)
        {
            if (propertyRowSet1 == null && propertyRowSet2 == null)
            {
                site.Log.Add(LogEntryKind.Debug, "Both of the two propertyRowSet are null.");
                return true;
            }
            else if (propertyRowSet1 == null)
            {
                site.Log.Add(LogEntryKind.Debug, "One of the two propertyRowSet is null and the RowCount field of the other propertyRowSet is {0}", propertyRowSet2.Value.CRows);

                return false;
            }
            else if (propertyRowSet2 == null)
            {
                site.Log.Add(LogEntryKind.Debug, "One of the two propertyRowSet is null and the RowCount field of the other propertyRowSet is {0}", propertyRowSet1.Value.CRows);

                return false;
            }

            if (propertyRowSet1.Value.CRows != propertyRowSet2.Value.CRows)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length of two propertyRowSet are not equal. The length of propertyRowSet1 is {0}, the length of propertyRowSet2 is {1}",
                    propertyRowSet1.Value.CRows,
                    propertyRowSet2.Value.CRows);

                return false;
            }

            for (int i = 0; i < propertyRowSet1.Value.CRows; i++)
            {
                if (!AdapterHelper.AreTwoPropertyRowEqual(propertyRowSet1.Value.ARow[i], propertyRowSet2.Value.ARow[i]))
                {
                    site.Log.Add(
                        LogEntryKind.Debug,
                        "The ARow {0} of the two propertyRowSet are not equal. The ARow value of propertyRowSet1 is {1}, the ARow value of propertyRowSet2 is {2}",
                        i,
                        propertyRowSet1.Value.ARow[i],
                        propertyRowSet2.Value.ARow[i]);

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether the two PropertyRow_r structures are equal.
        /// </summary>
        /// <param name="propertyRow1">The first element to be compared.</param>
        /// <param name="propertyRow2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoPropertyRowEqual(PropertyRow_r? propertyRow1, PropertyRow_r? propertyRow2)
        {
            if (propertyRow1.Value.CValues != propertyRow2.Value.CValues || propertyRow1.Value.Reserved != propertyRow2.Value.Reserved)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length or reserved of the two property row are not equal. The length of propertyRowSet1 is {0}, the length of propertyRowSet2 is {1}. The reserved of propertyRowSet1 is {2}, the reserved of propertyRowSet2 is {3}",
                    propertyRow1.Value.CValues,
                    propertyRow2.Value.CValues,
                    propertyRow1.Value.Reserved,
                    propertyRow2.Value.Reserved);

                return false;
            }

            for (int i = 0; i < propertyRow1.Value.CValues; i++)
            {
                if (!AdapterHelper.AreTwoPropertyValueEqual(propertyRow1.Value.LpProps[i], propertyRow2.Value.LpProps[i]))
                {
                    site.Log.Add(
                        LogEntryKind.Debug,
                        "The property row of index {0} of the two propertyRowSet are not equal. The property row of propertyRowSet1 is {1}, the property row of propertyRowSet2 is {2}",
                        i,
                        propertyRow1.Value.LpProps[i],
                        propertyRow2.Value.LpProps[i]);

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether the two PropertyTagArray_r structures are equal.
        /// </summary>
        /// <param name="propertyTagArray1">The first element to be compared.</param>
        /// <param name="propertyTagArray2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoPropertyTagArrayEqual(PropertyTagArray_r? propertyTagArray1, PropertyTagArray_r? propertyTagArray2)
        {
            if (propertyTagArray1.Value.CValues != propertyTagArray2.Value.CValues)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length the two property tag array are not equal. The length of propertyTagArray1 is {0}, the length of propertyTagArray2 is {1}.",
                    propertyTagArray1.Value.CValues,
                    propertyTagArray2.Value.CValues);

                return false;
            }

            for (int i = 0; i < propertyTagArray1.Value.CValues; i++)
            {
                if (propertyTagArray1.Value.AulPropTag[i] != propertyTagArray2.Value.AulPropTag[i])
                {
                    site.Log.Add(
                        LogEntryKind.Debug,
                        "The property tag of index {0} of the two property tag array are not equal. The property tag of propertyTagArray1 is {1}, the property tag of propertyTagArray2 is {2}",
                        i,
                        propertyTagArray1.Value.AulPropTag[i],
                        propertyTagArray2.Value.AulPropTag[i]);

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether the two PropertyValue_r structures are equal.
        /// </summary>
        /// <param name="propertyValue1">The first element to be compared.</param>
        /// <param name="propertyValue2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoPropertyValueEqual(PropertyValue_r propertyValue1, PropertyValue_r propertyValue2)
        {
            if (propertyValue1.PropTag != propertyValue2.PropTag || propertyValue1.Reserved != propertyValue2.Reserved)
            {
                return false;
            }

            switch (propertyValue1.PropTag & 0x0000ffff)
            {
                case 0x00000002:
                    if (propertyValue1.Value.I == propertyValue2.Value.I)
                    {
                        return true;
                    }
                    else
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The two property value are not equal. The property value of propertyValue1 is {0}, the property value of propertyValue2 is {1}",
                            propertyValue1.Value.I,
                            propertyValue2.Value.I);

                        return false;
                    }

                case 0x00000003:
                    if (propertyValue1.Value.L == propertyValue2.Value.L)
                    {
                        return true;
                    }
                    else
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The two property value are not equal. The property value of propertyValue1 is {0}, the property value of propertyValue2 is {1}",
                            propertyValue1.Value.L,
                            propertyValue2.Value.L);

                        return false;
                    }

                case 0x0000000b:
                    if (propertyValue1.Value.B == propertyValue2.Value.B)
                    {
                        return true;
                    }
                    else
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The two property value are not equal. The property value of propertyValue1 is {0}, the property value of propertyValue2 is {1}",
                            propertyValue1.Value.B,
                            propertyValue2.Value.B);

                        return false;
                    }

                case 0x0000001e:
                    if (AdapterHelper.AreTwoByteArrayEqual(propertyValue1.Value.LpszA, propertyValue2.Value.LpszA))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                case 0x00000102:
                    if (AdapterHelper.AreTwoBinary_rEqual(propertyValue1.Value.Bin, propertyValue2.Value.Bin))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                case 0x0000001f:
                    if (AdapterHelper.AreTwoByteArrayEqual(propertyValue1.Value.LpszW, propertyValue2.Value.LpszW))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                case 0x00000048:
                    if (AdapterHelper.AreTwoFlatUID_rArrayEqual(propertyValue1.Value.Lpguid, propertyValue2.Value.Lpguid))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                case 0x00000040:
                    if (propertyValue1.Value.Ft.HighDateTime == propertyValue2.Value.Ft.HighDateTime
                        && propertyValue1.Value.Ft.LowDateTime == propertyValue2.Value.Ft.LowDateTime)
                    {
                        return true;
                    }
                    else
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The two HighDateTime are not equal or the two LowDateTime are not equal. The HighDateTime of propertyValue1 is {0}, the HighDateTime of propertyValue2 is {1}. The LowDateTime of propertyValue1 is {2}, the LowDateTime of propertyValue2 is {3}.",
                            propertyValue1.Value.Ft.HighDateTime,
                            propertyValue2.Value.Ft.HighDateTime,
                            propertyValue1.Value.Ft.LowDateTime,
                            propertyValue2.Value.Ft.LowDateTime);

                        return false;
                    }

                case 0x0000000a:
                    if (propertyValue1.Value.Err == propertyValue2.Value.Err)
                    {
                        return true;
                    }
                    else
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The two property value are not equal. The property value of propertyValue1 is {0}, the property value of propertyValue2 is {1}",
                            propertyValue1.Value.Err,
                            propertyValue2.Value.Err);

                        return false;
                    }

                case 0x00001002:
                    if (AdapterHelper.AreTwoShortArray_rEqual(propertyValue1.Value.MVi, propertyValue2.Value.MVi))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                case 0x00001003:
                    if (AdapterHelper.AreTwoLongArray_rEqual(propertyValue1.Value.MVl, propertyValue2.Value.MVl))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                case 0x0000101e:
                    if (AdapterHelper.AreTwoStringArray_rEqual(propertyValue1.Value.MVszA, propertyValue2.Value.MVszA))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                case 0x00001102:
                    if (AdapterHelper.AreTwoBinaryArray_rEqual(propertyValue1.Value.MVbin, propertyValue2.Value.MVbin))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                case 0x00001048:
                    if (AdapterHelper.AreTwoFlatUIDArray_rEqual(propertyValue1.Value.MVguid, propertyValue2.Value.MVguid))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                case 0x0000101f:
                    if (AdapterHelper.AreTwoWStringArray_rEqual(propertyValue1.Value.MVszW, propertyValue2.Value.MVszW))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                case 0x00001040:
                    if (propertyValue1.Value.MVft.Equals(propertyValue2.Value.MVft))
                    {
                        return true;
                    }
                    else
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The two property value are not equal. The property value of propertyValue1 is {0}, the property value of propertyValue2 is {1}",
                            propertyValue1.Value.MVft,
                            propertyValue2.Value.MVft);

                        return false;
                    }

                case 0x00000001:
                case 0x0000000d:
                    if (propertyValue1.Value.Reserved == propertyValue2.Value.Reserved)
                    {
                        return true;
                    }
                    else
                    {
                        site.Log.Add(
                            LogEntryKind.Debug,
                            "The two property value are not equal. The property value of propertyValue1 is {0}, the property value of propertyValue2 is {1}",
                            propertyValue1.Value.Reserved,
                            propertyValue2.Value.Reserved);

                        return false;
                    }

                default:
                    return false;
            }
        }

        /// <summary>
        /// Compare whether the two byte array structures are equal.
        /// </summary>
        /// <param name="byteArray1">The first element to be compared.</param>
        /// <param name="byteArray2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoByteArrayEqual(byte[] byteArray1, byte[] byteArray2)
        {
            if (byteArray1 == null && byteArray2 == null)
            {
                return true;
            }

            if (byteArray1.Length != byteArray2.Length)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length of two byte array are not equal. The length of byteArray1 is {0}, the length of byteArray2 is {1}",
                    byteArray1.Length,
                    byteArray2.Length);

                return false;
            }

            for (int i = 0; i < byteArray1.Length; i++)
            {
                if (byteArray1[i] != byteArray2[i])
                {
                    site.Log.Add(
                        LogEntryKind.Debug,
                        "The byte value of index {0} of two byte array are not equal. The byte value of byteArray1 is {1}, the byte value of byteArray2 is {2}",
                        i,
                        byteArray1[i],
                        byteArray2[i]);

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether the two FlatUID_r array structures are equal.
        /// </summary>
        /// <param name="flatUidArray1">The first element to be compared.</param>
        /// <param name="flatUidArray2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoFlatUID_rArrayEqual(FlatUID_r[] flatUidArray1, FlatUID_r[] flatUidArray2)
        {
            if (flatUidArray1.Length != flatUidArray2.Length)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length of two flat UID array are not equal. The length of flatUidArray1 is {0}, the length of flatUidArray2 is {1}",
                    flatUidArray1.Length,
                    flatUidArray2.Length);

                return false;
            }

            for (int i = 0; i < flatUidArray1.Length; i++)
            {
                if (!AdapterHelper.AreTwoByteArrayEqual(flatUidArray1[i].Ab, flatUidArray2[i].Ab))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether the two ShortArray_r structures are equal.
        /// </summary>
        /// <param name="shortArray1">The first element to be compared.</param>
        /// <param name="shortArray2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoShortArray_rEqual(ShortArray_r shortArray1, ShortArray_r shortArray2)
        {
            if (shortArray1.CValues != shortArray2.CValues)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length of two short array are not equal. The length of shortArray1 is {0}, the length of shortArray2 is {1}",
                    shortArray1.CValues,
                    shortArray2.CValues);

                return false;
            }

            for (int i = 0; i < shortArray1.CValues; i++)
            {
                if (shortArray1.Lpi[i] != shortArray2.Lpi[i])
                {
                    site.Log.Add(
                        LogEntryKind.Debug,
                        "The short value of index {0} of two short array are not equal. The short value of shortArray1 is {1}, the short value of shortArray2 is {2}",
                        i,
                        shortArray1.Lpi[i],
                        shortArray2.Lpi[i]);

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether the two LongArray_r structures are equal.
        /// </summary>
        /// <param name="longArray1">The first element to be compared.</param>
        /// <param name="longArray2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoLongArray_rEqual(LongArray_r longArray1, LongArray_r longArray2)
        {
            if (longArray1.CValues != longArray2.CValues)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length of two long array are not equal. The length of longArray1 is {0}, the length of longArray2 is {1}",
                    longArray1.CValues,
                    longArray2.CValues);

                return false;
            }

            for (int i = 0; i < longArray1.CValues; i++)
            {
                if (longArray1.Lpl[i] != longArray2.Lpl[i])
                {
                    site.Log.Add(
                        LogEntryKind.Debug,
                        "The long value of index {0} of two long array are not equal. The long value of longArray1 is {1}, the long value of longArray2 is {2}",
                        i,
                        longArray1.Lpl[i],
                        longArray2.Lpl[i]);

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether the two StringArray_r structures are equal.
        /// </summary>
        /// <param name="stringArray1">The first element to be compared.</param>
        /// <param name="stringArray2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoStringArray_rEqual(StringArray_r stringArray1, StringArray_r stringArray2)
        {
            if (stringArray1.CValues != stringArray2.CValues)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length of two string array are not equal. The length of stringArray1 is {0}, the length of stringArray2 is {1}",
                    stringArray1.CValues,
                    stringArray2.CValues);

                return false;
            }

            for (int i = 0; i < stringArray2.CValues; i++)
            {
                if (stringArray1.LppszA[i] != stringArray2.LppszA[i])
                {
                    site.Log.Add(
                        LogEntryKind.Debug,
                        "The string value of index {0} of two string array are not equal. The string value of stringArray1 is {1}, the string value of stringArray2 is {2}",
                        i,
                        stringArray1.LppszA[i],
                        stringArray2.LppszA[i]);

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether the two BinaryArray_r structures are equal.
        /// </summary>
        /// <param name="binaryArray1">The first element to be compared.</param>
        /// <param name="binaryArray2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoBinaryArray_rEqual(BinaryArray_r binaryArray1, BinaryArray_r binaryArray2)
        {
            if (binaryArray1.CValues != binaryArray2.CValues)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length of two binary array are not equal. The length of binaryArray1 is {0}, the length of binaryArray2 is {1}",
                    binaryArray1.CValues,
                    binaryArray2.CValues);

                return false;
            }

            for (int i = 0; i < binaryArray1.CValues; i++)
            {
                if (AdapterHelper.AreTwoBinary_rEqual(binaryArray1.Lpbin[i], binaryArray2.Lpbin[i]))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether the two Binary_r structures are equal.
        /// </summary>
        /// <param name="binary1">The first element to be compared.</param>
        /// <param name="binary2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoBinary_rEqual(Binary_r binary1, Binary_r binary2)
        {
            if (binary1.Cb != binary2.Cb)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length of two binary are not equal. The length of binary1 is {0}, the length of binary2 is {1}",
                    binary1.Cb,
                    binary2.Cb);

                return false;
            }

            for (int i = 0; i < binary1.Cb; i++)
            {
                if (binary1.Lpb[i] != binary2.Lpb[i])
                {
                    site.Log.Add(
                        LogEntryKind.Debug,
                        "The uninterpreted bytes value of index {0} of two binary are not equal. The uninterpreted bytes value of binary1 is {1}, the uninterpreted bytes value of binary2 is {2}",
                        i,
                        binary1.Lpb[i],
                        binary2.Lpb[i]);

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether the two FlatUIDArray_r structures are equal.
        /// </summary>
        /// <param name="flatUIDArray1">The first element to be compared.</param>
        /// <param name="flatUIDArray2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoFlatUIDArray_rEqual(FlatUIDArray_r flatUIDArray1, FlatUIDArray_r flatUIDArray2)
        {
            if (flatUIDArray1.CValues != flatUIDArray2.CValues
                || !AdapterHelper.AreTwoFlatUID_rArrayEqual(flatUIDArray1.Lpguid, flatUIDArray2.Lpguid))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Compare whether the two WStringArray_r structures are equal.
        /// </summary>
        /// <param name="wstringArray1">The first element to be compared.</param>
        /// <param name="wstringArray2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoWStringArray_rEqual(WStringArray_r wstringArray1, WStringArray_r wstringArray2)
        {
            if (wstringArray1.CValues != wstringArray2.CValues)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The length of two wstring array are not equal. The length of wstringArray1 is {0}, the length of wstringArray2 is {1}",
                    wstringArray1.CValues,
                    wstringArray2.CValues);

                return false;
            }

            for (int i = 0; i < wstringArray1.CValues; i++)
            {
                if (wstringArray1.LppszW[i] != wstringArray2.LppszW[i])
                {
                    site.Log.Add(
                        LogEntryKind.Debug,
                        "The wstring value of index {0} of two wstring array are not equal. The wstring value of wstringArray1 is {1}, the wstring value of wstringArray2 is {2}",
                        i,
                        wstringArray1.LppszW[i],
                        wstringArray2.LppszW[i]);

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Compare whether the two Permanent Entry ID structures are equal.
        /// </summary>
        /// <param name="permanentEntryID1">The first element to be compared.</param>
        /// <param name="permanentEntryID2">The second element to be compared.</param>
        /// <returns>If they are equal, return true, else false.</returns>
        public static bool AreTwoPermanentEntryIDEqual(PermanentEntryID? permanentEntryID1, PermanentEntryID? permanentEntryID2)
        {
            site.Log.Add(LogEntryKind.Debug, "Compare whether the two Permanent Entry ID structures are equal");

            if (permanentEntryID1.Value.DisplayTypeString != permanentEntryID2.Value.DisplayTypeString)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The DisplayTypeString of the two Permanent Entry ID are not equal, the first is:\r\n{0}, and the second is:\r\n{1}",
                    permanentEntryID1.Value.DisplayTypeString,
                    permanentEntryID2.Value.DisplayTypeString);
                return false;
            }

            if (!permanentEntryID1.Value.DistinguishedName.Equals(permanentEntryID2.Value.DistinguishedName, StringComparison.CurrentCultureIgnoreCase))
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The DistinguishedName of the two Permanent Entry ID are not equal, the first is:\r\n{0}, and the second is:\r\n{1}",
                    permanentEntryID1.Value.DistinguishedName,
                    permanentEntryID2.Value.DistinguishedName);
                return false;
            }

            if (permanentEntryID1.Value.IDType != permanentEntryID2.Value.IDType)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The IDType of the two Permanent Entry ID are not equal, the first is:\r\n{0}, and the second is:\r\n{1}",
                    permanentEntryID1.Value.IDType,
                    permanentEntryID2.Value.IDType);
                return false;
            }

            if (!AdapterHelper.AreTwoByteArrayEqual(permanentEntryID1.Value.ProviderUID.Ab, permanentEntryID2.Value.ProviderUID.Ab))
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The ProviderUID of the two Permanent Entry ID are not equal");
                return false;
            }

            if (permanentEntryID1.Value.R1 != permanentEntryID2.Value.R1)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The R1 of the two Permanent Entry ID are not equal, the first is:\r\n{0}, and the second is:\r\n{1}",
                    permanentEntryID1.Value.R1,
                    permanentEntryID2.Value.R1);
                return false;
            }

            if (permanentEntryID1.Value.R2 != permanentEntryID2.Value.R2)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The R2 of the two Permanent Entry ID are not equal, the first is:\r\n{0}, and the second is:\r\n{1}",
                    permanentEntryID1.Value.R2,
                    permanentEntryID2.Value.R2);
                return false;
            }

            if (permanentEntryID1.Value.R3 != permanentEntryID2.Value.R3)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The R3 of the two Permanent Entry ID are not equal, the first is:\r\n{0}, and the second is:\r\n{1}",
                    permanentEntryID1.Value.R3,
                    permanentEntryID2.Value.R3);
                return false;
            }

            if (permanentEntryID1.Value.R4 != permanentEntryID2.Value.R4)
            {
                site.Log.Add(
                    LogEntryKind.Debug,
                    "The R4 of the two Permanent Entry ID are not equal, the first is:\r\n{0}, and the second is:\r\n{1}",
                    permanentEntryID1.Value.R4,
                    permanentEntryID2.Value.R4);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Find property value with specified property tag in PropertyRowSet_r.
        /// </summary>
        /// <param name="rowSet">The PropertyRowSet_r to find specified tag.</param>
        /// <param name="specifiedPropTag">The specified tag.</param>
        /// <param name="property">The found property value with specified tag.</param>
        /// <returns>If the specified tag is found, return true, else false.</returns>
        public static bool FindFirstSpecifiedPropTagValueInRowSet(PropertyRowSet_r rowSet, uint specifiedPropTag, out PropertyValue_r? property)
        {
            property = null;

            if (rowSet.CRows == 0)
            {
                return false;
            }

            foreach (PropertyRow_r row in rowSet.ARow)
            {
                foreach (PropertyValue_r pro in row.LpProps)
                {
                    if (pro.PropTag == specifiedPropTag)
                    {
                        property = pro;
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// To determine whether DT_CONTAINERType is correct.
        /// </summary>
        /// <param name="rows">A nullable PropertyRowSet_r instance.</param>
        /// <returns>If DT_CONTAINERType is correct, return true, else false.</returns>
        public static bool IsDT_CONTAINERTypeCorrect(PropertyRowSet_r? rows)
        {
            PermanentEntryID entryID = new PermanentEntryID();

            foreach (PropertyRow_r row in rows.Value.ARow)
            {
                foreach (PropertyValue_r val in row.LpProps)
                {
                    if (val.PropTag == (uint)AulProp.PidTagEntryId)
                    {
                        entryID = AdapterHelper.ParsePermanentEntryIDFromBytes(val.Value.Bin.Lpb);
                    }
                }

                if (entryID.DisplayTypeString == DisplayTypeValue.DT_CONTAINER)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// To determine whether the string value in the PropertyRowSet_r structure is of type PtypString.
        /// </summary>
        /// <param name="rowSet">The PropertyRowSet_r structure to be checked.</param>
        /// <returns>If the string value in the PropertyRowSet_r structure is PtypString type, return true, else false.</returns>
        public static bool IsPtypString(PropertyRowSet_r rowSet)
        {
            if (rowSet.CRows == 0)
            {
                return false;
            }

            foreach (PropertyRow_r row in rowSet.ARow)
            {
                foreach (PropertyValue_r pro in row.LpProps)
                {
                    if ((pro.PropTag & 0x0000ffff) == (uint)PropertyTypeValue.PtypString8)
                    {
                        return false;
                    }

                    if ((pro.PropTag & 0x0000ffff) == (uint)PropertyTypeValue.PtypMultipleString8)
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// To determine whether the string value in the PropertyRowSet_r structure is of type PtypString8.
        /// </summary>
        /// <param name="rowSet">The PropertyRowSet_r structure to be checked.</param>
        /// <returns>If the string value in the PropertyRowSet_r structure is PtypString8 type, return true, else false.</returns>
        public static bool IsPtypString8(PropertyRowSet_r rowSet)
        {
            if (rowSet.CRows == 0)
            {
                return false;
            }

            foreach (PropertyRow_r row in rowSet.ARow)
            {
                foreach (PropertyValue_r pro in row.LpProps)
                {
                    if ((pro.PropTag & 0x0000ffff) == (uint)PropertyTypeValue.PtypString)
                    {
                        return false;
                    }

                    if ((pro.PropTag & 0x0000ffff) == (uint)PropertyTypeValue.PtypMultipleString)
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// To determine whether the property rows are sorted by name.
        /// </summary>
        /// <param name="rows">A nullable PropertyRowSet_r instance.</param>
        /// <param name="indexOfPidTagDisplayName">The position of PidTagDisplayName in input parameters propTags.</param>
        /// <returns>If the property rows are sorted by name, return true, else false.</returns>
        public static bool IsSortByName(PropertyRowSet_r? rows, int indexOfPidTagDisplayName)
        {
            // store name of each row
            string[] names = new string[rows.Value.CRows];
            Site.Log.Add(LogEntryKind.Debug, "The row sorted order returned by display name is:");
            for (int i = 0; i < rows.Value.CRows; i++)
            {
                names[i] = System.Text.Encoding.Default.GetString(rows.Value.ARow[i].LpProps[indexOfPidTagDisplayName].Value.LpszA);
                Site.Log.Add(LogEntryKind.Debug, "{0}", names[i]);
            }

            for (int i = 0; i <= names.Length - 2; i++)
            {
                if (string.Compare(names[i], names[i + 1]) > 0)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Whether the propRowSet follows the propTags.
        /// </summary>
        /// <param name="propTags">The PropertyRowSet_r must follow this.</param>
        /// <param name="propRowSet">It must follow the PropertyTagArray_r.</param>
        /// <returns>If propRowSet follows propTags, return true, else false.</returns>
        public static bool IsRowSetSubjectToPropTags(PropertyTagArray_r? propTags, PropertyRowSet_r? propRowSet)
        {
            bool result = true;
            foreach (PropertyRow_r propRow in propRowSet.Value.ARow)
            {
                if (!AdapterHelper.IsRowSubjectToPropTags(propTags, propRow))
                {
                    result = false;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Whether the propRow follows the propTags.
        /// </summary>
        /// <param name="propTags">The PropertyRow_r must follow this.</param>
        /// <param name="propRow">It must follow the PropertyTagArray_r.</param>
        /// <returns>A Boolean value indicates whether the propRow follows the propTags.
        /// The value "true" indicates the propRow follows the propTags.</returns>
        public static bool IsRowSubjectToPropTags(PropertyTagArray_r? propTags, PropertyRow_r? propRow)
        {
            bool result = true;
            for (int i = 0; i < propTags.Value.CValues; i++)
            {
                if (propTags.Value.AulPropTag[i] != propRow.Value.LpProps[i].PropTag)
                {
                    site.Log.Add(
                        LogEntryKind.Debug,
                        "The property row {0} does not follow the corresponding property tag. The value of property row is {1}, the value of property tag is {2}",
                        i,
                        propRow.Value.LpProps[i].PropTag,
                        propTags.Value.AulPropTag[i]);

                    result = false;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Check whether a specific display type is returned as part of the EntryID field in the property row set.
        /// </summary>
        /// <param name="rows">The property rows returned by the server.</param>
        /// <param name="specifieDisplayType">A specific display type.</param>
        /// <returns>If the specific display type exists in the rows, return true, else false.</returns>
        public static bool CheckIfSpecificDisplayTypeExists(PropertyRow_r[] rows, DisplayTypeValue specifieDisplayType)
        {
            DisplayTypeValue displayType = (DisplayTypeValue)0xff;
            foreach (PropertyRow_r propertyRow in rows)
            {
                for (int i = 0; i <= propertyRow.CValues - 1; i++)
                {
                    if (propertyRow.LpProps[i].PropTag == (uint)AulProp.PidTagEntryId)
                    {
                        // The PidTagEntryId property is in the form of a PermanentEntryID structure. According to the definition of PermanentEntryID structure, the display type is defined from the 25th byte. 
                        displayType = (DisplayTypeValue)BitConverter.ToInt32(propertyRow.LpProps[i].Value.Bin.Lpb, 24);

                        if (displayType == specifieDisplayType)
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }
        #endregion

        #region Converter Methods
        /// <summary>
        /// Convert the PtypSting type to PtypString8 type.
        /// </summary>
        /// <param name="originalStringType">The original property tag with the PtypString type.</param>
        /// <returns>The converted property tag with the PtypString8 type.</returns>
        public static uint ConvertStringToString8(uint originalStringType)
        {
            uint result = originalStringType;
            if ((originalStringType & 0x0000FFFF) == 0x0000001F)
            {
                result = result & 0xFFFF001E;
            }
            else
            {
                throw new ArgumentException("The property type should be PtypString.");
            }

            return result;
        }

        /// <summary>
        /// Convert the PtypSting8 type to PtypString type.
        /// </summary>
        /// <param name="originalString8Type">The original property tag with the PtypString8 type.</param>
        /// <returns>The converted property tag with the PtypString type.</returns>
        public static uint ConvertString8ToString(uint originalString8Type)
        {
            uint result = originalString8Type;
            if ((originalString8Type & 0x0000FFFF) == 0x0000001E)
            {
                result = result | 0x0000001F;
            }
            else
            {
                throw new ArgumentException("The property type should be PtypString8.");
            }

            return result;
        }

        /// <summary>
        /// Convert a Restriction_r structure to the byte array of a Restriction structure.
        /// </summary>
        /// <param name="res_r">The Restriction_r structure instance.</param>
        /// <returns>The byte array of the Restriction structure converted.</returns>
        public static byte[] ConvertRestriction_rToRestriction(Restriction_r res_r)
        {
            List<byte> filter = new List<byte>();
            switch (res_r.Rt)
            {
                // AndRestriction_r
                case 0x00000000:
                    AndRestriction andRestriction = new AndRestriction();
                    andRestriction.RestrictCount = res_r.Res.ResAnd.CRes;
                    andRestriction.Restricts = new byte[andRestriction.RestrictCount][];
                    for (int i = 0; i < andRestriction.RestrictCount; i++)
                    {
                        byte[] restriction = ConvertRestriction_rToRestriction(res_r.Res.ResAnd.LpRes[i]);
                        andRestriction.Restricts[i] = restriction;
                    }

                    filter.AddRange(andRestriction.Serialize());
                    break;

                // OrRestriction_r
                case 0x00000001:
                    OrRestriction restrictionOfOr = new OrRestriction();
                    restrictionOfOr.RestrictCount = res_r.Res.ResOr.CRes;
                    restrictionOfOr.Restricts = new byte[restrictionOfOr.RestrictCount][];
                    for (int i = 0; i < restrictionOfOr.RestrictCount; i++)
                    {
                        byte[] restriction = ConvertRestriction_rToRestriction(res_r.Res.ResOr.LpRes[i]);
                        restrictionOfOr.Restricts[i] = restriction;
                    }

                    filter.AddRange(restrictionOfOr.Serialize());
                    break;

                // NotRestriction_r
                case 0x00000002:
                    NotRestriction notRestriction = new NotRestriction();
                    byte[] restrictionNot = ConvertRestriction_rToRestriction(res_r.Res.ResNot.LpRes[0]);
                    notRestriction.Restrict = restrictionNot;
                    filter.AddRange(notRestriction.Serialize());
                    break;

                // ContentRestriction_r
                case 0x00000003:
                    PropertyTag propertyTagContent = new PropertyTag();
                    propertyTagContent.PropertyId = (ushort)((res_r.Res.ResContent.PropTag & 0xFFFF0000) >> 16);
                    propertyTagContent.PropertyType = (ushort)(res_r.Res.ResContent.PropTag & 0x0000FFFF);
                    TaggedPropertyValue taggedContentPropertyValue = new TaggedPropertyValue();
                    taggedContentPropertyValue.PropertyTag = propertyTagContent;

                    // Get the actual property value by removing the PropTag and Reserved fields (the first 8 bytes in the PropertyValue_r structure).
                    taggedContentPropertyValue.Value = new byte[res_r.Res.ResContent.Prop[0].Serialize().Length - 8];
                    Array.Copy(res_r.Res.ResContent.Prop[0].Serialize(), 8, taggedContentPropertyValue.Value, 0, res_r.Res.ResContent.Prop[0].Serialize().Length - 8);
                    ContentRestriction contentRestriction = new ContentRestriction();
                    contentRestriction.PropTag = propertyTagContent;
                    contentRestriction.FuzzyLevel = res_r.Res.ResContent.FuzzyLevel;
                    contentRestriction.TaggedValue = taggedContentPropertyValue;
                    filter.AddRange(contentRestriction.Serialize());
                    break;

                // PropertyRestriction_r
                case 0x00000004:
                    PropertyTag propertyTag = new PropertyTag();
                    propertyTag.PropertyId = (ushort)((res_r.Res.ResProperty.PropTag & 0xFFFF0000) >> 16);
                    propertyTag.PropertyType = (ushort)(res_r.Res.ResProperty.PropTag & 0x0000FFFF);
                    TaggedPropertyValue taggedPropertyValue = new TaggedPropertyValue();
                    taggedPropertyValue.PropertyTag = propertyTag;

                    // Get the actual property value by removing the PropTag and Reserved fields (the first 8 bytes in the PropertyValue_r structure).
                    taggedPropertyValue.Value = new byte[res_r.Res.ResProperty.Prop[0].Serialize().Length - 8];
                    Array.Copy(res_r.Res.ResProperty.Prop[0].Serialize(), 8, taggedPropertyValue.Value, 0, res_r.Res.ResProperty.Prop[0].Serialize().Length - 8);
                    PropertyRestriction propertyRestriction = new PropertyRestriction();
                    propertyRestriction.PropTag = propertyTag;
                    propertyRestriction.RelOp = (byte)res_r.Res.ResProperty.Relop;
                    propertyRestriction.TaggedValue = taggedPropertyValue;
                    filter.AddRange(propertyRestriction.Serialize());
                    break;

                // ExistRestriction_r
                case 0x00000008:
                    ushort propertyID = (ushort)((res_r.Res.ResExist.PropTag & 0xFFFF0000) >> 16);
                    ushort propertyType = (ushort)(res_r.Res.ResExist.PropTag & 0x0000FFFF);
                    ExistRestriction existRestriction = new ExistRestriction()
                    {
                        PropTag = new PropertyTag(propertyID, propertyType)
                    };
                    filter.AddRange(existRestriction.Serialize());
                    break;

                default:
                    throw new ArgumentException("Unknown type of RestrictionUnion_r.");
            }

            return filter.ToArray();
        }
        #endregion

        #region Parser Methods
        /// <summary>
        /// Parse EphemeralEntryID structure from byte array.
        /// </summary>
        /// <param name="bytes">The byte array to be parsed.</param>
        /// <returns>An EphemeralEntryID structure instance.</returns>
        public static EphemeralEntryID ParseEphemeralEntryIDFromBytes(byte[] bytes)
        {
            int index = 0;

            EphemeralEntryID entryID = new EphemeralEntryID
            {
                IDType = bytes[index++],
                R1 = bytes[index++],
                R2 = bytes[index++],
                R3 = bytes[index++],
                ProviderUID = new FlatUID_r
                {
                    Ab = new byte[Constants.FlatUIDByteSize]
                }
            };
            for (int i = 0; i < Constants.FlatUIDByteSize; i++)
            {
                entryID.ProviderUID.Ab[i] = bytes[index++];
            }

            // R4: 4 bytes
            entryID.R4 = (uint)BitConverter.ToInt32(bytes, index);
            index += 4;

            // DisplayType: 4 bytes
            entryID.DisplayType = (DisplayTypeValue)BitConverter.ToInt32(bytes, index);
            index += 4;

            // Mid: 4 bytes
            entryID.Mid = (uint)BitConverter.ToInt32(bytes, index);
            index += 4;

            return entryID;
        }

        /// <summary>
        /// Parse PermanentEntryID structure from byte array.
        /// </summary>
        /// <param name="bytes">The byte array to be parsed.</param>
        /// <returns>An PermanentEntryID structure instance.</returns>
        public static PermanentEntryID ParsePermanentEntryIDFromBytes(byte[] bytes)
        {
            int index = 0;

            PermanentEntryID entryID = new PermanentEntryID
            {
                IDType = bytes[index++],
                R1 = bytes[index++],
                R2 = bytes[index++],
                R3 = bytes[index++],
                ProviderUID = new FlatUID_r
                {
                    Ab = new byte[Constants.FlatUIDByteSize]
                }
            };
            for (int i = 0; i < Constants.FlatUIDByteSize; i++)
            {
                entryID.ProviderUID.Ab[i] = bytes[index++];
            }

            // R4: 4 bytes
            entryID.R4 = (uint)BitConverter.ToInt32(bytes, index);
            index += 4;

            // DisplayType: 4 bytes
            entryID.DisplayTypeString = (DisplayTypeValue)BitConverter.ToInt32(bytes, index);
            index += 4;

            // DistinguishedName: variable 
            entryID.DistinguishedName = System.Text.Encoding.Default.GetString(bytes, index, bytes.Length - index - 1);
            return entryID;
        }

        /// <summary>
        /// Parse a STAT structure instance from pointer.
        /// </summary>
        /// <param name="ptr">Pointer points to the memory.</param>
        /// <returns>A STAT structure.</returns>
        public static STAT ParseStat(IntPtr ptr)
        {
            STAT stat = new STAT();
            int offset = 0;

            stat.SortType = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.ContainerID = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.CurrentRec = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.Delta = Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.NumPos = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.TotalRecs = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.CodePage = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.TemplateLocale = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(int);

            stat.SortLocale = (uint)Marshal.ReadInt32(ptr, offset);

            return stat;
        }

        /// <summary>
        /// Parse Binary_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of Binary_r structure.</returns>
        public static Binary_r ParseBinary_r(IntPtr ptr)
        {
            Binary_r b_r = new Binary_r
            {
                Cb = (uint)Marshal.ReadInt32(ptr)
            };

            if (b_r.Cb == 0)
            {
                b_r.Lpb = null;
            }
            else
            {
                b_r.Lpb = new byte[b_r.Cb];
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr baddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    for (uint i = 0; i < b_r.Cb; i++)
                    {
                        b_r.Lpb[i] = Marshal.ReadByte(baddr, (int)i);
                    }
                }
                else
                {
                    for (uint i = 0; i < b_r.Cb; i++)
                    {
                        b_r.Lpb[i] = Marshal.ReadByte(ptr, (int)i + sizeof(uint));
                    }
                }
            }

            return b_r;
        }

        /// <summary>
        /// Parse GUIDs.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of GUIDs.</returns>
        public static FlatUID_r ParseFlatUID_r(IntPtr ptr)
        {
            FlatUID_r fuid_r = new FlatUID_r
            {
                Ab = new byte[Constants.FlatUIDByteSize]
            };
            for (int i = 0; i < Constants.FlatUIDByteSize; i++)
            {
                fuid_r.Ab[i] = Marshal.ReadByte(ptr, i);
            }

            return fuid_r;
        }

        /// <summary>
        /// Parse ShortArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of ShortArray_r structure.</returns>
        public static ShortArray_r ParseShortArray_r(IntPtr ptr)
        {
            ShortArray_r shortArray = new ShortArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (shortArray.CValues == 0)
            {
                shortArray.Lpi = null;
            }
            else
            {
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr saaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    ptr = saaddr;
                }

                shortArray.Lpi = new short[shortArray.CValues];
                int offset = 0;
                for (uint i = 0; i < shortArray.CValues; i++, offset += sizeof(short))
                {
                    shortArray.Lpi[i] = Marshal.ReadInt16(ptr, offset);
                }
            }

            return shortArray;
        }

        /// <summary>
        /// Parse LongArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of LongArray_r structure.</returns>
        public static LongArray_r ParseLongArray_r(IntPtr ptr)
        {
            LongArray_r longArray = new LongArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (longArray.CValues == 0)
            {
                longArray.Lpl = null;
            }
            else
            {
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr laaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    ptr = laaddr;
                }

                longArray.Lpl = new int[longArray.CValues];
                int offset = 0;
                for (uint i = 0; i < longArray.CValues; i++, offset += sizeof(int))
                {
                    longArray.Lpl[i] = Marshal.ReadInt32(ptr, offset);
                }
            }

            return longArray;
        }

        /// <summary>
        /// Parse String_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of String_r structure.</returns>
        public static byte[] ParseString_r(IntPtr ptr)
        {
            List<byte> stringArray = new List<byte>();
            if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
            {
                IntPtr szaPtr = new IntPtr(Marshal.ReadInt32(ptr));

                if (szaPtr == IntPtr.Zero)
                {
                    return null;
                }

                ptr = szaPtr;
            }

            int offsetOfszA = 0;
            byte curValueOfszA = 0;
            ArrayList listOfszA = new ArrayList();
            while (Marshal.ReadByte(ptr, offsetOfszA) != '\0')
            {
                curValueOfszA = Marshal.ReadByte(ptr, offsetOfszA);
                offsetOfszA++;
                listOfszA.Add(curValueOfszA);
            }

            for (int i = 0; i < listOfszA.Count; i++)
            {
                stringArray.Add(byte.Parse(listOfszA[i].ToString()));
            }

            return stringArray.ToArray();
        }

        /// <summary>
        /// Parse WString_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of WString_r structure.</returns>
        public static byte[] ParseWString_r(IntPtr ptr)
        {
            List<byte> stringArray = new List<byte>();

            if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
            {
                IntPtr szwPtr = new IntPtr(Marshal.ReadInt32(ptr));

                if (szwPtr == IntPtr.Zero)
                {
                    return null;
                }

                ptr = szwPtr;
            }

            int offsetOfszW = 0;
            byte curValueOfszW = 0;
            short shortValueOfszW = Marshal.ReadInt16(ptr, offsetOfszW);
            ArrayList listOfszW = new ArrayList();

            while (shortValueOfszW != '\0')
            {
                curValueOfszW = Marshal.ReadByte(ptr, offsetOfszW);
                offsetOfszW++;
                listOfszW.Add(curValueOfszW);

                if (offsetOfszW % 2 == 0)
                {
                    shortValueOfszW = Marshal.ReadInt16(ptr, offsetOfszW);
                }
            }

            for (int i = 0; i < listOfszW.Count; i++)
            {
                stringArray.Add(byte.Parse(listOfszW[i].ToString()));
            }

            return stringArray.ToArray();
        }

        /// <summary>
        /// Parse StringArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of StringArray_r structure.</returns>
        public static StringArray_r ParseStringArray_r(IntPtr ptr)
        {
            StringArray_r stringArray = new StringArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (stringArray.CValues == 0)
            {
                stringArray.LppszA = null;
            }
            else
            {
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr szaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    stringArray.LppszA = new string[stringArray.CValues];
                    int offset = 0;
                    for (uint i = 0; i < stringArray.CValues; i++)
                    {
                        stringArray.LppszA[i] = Marshal.PtrToStringAnsi(new IntPtr(Marshal.ReadInt32(szaddr, offset)));
                        offset += 4;
                    }
                }
                else
                {
                    stringArray.LppszA = new string[stringArray.CValues];
                    for (uint i = 0; i < stringArray.CValues; i++)
                    {
                        stringArray.LppszA[i] = BitConverter.ToString(ParseString_r(ptr));
                    }
                }
            }

            return stringArray;
        }

        /// <summary>
        /// Parse BinaryArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of BinaryArray_r structure.</returns>
        public static BinaryArray_r ParseBinaryArray_r(IntPtr ptr)
        {
            BinaryArray_r binaryArray = new BinaryArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (binaryArray.CValues == 0)
            {
                binaryArray.Lpbin = null;
            }
            else
            {
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr baaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    binaryArray.Lpbin = new Binary_r[binaryArray.CValues];
                    for (uint i = 0; i < binaryArray.CValues; i++)
                    {
                        binaryArray.Lpbin[i] = ParseBinary_r(baaddr);
                        baaddr = new IntPtr(baaddr.ToInt32() + 8);
                    }
                }
                else
                {
                    binaryArray.Lpbin = new Binary_r[binaryArray.CValues];
                    for (uint i = 0; i < binaryArray.CValues; i++)
                    {
                        binaryArray.Lpbin[i] = ParseBinary_r(ptr);
                    }
                }
            }

            return binaryArray;
        }

        /// <summary>
        /// Parse WStringArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of WStringArray_r structure.</returns>
        public static WStringArray_r ParseWStringArray_r(IntPtr ptr)
        {
            WStringArray_r wsa_r = new WStringArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (wsa_r.CValues == 0)
            {
                wsa_r.LppszW = null;
            }
            else
            {
                if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                {
                    IntPtr szwaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                    wsa_r.LppszW = new string[wsa_r.CValues];
                    for (uint i = 0; i < wsa_r.CValues; i++)
                    {
                        wsa_r.LppszW[i] = Marshal.PtrToStringUni(new IntPtr(Marshal.ReadInt32(szwaddr)));
                        szwaddr = new IntPtr(szwaddr.ToInt32() + 4);
                    }
                }
                else
                {
                    wsa_r.LppszW = new string[wsa_r.CValues];
                    for (uint i = 0; i < wsa_r.CValues; i++)
                    {
                        wsa_r.LppszW[i] = BitConverter.ToString(ParseWString_r(ptr));
                    }
                }
            }

            return wsa_r;
        }

        /// <summary>
        /// Parse FlatUIDArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of FlatUIDArray_r structure.</returns>
        public static FlatUIDArray_r ParseFlatUIDArray_r(IntPtr ptr)
        {
            FlatUIDArray_r fuida_r = new FlatUIDArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            if (fuida_r.CValues == 0)
            {
                fuida_r.Lpguid = null;
            }
            else
            {
                fuida_r.Lpguid = new FlatUID_r[fuida_r.CValues];
                IntPtr fuidaddr = new IntPtr(Marshal.ReadInt32(ptr, sizeof(uint)));
                for (uint i = 0; i < fuida_r.CValues; i++)
                {
                    fuida_r.Lpguid[i] = ParseFlatUID_r(new IntPtr(Marshal.ReadInt32(fuidaddr)));
                    fuidaddr = new IntPtr(fuidaddr.ToInt32() + 4);
                }
            }

            return fuida_r;
        }

        /// <summary>
        /// Parse PROP_VAL_UNION structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <param name="proptype">Property Types.</param>
        /// <returns>Instance of PROP_VAL_UNION structure.</returns>
        public static PROP_VAL_UNION ParsePROP_VAL_UNION(IntPtr ptr, PropertyTypeValue proptype)
        {
            PROP_VAL_UNION pvu = new PROP_VAL_UNION();

            switch (proptype)
            {
                case PropertyTypeValue.PtypInteger16:
                    pvu.I = Marshal.ReadInt16(ptr);
                    break;

                case PropertyTypeValue.PtypInteger32:
                    pvu.L = Marshal.ReadInt32(ptr);
                    break;

                case PropertyTypeValue.PtypBoolean:
                    if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                    {
                        pvu.B = (ushort)Marshal.ReadInt16(ptr);
                    }
                    else
                    {
                        pvu.B = (byte)Marshal.ReadByte(ptr);
                    }

                    break;

                case PropertyTypeValue.PtypString8:
                    pvu.LpszA = ParseString_r(ptr);
                    break;

                case PropertyTypeValue.PtypBinary:
                    pvu.Bin = ParseBinary_r(ptr);
                    break;

                case PropertyTypeValue.PtypString:
                    pvu.LpszW = ParseWString_r(ptr);
                    break;

                case PropertyTypeValue.PtypGuid:
                    if (AdapterHelper.Transport.ToLower(System.Globalization.CultureInfo.CurrentCulture) != "mapi_http")
                    {
                        IntPtr uidaddr = new IntPtr(Marshal.ReadInt32(ptr));

                        if (uidaddr == IntPtr.Zero)
                        {
                            pvu.Lpguid = null;
                        }
                        else
                        {
                            pvu.Lpguid = new FlatUID_r[1];
                            pvu.Lpguid[0] = ParseFlatUID_r(uidaddr);
                        }
                    }
                    else
                    {
                        if (ptr == IntPtr.Zero)
                        {
                            pvu.Lpguid = null;
                        }
                        else
                        {
                            pvu.Lpguid = new FlatUID_r[1];
                            pvu.Lpguid[0] = ParseFlatUID_r(ptr);
                        }
                    }

                    break;

                case PropertyTypeValue.PtypTime:
                    pvu.Ft.LowDateTime = (uint)Marshal.ReadInt32(ptr);
                    pvu.Ft.HighDateTime = (uint)Marshal.ReadInt32(ptr, sizeof(uint));
                    break;

                case PropertyTypeValue.PtypErrorCode:
                    pvu.Err = Marshal.ReadInt32(ptr);
                    break;

                case PropertyTypeValue.PtypMultipleInteger16:
                    pvu.MVi = ParseShortArray_r(ptr);
                    break;

                case PropertyTypeValue.PtypMultipleInteger32:
                    pvu.MVl = ParseLongArray_r(ptr);
                    break;

                case PropertyTypeValue.PtypMultipleString8:
                    uint isFound = (uint)Marshal.ReadInt32(ptr);

                    if (isFound == (uint)ErrorCodeValue.NotFound)
                    {
                        pvu.MVszA.CValues = 0;
                        pvu.MVszA.LppszA = null;
                    }
                    else
                    {
                        pvu.MVszA = ParseStringArray_r(ptr);
                    }

                    break;

                case PropertyTypeValue.PtypMultipleBinary:
                    pvu.MVbin = ParseBinaryArray_r(ptr);
                    break;

                case PropertyTypeValue.PtypMultipleString:
                    pvu.MVszW = ParseWStringArray_r(ptr);
                    break;

                case PropertyTypeValue.PtypMultipleGuid:
                    pvu.MVguid = ParseFlatUIDArray_r(ptr);
                    break;

                case PropertyTypeValue.PtypNull:
                case PropertyTypeValue.PtypEmbeddedTable:
                    pvu.Reserved = Marshal.ReadInt32(ptr);
                    break;

                default:
                    throw new ParseException("Parsing PROP_VAL_UNION failed!");
            }

            return pvu;
        }

        /// <summary>
        /// Parse PROP_VAL_UNION structure.
        /// </summary>
        /// <param name="propertyValue">The property value used for parsing.</param>
        /// <param name="proptype">The Property Types used for parsing.</param>
        /// <returns>Instance of PROP_VAL_UNION structure.</returns>
        public static PROP_VAL_UNION ParsePROP_VAL_UNION(PropertyValue propertyValue, PropertyTypeValue proptype)
        {
            PROP_VAL_UNION pvu = new PROP_VAL_UNION();
            IntPtr valuePtr = IntPtr.Zero;
            valuePtr = Marshal.AllocHGlobal(propertyValue.Value.Length);
            Marshal.Copy(propertyValue.Value, 0, valuePtr, propertyValue.Value.Length);
            pvu = AdapterHelper.ParsePROP_VAL_UNION(valuePtr, proptype);
            Marshal.FreeHGlobal(valuePtr);
            return pvu;
        }

        /// <summary>
        /// Parse PropertyValue_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of PropertyValue_r structure.</returns>
        public static PropertyValue_r ParsePropertyValue_r(IntPtr ptr)
        {
            PropertyValue_r protertyValue = new PropertyValue_r();

            int offset = 0;

            protertyValue.PropTag = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(uint);

            protertyValue.Reserved = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(uint);

            IntPtr newPtr = new IntPtr(ptr.ToInt32() + offset);
            protertyValue.Value = ParsePROP_VAL_UNION(newPtr, (PropertyTypeValue)(protertyValue.PropTag & 0x0000FFFF));

            return protertyValue;
        }

        /// <summary>
        /// Parse PropertyRow_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of PropertyRow_r structure.</returns>
        public static PropertyRow_r ParsePropertyRow_r(IntPtr ptr)
        {
            PropertyRow_r protertyRow = new PropertyRow_r();
            int offset = 0;

            protertyRow.Reserved = (uint)Marshal.ReadInt32(ptr);
            offset += sizeof(uint);

            protertyRow.CValues = (uint)Marshal.ReadInt32(ptr, offset);
            offset += sizeof(uint);

            if (protertyRow.CValues == 0)
            {
                protertyRow.LpProps = null;
            }
            else
            {
                protertyRow.LpProps = new PropertyValue_r[protertyRow.CValues];
                IntPtr pvaddr = new IntPtr(Marshal.ReadInt32(ptr, offset));

                const int PropertyValueLengthInBytes = 16;

                for (uint i = 0; i < protertyRow.CValues; i++)
                {
                    protertyRow.LpProps[i] = ParsePropertyValue_r(pvaddr);
                    pvaddr = new IntPtr(pvaddr.ToInt32() + PropertyValueLengthInBytes);
                }
            }

            return protertyRow;
        }

        /// <summary>
        /// Parse PropertyRowSet_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of PropertyRowSet_r structure.</returns>
        public static PropertyRowSet_r ParsePropertyRowSet_r(IntPtr ptr)
        {
            PropertyRowSet_r prs_r = new PropertyRowSet_r();
            int offset = 0;

            prs_r.CRows = (uint)Marshal.ReadInt32(ptr);
            offset += sizeof(uint);

            const int PropertyRowLengthInBytes = 12;

            if (prs_r.CRows == 0)
            {
                prs_r.ARow = null;
            }
            else
            {
                ptr = new IntPtr(ptr.ToInt32() + offset);
                prs_r.ARow = new PropertyRow_r[prs_r.CRows];

                for (uint i = 0; i < prs_r.CRows; i++)
                {
                    prs_r.ARow[i] = ParsePropertyRow_r(ptr);
                    ptr = new IntPtr(ptr.ToInt32() + PropertyRowLengthInBytes);
                }
            }

            return prs_r;
        }

        /// <summary>
        /// Parse PropertyTagArray_r structure.
        /// </summary>
        /// <param name="ptr">A pointer points to the allocated memory.</param>
        /// <returns>Instance of PropertyRowSet_r structure.</returns>
        public static PropertyTagArray_r ParsePropertyTagArray_r(IntPtr ptr)
        {
            int offset = 0;
            PropertyTagArray_r pta_r = new PropertyTagArray_r
            {
                CValues = (uint)Marshal.ReadInt32(ptr)
            };

            offset += sizeof(uint);

            if (pta_r.CValues == 0)
            {
                pta_r.AulPropTag = null;
            }
            else
            {
                pta_r.AulPropTag = new uint[pta_r.CValues];

                for (int i = 0; i < pta_r.CValues; i++)
                {
                    pta_r.AulPropTag[i] = (uint)Marshal.ReadInt32(ptr, offset);
                    offset += sizeof(uint);
                }
            }

            return pta_r;
        }

        /// <summary>
        /// Extract the dictionary for display name and PermanentEntryID.
        /// </summary>
        /// <param name="rows">The returned rows which contain PermanentEntryID.</param>
        /// <returns>The dictionary for display name and PermanentEntryID.</returns>
        public static Dictionary<string, PermanentEntryID?> ExtractPermanentEntryIDAndDisplayname(PropertyRowSet_r? rows)
        {
            Dictionary<string, PermanentEntryID?> dict = new Dictionary<string, PermanentEntryID?>();
            string displayName = string.Empty;
            PermanentEntryID? permanentEntryID = null;

            foreach (PropertyRow_r row in rows.Value.ARow)
            {
                foreach (PropertyValue_r propertyValue in row.LpProps)
                {
                    Site.Log.Add(LogEntryKind.Debug, "The PropTag value is {0}.", (AulProp)propertyValue.PropTag);
                    switch ((AulProp)propertyValue.PropTag)
                    {
                        case AulProp.PidTagDisplayName:
                            displayName = System.Text.Encoding.UTF8.GetString(propertyValue.Value.LpszA);
                            break;
                        case AulProp.PidTagEntryId:
                            permanentEntryID = AdapterHelper.ParsePermanentEntryIDFromBytes(propertyValue.Value.Bin.Lpb);
                            break;

                        default:
                            break;
                    }
                }
            }

            Site.Log.Add(LogEntryKind.Debug, "Got the Address Creation Templates object {0}.", displayName);
            Site.Assert.IsNotNull(permanentEntryID, "The Permanent Entry ID of the {0} should not null.", displayName);
            dict.Add(displayName, permanentEntryID);
            permanentEntryID = null;

            return dict;
        }

        /// <summary>
        /// Parse PropertyRowSet_r structure.
        /// </summary>
        /// <param name="columns">The columns which contain property tags.</param>
        /// <param name="rowCount">The row count of the PropertyRowSet_r.</param>
        /// <param name="rowData">The row data which contain the property values.</param>
        /// <returns>Instance of PropertyRowSet_r structure.</returns>
        public static PropertyRowSet_r ParsePropertyRowSet_r(LargePropTagArray columns, uint rowCount, AddressBookPropertyRow[] rowData)
        {
            PropertyRowSet_r propertyRowSet_r = new PropertyRowSet_r();

            propertyRowSet_r.CRows = rowCount;
            propertyRowSet_r.ARow = new PropertyRow_r[rowCount];

            for (int i = 0; i < rowCount; i++)
            {
                propertyRowSet_r.ARow[i].Reserved = 0;
                propertyRowSet_r.ARow[i].CValues = columns.PropertyTagCount;
                propertyRowSet_r.ARow[i].LpProps = new PropertyValue_r[columns.PropertyTagCount];
                for (int j = 0; j < columns.PropertyTagCount; j++)
                {
                    propertyRowSet_r.ARow[i].LpProps[j].PropTag = BitConverter.ToUInt32(columns.PropertyTags[j].Serialize(), 0);
                    propertyRowSet_r.ARow[i].LpProps[j].Reserved = 0;
                    propertyRowSet_r.ARow[i].LpProps[j].Value = AdapterHelper.ParsePROP_VAL_UNION(rowData[i].ValueArray.ToArray()[j], (PropertyTypeValue)columns.PropertyTags[j].PropertyType);
                }
            }

            return propertyRowSet_r;
        }

        /// <summary>
        /// Parse PropertyRowSet_r structure.
        /// </summary>
        /// <param name="rowsCount">The row count of the PropertyRowSet_r.</param>
        /// <param name="rows">The rows which contains property tags and property values.</param>
        /// <returns>Instance of PropertyRowSet_r structure.</returns>
        public static PropertyRowSet_r ParsePropertyRowSet_r(uint rowsCount, AddressBookPropValueList[] rows)
        {
            PropertyRowSet_r propertyRowSet_r = new PropertyRowSet_r();

            propertyRowSet_r.CRows = rowsCount;
            if (rowsCount == 0)
            {
                propertyRowSet_r.ARow = null;
            }
            else
            {
                propertyRowSet_r.ARow = new PropertyRow_r[rowsCount];

                for (int i = 0; i < rowsCount; i++)
                {
                    propertyRowSet_r.ARow[i].Reserved = 0;
                    propertyRowSet_r.ARow[i].CValues = rows[i].PropertyValueCount;
                    propertyRowSet_r.ARow[i].LpProps = new PropertyValue_r[rows[i].PropertyValueCount];
                    for (int j = 0; j < rows[i].PropertyValueCount; j++)
                    {
                        propertyRowSet_r.ARow[i].LpProps[j].PropTag = BitConverter.ToUInt32(rows[i].PropertyValues[j].PropertyTag.Serialize(), 0);
                        propertyRowSet_r.ARow[i].LpProps[j].Reserved = 0;
                        propertyRowSet_r.ARow[i].LpProps[j].Value = AdapterHelper.ParsePROP_VAL_UNION((PropertyValue)rows[i].PropertyValues[j], (PropertyTypeValue)rows[i].PropertyValues[j].PropertyTag.PropertyType);
                    }
                }
            }

            return propertyRowSet_r;
        }

        /// <summary>
        /// Parse the PropertyRow_r structure.
        /// </summary>
        /// <param name="row">The row which contains property tags and property values</param>
        /// <returns>Instance of the PropertyRow_r.</returns>
        public static PropertyRow_r ParsePropertyRow_r(AddressBookPropValueList row)
        {
            PropertyRow_r propertyRow_r = new PropertyRow_r();
            propertyRow_r.Reserved = 0;
            propertyRow_r.CValues = row.PropertyValueCount;
            propertyRow_r.LpProps = new PropertyValue_r[row.PropertyValueCount];
            for (int i = 0; i < row.PropertyValueCount; i++)
            {
                propertyRow_r.LpProps[i].PropTag = BitConverter.ToUInt32(row.PropertyValues[i].PropertyTag.Serialize(), 0);
                propertyRow_r.LpProps[i].Reserved = 0;
                propertyRow_r.LpProps[i].Value = AdapterHelper.ParsePROP_VAL_UNION((PropertyValue)row.PropertyValues[i], (PropertyTypeValue)row.PropertyValues[i].PropertyTag.PropertyType);
            }

            return propertyRow_r;
        }

        /// <summary>
        /// Parse the PropertyTagArray_r structure. 
        /// </summary>
        /// <param name="minimalIdCount">The count of the minimal IDs.</param>
        /// <param name="minimalIds">The minimal IDs used by PropertyTagArray_r.</param>
        /// <returns>Instance of the PropertyTagArray_r.</returns>
        public static PropertyTagArray_r ParsePropertyTagArray_r(uint minimalIdCount, uint[] minimalIds)
        {
            PropertyTagArray_r propertyTagArray_r = new PropertyTagArray_r();
            propertyTagArray_r.CValues = minimalIdCount;
            propertyTagArray_r.AulPropTag = minimalIds;
            return propertyTagArray_r;
        }
        #endregion
        #endregion
    }
}