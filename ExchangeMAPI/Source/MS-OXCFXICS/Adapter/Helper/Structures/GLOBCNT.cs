namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// An auto-incrementing 6-byte value.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct GLOBCNT
    {
        /// <summary>
        /// The 1st byte of GLOBCNT.
        /// </summary>
        [MarshalAs(UnmanagedType.U1)]
        public byte Byte1;

        /// <summary>
        /// The 2nd byte of GLOBCNT.
        /// </summary>
        [MarshalAs(UnmanagedType.U1)]
        public byte Byte2;

        /// <summary>
        /// The 3rd byte of GLOBCNT.
        /// </summary>
        [MarshalAs(UnmanagedType.U1)]
        public byte Byte3;

        /// <summary>
        /// The 4th byte of GLOBCNT.
        /// </summary>
        [MarshalAs(UnmanagedType.U1)]
        public byte Byte4;

        /// <summary>
        /// The 5th byte of GLOBCNT.
        /// </summary>
        [MarshalAs(UnmanagedType.U1)]
        public byte Byte5;

        /// <summary>
        /// The 6th byte of GLOBCNT.
        /// </summary>
        [MarshalAs(UnmanagedType.U1)]
        public byte Byte6;

        /// <summary>
        /// Gets the byte at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index of bytes.</param>
        /// <returns>The byte at the specified index.</returns>
        public byte this[int index]
        { 
            get
            {
                if (index < 0 || index > 5)
                {
                    AdapterHelper.Site.Assert.Fail("The index is out of range.");
                }

                byte[] tmp = StructureSerializer.Serialize(this);
                return tmp[index];
            }
        }

        /// <summary>
        /// Indicates whether the 1st instance is greater than the 2nd one.
        /// </summary>
        /// <param name="item1">The 1st GLOBCNT instance to compare.</param>
        /// <param name="item2">The 2st GLOBCNT instance to compare.</param>
        /// <returns>True if the 1st instance is greater than the 2nd one.</returns>
        public static bool operator >(GLOBCNT item1, GLOBCNT item2)
        {
            byte[] tmp1 = StructureSerializer.Serialize(item1);
            byte[] tmp2 = StructureSerializer.Serialize(item2);
            for (int i = 0; i < tmp1.Length; i++)
            {
                if (tmp1[i] == tmp2[i])
                {
                    continue;
                }
                else
                {
                    return tmp1[i] > tmp2[i];
                }
            }

            return false;
        }

        /// <summary>
        /// Indicates whether the 1st instance is less than the 2nd one
        /// </summary>
        /// <param name="item1">The 1st GLOBCNT instance to compare.</param>
        /// <param name="item2">The 2st GLOBCNT instance to compare.</param>
        /// <returns>True if the 1st instance is less than the 2nd one.</returns>
        public static bool operator <(GLOBCNT item1, GLOBCNT item2)
        {
            byte[] tmp1 = StructureSerializer.Serialize(item1);
            byte[] tmp2 = StructureSerializer.Serialize(item2);
            for (int i = 0; i < tmp1.Length; i++)
            {
                if (tmp1[i] == tmp2[i])
                {
                    continue;
                }
                else
                {
                    return tmp1[i] < tmp2[i];
                }
            }

            return false;
        }

        /// <summary>
        /// Indicates whether two GLOBCNT instances and a specified object are equal.
        /// </summary>
        /// <param name="item1">The 1st GLOBCNT instance to compare.</param>
        /// <param name="item2">The 2st GLOBCNT instance to compare.</param>
        /// <returns>True if the two instances represent the same value.</returns>
        public static bool operator ==(GLOBCNT item1, GLOBCNT item2)
        {
            byte[] tmp1 = StructureSerializer.Serialize(item1);
            byte[] tmp2 = StructureSerializer.Serialize(item2);
            for (int i = 0; i < tmp1.Length; i++)
            {
                if (tmp1[i] == tmp2[i])
                {
                    continue;
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Indicates whether two GLOBCNT instances and a specified object are not equal.
        /// </summary>
        /// <param name="item1">The 1st GLOBCNT instance to compare.</param>
        /// <param name="item2">The 2st GLOBCNT instance to compare.</param>
        /// <returns>True if the two instances do represent different values.</returns>
        public static bool operator !=(GLOBCNT item1, GLOBCNT item2)
        {
            byte[] tmp1 = StructureSerializer.Serialize(item1);
            byte[] tmp2 = StructureSerializer.Serialize(item2);
            for (int i = 0; i < tmp1.Length; i++)
            {
                if (tmp1[i] == tmp2[i])
                {
                    continue;
                }
                else
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Indicates whether the 1st instance is greater than or equals to the 2nd one
        /// </summary>
        /// <param name="item1">The 1st GLOBCNT instance to compare.</param>
        /// <param name="item2">The 2st GLOBCNT instance to compare.</param>
        /// <returns>True if the 1st instance is greater than or equals to the 2nd one.</returns>
        public static bool operator >=(GLOBCNT item1, GLOBCNT item2)
        {
            return item1 > item2 || item1 == item2;
        }

        /// <summary>
        /// Indicates whether the 1st instance is less than or equals to the 2nd one.
        /// </summary>
        /// <param name="item1">The 1st GLOBCNT instance to compare.</param>
        /// <param name="item2">The 2st GLOBCNT instance to compare.</param>
        /// <returns>True if the 1st instance is less than or equals to the 2nd one.</returns>
        public static bool operator <=(GLOBCNT item1, GLOBCNT item2)
        {
            return item1 < item2 || item1 == item2;
        }

        /// <summary>
        /// Plus a GLOBCNT by 1.
        /// </summary>
        /// <param name="g">A GLOBCNT instance.</param>
        /// <returns>The next GLOBCNT value.</returns>
        public static GLOBCNT Inc(GLOBCNT g)
        {
            byte c = 0;
            byte[] tmp = StructureSerializer.Serialize(g);
            if (tmp[5] == byte.MaxValue)
            {
                tmp[5] = 0;
                c = 1;
            }
            else
            {
                tmp[5]++;
                c = 0;
            }

            for (int i = 4; i > -1; i--)
            {
                if (tmp[i] == byte.MaxValue && c == 1)
                {
                    tmp[i] = 0;
                    c = 1;
                }
                else
                {
                    tmp[i] += c;
                    c = 0;
                }
            }

            return StructureSerializer.Deserialize<GLOBCNT>(tmp);
        }

        /// <summary>
        /// Indicates whether this instance and a specified object are equal.
        /// </summary>
        /// <param name="obj">Another object to compare to.</param>
        /// <returns>True if object and this instance represent the same value.</returns>
        public override bool Equals(object obj)
        {
            byte[] tmp1 = StructureSerializer.Serialize(obj);
            byte[] tmp2 = StructureSerializer.Serialize(this);
            for (int i = 0; i < tmp1.Length; i++)
            {
                if (tmp1[i] != tmp2[i])
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Returns the hash code for this instance.
        /// </summary>
        /// <returns>A 32-bit signed integer that is the hash code for this instance.</returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}