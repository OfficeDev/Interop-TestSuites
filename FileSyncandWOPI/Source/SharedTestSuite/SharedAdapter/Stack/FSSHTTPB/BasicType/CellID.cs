namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to specify the cell identifier.
    /// </summary>
    public class CellID : BasicObject
    {
        /// <summary>
        /// Initializes a new instance of the CellID class with specified ExGuids.
        /// </summary>
        /// <param name="extendGuid1">Specify the first ExGuid.</param>
        /// <param name="extendGuid2">Specify the second ExGuid.</param>
        public CellID(ExGuid extendGuid1, ExGuid extendGuid2)
        {
            this.ExtendGUID1 = extendGuid1;
            this.ExtendGUID2 = extendGuid2;
        }

        /// <summary>
        /// Initializes a new instance of the CellID class, this is the copy constructor.
        /// </summary>
        /// <param name="cellId">Specify the CellID.</param>
        public CellID(CellID cellId)
        {
            if (cellId.ExtendGUID1 != null)
            {
                this.ExtendGUID1 = new ExGuid(cellId.ExtendGUID1);
            }

            if (cellId.ExtendGUID2 != null)
            {
                this.ExtendGUID2 = new ExGuid(cellId.ExtendGUID2);
            }
        }

        /// <summary>
        /// Initializes a new instance of the CellID class, this is default constructor.
        /// </summary>
        public CellID()
        {
        }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the first cell identifier.
        /// </summary>
        public ExGuid ExtendGUID1 { get; set; }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the second cell identifier.
        /// </summary>
        public ExGuid ExtendGUID2 { get; set; }

        /// <summary>
        /// This method is used to convert the element of CellID basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of CellID.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.ExtendGUID1.SerializeToByteList());
            byteList.AddRange(this.ExtendGUID2.SerializeToByteList());
            return byteList;
        }

        /// <summary>
        /// Override the Equals method.
        /// </summary>
        /// <param name="obj">Specify the object.</param>
        /// <returns>Return true if equals, otherwise return false.</returns>
        public override bool Equals(object obj)
        {
            CellID another = obj as CellID;

            if (another == null)
            {
                return false;
            }

            if (another.ExtendGUID1 != null && another.ExtendGUID2 != null && this.ExtendGUID1 != null && this.ExtendGUID2 != null)
            {
                return another.ExtendGUID1.Equals(this.ExtendGUID1) && another.ExtendGUID2.Equals(this.ExtendGUID2);
            }

            return false;
        }

        /// <summary>
        /// Override the GetHashCode.
        /// </summary>
        /// <returns>Return the hash value.</returns>
        public override int GetHashCode()
        {
            return this.ExtendGUID1.GetHashCode() + this.ExtendGUID2.GetHashCode();
        }

        /// <summary>
        /// This method is used to deserialize the CellID basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the CellID basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;

            this.ExtendGUID1 = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.ExtendGUID2 = BasicObject.Parse<ExGuid>(byteArray, ref index);

            return index - startIndex;
        }
    }

    /// <summary>
    /// This class is used to specify the array of cell IDs.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class CellIDArray : BasicObject
    {
        /// <summary>
        /// Initializes a new instance of the CellIDArray class.
        /// </summary>
        /// <param name="count">Specify the number of CellID in the CellID array.</param>
        /// <param name="content">Specify the list of CellID.</param>
        public CellIDArray(ulong count, List<CellID> content)
        {
            this.Count = count;
            this.Content = content;
        }

        /// <summary>
        /// Initializes a new instance of the CellIDArray class, this is copy constructor.
        /// </summary>
        /// <param name="cellIdArray">Specify the CellIDArray.</param>
        public CellIDArray(CellIDArray cellIdArray)
        {
            this.Count = cellIdArray.Count;
            if (cellIdArray.Content != null)
            {
                foreach (CellID cellId in cellIdArray.Content)
                {
                    this.Content.Add(new CellID(cellId));
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the CellIDArray class, this is default constructor.
        /// </summary>
        public CellIDArray()
        {
            this.Content = new List<CellID>();
        }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the count of cell IDs in the array. 
        /// </summary>
        public ulong Count { get; set; }
        
        /// <summary>
        /// Gets or sets a cell ID list that specifies a list of cells.
        /// </summary>
        public List<CellID> Content { get; set; }

        /// <summary>
        /// This method is used to convert the element of CellIDArray basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of CellIDArray.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange((new Compact64bitInt(this.Count)).SerializeToByteList());
            if (this.Content != null)
            {
                foreach (CellID extendGuid in this.Content)
                {
                    byteList.AddRange(extendGuid.SerializeToByteList());
                }
            }

            return byteList;
        }

        /// <summary>
        /// This method is used to deserialize the CellIDArray basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the CellIDArray basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex) 
        {
            int index = startIndex;

            this.Count = BasicObject.Parse<Compact64bitInt>(byteArray, ref index).DecodedValue;

            for (ulong i = 0; i < this.Count; i++)
            {
                this.Content.Add(BasicObject.Parse<CellID>(byteArray, ref index));
            }

            return index - startIndex;
        }
    }
}