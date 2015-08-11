namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Data Element
    /// </summary>
    public class DataElement : StreamObject
    {
        /// <summary>
        /// Data Element Data Type Mapping
        /// </summary>
        private static Dictionary<DataElementType, Type> dataElementDataTypeMapping = null;

        /// <summary>
        /// Initializes static members of the DataElement class
        /// </summary>
        static DataElement()
        {
            dataElementDataTypeMapping = new Dictionary<DataElementType, Type>();
            Type dataElementEnumType = typeof(DataElementType);
            foreach (object value in dataElementEnumType.GetEnumValues())
            {
                dataElementDataTypeMapping.Add((DataElementType)value, Type.GetType("Microsoft.Protocols.TestSuites.SharedAdapter." + dataElementEnumType.GetEnumName(value)));
            }
        }

        /// <summary>
        /// Initializes a new instance of the DataElement class.
        /// </summary>
        /// <param name="type">data element type</param>
        /// <param name="data">Specifies the data of the element.</param>
        public DataElement(DataElementType type, DataElementData data)
            : base(StreamObjectTypeHeaderStart.DataElement)
        {
            if (!dataElementDataTypeMapping.Keys.Contains(type))
            {
                throw new InvalidOperationException("Invalid argument type value" + (int)type);
            }

            this.DataElementType = type;
            this.Data = data;
            this.DataElementExtendedGUID = new ExGuid(SequenceNumberGenerator.GetCurrentSerialNumber(), Guid.NewGuid());
            this.SerialNumber = new SerialNumber(Guid.NewGuid(), SequenceNumberGenerator.GetCurrentSerialNumber());
        }

        /// <summary>
        /// Initializes a new instance of the DataElement class.
        /// </summary>
        public DataElement()
            : base(StreamObjectTypeHeaderStart.DataElement)
        {
        }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the data element.
        /// </summary>
        public ExGuid DataElementExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets a serial number that specifies the data element.
        /// </summary>
        public SerialNumber SerialNumber { get; set; }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the value of the storage index data element type.
        /// </summary>
        public DataElementType DataElementType { get; set; }

        /// <summary>
        /// Gets or sets a data element fragment.
        /// </summary>
        public DataElementData Data { get; set; }

        /// <summary>
        /// Used to get data.
        /// </summary>
        /// <typeparam name="T">Type of element</typeparam>
        /// <returns>Data of the element</returns>
        public T GetData<T>()
            where T : DataElementData
        {
            if (this.Data is T)
            {
                return this.Data as T;
            }
            else
            {
                throw new InvalidOperationException(string.Format("Unable to cast DataElementData to the type {0}, its actual type is {1}", typeof(T).Name, this.Data.GetType().Name));
            }
        }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;

            try
            {
                this.DataElementExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
                this.SerialNumber = BasicObject.Parse<SerialNumber>(byteArray, ref index);
                this.DataElementType = (DataElementType)BasicObject.Parse<Compact64bitInt>(byteArray, ref index).DecodedValue;
            }
            catch (BasicObjectParseErrorException e)
            {
                throw new DataElementParseErrorException(index, e);
            }

            if (index - currentIndex != lengthOfItems)
            {
                throw new DataElementParseErrorException(currentIndex, "Failed to check the data element header length, whose value does not cover the DataElementExtendedGUID, SerialNumber and DataElementType", null);
            }

            if (dataElementDataTypeMapping.ContainsKey(this.DataElementType))
            {
                this.Data = Activator.CreateInstance(dataElementDataTypeMapping[this.DataElementType]) as DataElementData;

                try
                {
                    index += this.Data.DeserializeDataElementDataFromByteArray(byteArray, index);
                }
                catch (BasicObjectParseErrorException e1)
                {
                    throw new DataElementParseErrorException(index, e1);
                }
                catch (StreamObjectParseErrorException e2)
                {
                    throw new DataElementParseErrorException(index, e2);
                }
            }
            else
            {
                throw new DataElementParseErrorException(index, "Failed to create specific data element instance with the type " + this.DataElementType, null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The element length</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int startIndex = byteList.Count;
            byteList.AddRange(this.DataElementExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.SerialNumber.SerializeToByteList());
            byteList.AddRange(new Compact64bitInt((uint)this.DataElementType).SerializeToByteList());

            int headerLength = byteList.Count - startIndex;
            byteList.AddRange(this.Data.SerializeToByteList());

            return headerLength;
        }
    }
}