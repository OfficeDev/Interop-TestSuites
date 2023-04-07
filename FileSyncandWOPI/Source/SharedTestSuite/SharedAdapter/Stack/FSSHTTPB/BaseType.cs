namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// FSSHTTPB Serialize interface.
    /// </summary>
    public interface IFSSHTTPBSerializable
    {
        /// <summary>
        /// Serialize to byte list.
        /// </summary>
        /// <returns>The byte list.</returns>
        List<byte> SerializeToByteList();
    }

    /// <summary>
    /// Base object for FSSHTTPB.
    /// </summary>
    public abstract class BasicObject : IFSSHTTPBSerializable
    {
        /// <summary>
        /// Used to parse byte array to special object.
        /// </summary>
        /// <typeparam name="T">The type of target object.</typeparam>
        /// <param name="byteArray">The byte array contains raw data.</param>
        /// <param name="index">The index special where to start.</param>
        /// <returns>The instance of target object.</returns>
        public static T Parse<T>(byte[] byteArray, ref int index)
            where T : BasicObject, new()
        {
            T fsshttpbObject = Activator.CreateInstance<T>();
            try
            {
                index += fsshttpbObject.DeserializeFromByteArray(byteArray, index);
            }
            catch (InvalidOperationException e)
            {
                throw new BasicObjectParseErrorException(index, typeof(T).Name, e.Message, e);
            }

            return fsshttpbObject;
        }

        /// <summary>
        /// Used to return the length of this element.
        /// </summary>
        /// <param name="byteArray">The byte list.</param>
        /// <param name="startIndex">The start position.</param>
        /// <returns>The element length.</returns>
        public int DeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int length = this.DoDeserializeFromByteArray(byteArray, startIndex);

            // Invoke the basic object related capture code.
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                new MsfsshttpbAdapterCapture().InvokeCaptureMethod(this.GetType(), this, SharedContext.Current.Site);
            }

            return length;
        }

        /// <summary>
        /// Used to serialize item to byte list.
        /// </summary>
        /// <returns>The byte list.</returns>
        public abstract List<byte> SerializeToByteList();

        /// <summary>
        /// Used to return the length of this element.
        /// </summary>
        /// <param name="byteArray">The byte list.</param>
        /// <param name="startIndex">The start position.</param>
        /// <returns>The element length</returns>
        protected abstract int DoDeserializeFromByteArray(byte[] byteArray, int startIndex);
    }

    /// <summary>
    /// Base class of stream object.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public abstract class StreamObject : IFSSHTTPBSerializable
    {
        /// <summary>
        /// Hash set contains the StreamObjectTypeHeaderStart type.
        /// </summary>
        private static HashSet<StreamObjectTypeHeaderStart> compoundTypes = new HashSet<StreamObjectTypeHeaderStart>
        {
            StreamObjectTypeHeaderStart.DataElement,
            StreamObjectTypeHeaderStart.Knowledge,
            StreamObjectTypeHeaderStart.CellKnowledge,
            StreamObjectTypeHeaderStart.DataElementPackage,
            StreamObjectTypeHeaderStart.ObjectGroupDeclarations,
            StreamObjectTypeHeaderStart.ObjectGroupData,
            StreamObjectTypeHeaderStart.WaterlineKnowledge,
            StreamObjectTypeHeaderStart.ContentTagKnowledge,
            StreamObjectTypeHeaderStart.Request,
            StreamObjectTypeHeaderStart.FsshttpbSubResponse,
            StreamObjectTypeHeaderStart.SubRequest,
            StreamObjectTypeHeaderStart.ReadAccessResponse,
            StreamObjectTypeHeaderStart.SpecializedKnowledge,
            StreamObjectTypeHeaderStart.WriteAccessResponse,
            StreamObjectTypeHeaderStart.QueryChangesFilter,
            StreamObjectTypeHeaderStart.ResponseError,
            StreamObjectTypeHeaderStart.UserAgent,
            StreamObjectTypeHeaderStart.FragmentKnowledge,
            StreamObjectTypeHeaderStart.ObjectGroupMetadataDeclarations,
            StreamObjectTypeHeaderStart.LeafNodeObject,
            StreamObjectTypeHeaderStart.IntermediateNodeObject,
            StreamObjectTypeHeaderStart.FsshttpbResponse
        };

        /// <summary>
        /// The dictionary of StreamObjectTypeHeaderStart and type.
        /// </summary>
        private static Dictionary<StreamObjectTypeHeaderStart, Type> streamObjectTypeMapping = null;

        /// <summary>
        /// Initializes static members of the StreamObject class.
        /// </summary>
        static StreamObject()
        {
            streamObjectTypeMapping = new Dictionary<StreamObjectTypeHeaderStart, Type>();
            Type startHeaderEnumType = typeof(StreamObjectTypeHeaderStart);
            foreach (object value in startHeaderEnumType.GetEnumValues())
            {
                StreamObjectTypeMapping.Add((StreamObjectTypeHeaderStart)value, Type.GetType("Microsoft.Protocols.TestSuites.SharedAdapter." + startHeaderEnumType.GetEnumName(value)));
            }
        }

        /// <summary>
        /// Initializes a new instance of the StreamObject class.
        /// </summary>
        /// <param name="streamObjectType">The instance of StreamObjectTypeHeaderStart.</param>
        protected StreamObject(StreamObjectTypeHeaderStart streamObjectType)
        {
            this.StreamObjectType = streamObjectType;
        }

        /// <summary>
        /// Gets the StreamObjectTypeHeaderStart
        /// </summary>
        public static HashSet<StreamObjectTypeHeaderStart> CompoundTypes
        {
            get
            {
                return compoundTypes;
            }
        }

        /// <summary>
        /// Gets the StreamObjectTypeMapping
        /// </summary>
        public static Dictionary<StreamObjectTypeHeaderStart, Type> StreamObjectTypeMapping
        {
            get
            {
                return streamObjectTypeMapping;
            }
        }

        /// <summary>
        /// Gets the StreamObjectTypeHeaderStart.
        /// </summary>
        public StreamObjectTypeHeaderStart StreamObjectType { get; private set; }

        /// <summary>
        /// Gets the length of items.
        /// </summary>
        public int LengthOfItems { get; private set; }

        /// <summary>
        /// Gets or sets the stream object header start.
        /// </summary>
        internal StreamObjectHeaderStart StreamObjectHeaderStart { get; set; }

        /// <summary>
        /// Gets or sets the stream object header end.
        /// </summary>
        internal StreamObjectHeaderEnd StreamObjectHeaderEnd { get; set; }

        /// <summary>
        /// Get current stream object.
        /// </summary>
        /// <typeparam name="T">The type of target object.</typeparam>
        /// <param name="byteArray">The byte array which contains message.</param>
        /// <param name="index">The position where to start.</param>
        /// <returns>The current object instance.</returns>
        public static T GetCurrent<T>(byte[] byteArray, ref int index)
            where T : StreamObject
        {
            int tmpIndex = index;
            int length = 0;
            StreamObjectHeaderStart streamObjectHeader;
            if ((length = StreamObjectHeaderStart.TryParse(byteArray, tmpIndex, out streamObjectHeader)) == 0)
            {
                throw new StreamObjectParseErrorException(tmpIndex, typeof(T).Name, "Failed to extract either 16bit or 32bit stream object header in the current index.", null);
            }

            tmpIndex += length;

            StreamObject streamObject = ParseStreamObject(streamObjectHeader, byteArray, ref tmpIndex);

            if (!(streamObject is T))
            {
                throw new StreamObjectParseErrorException(tmpIndex, typeof(T).Name, string.Format("Failed to get stream object as expect type {0}, actual type is {1}", typeof(T).Name, StreamObjectTypeMapping[streamObjectHeader.Type].Name), null);
            }

            // Store the current index to the ref parameter index.
            index = tmpIndex;
            return streamObject as T;
        }

        /// <summary>
        /// Parse stream object from byte array.
        /// </summary>
        /// <param name="header">The instance of StreamObjectHeaderStart.</param>
        /// <param name="byteArray">The byte array.</param>
        /// <param name="index">The position where to start.</param>
        /// <returns>The instance of StreamObject.</returns>
        public static StreamObject ParseStreamObject(StreamObjectHeaderStart header, byte[] byteArray, ref int index)
        {
            if (StreamObjectTypeMapping.Keys.Contains(header.Type))
            {
                StreamObject streamObject = Activator.CreateInstance(StreamObjectTypeMapping[header.Type]) as StreamObject;
                
                try
                {
                    index += streamObject.DeserializeFromByteArray(header, byteArray, index);
                }
                catch (BasicObjectParseErrorException e)
                {
                    throw new StreamObjectParseErrorException(index, StreamObjectTypeMapping[header.Type].Name, e);
                }

                return streamObject;
            }

            int tmpIndex = index;
            tmpIndex -= header.HeaderType == StreamObjectHeaderStart.StreamObjectHeaderStart16bit ? 2 : 4;
            throw new StreamObjectParseErrorException(tmpIndex, "Unknown", string.Format("Failed to create the specified stream object instance, the type {0} of stream object header in the current index is not defined", (int)header.Type), null);
        }

        /// <summary>
        /// Try to get current object, true will returned if success.
        /// </summary>
        /// <typeparam name="T">The type of target object.</typeparam>
        /// <param name="byteArray">The byte array.</param>
        /// <param name="index">The position where to start.</param>
        /// <param name="streamObject">The instance that want to get.</param>
        /// <returns>The result of whether get success.</returns>
        public static bool TryGetCurrent<T>(byte[] byteArray, ref int index, out T streamObject)
            where T : StreamObject
        {
            int tmpIndex = index;
            streamObject = null;

            int length = 0;
            StreamObjectHeaderStart streamObjectHeader;
            if ((length = StreamObjectHeaderStart.TryParse(byteArray, tmpIndex, out streamObjectHeader)) == 0)
            {
                return false;
            }

            tmpIndex += length;
            if (StreamObjectTypeMapping.Keys.Contains(streamObjectHeader.Type) && StreamObjectTypeMapping[streamObjectHeader.Type] == typeof(T))
            {
                streamObject = ParseStreamObject(streamObjectHeader, byteArray, ref tmpIndex) as T;
            }
            else
            {
                return false;
            }

            index = tmpIndex;
            return true;
        }

        /// <summary>
        /// Serialize item to byte list.
        /// </summary>
        /// <returns>The byte list.</returns>
        public List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();

            int lengthOfItems = this.SerializeItemsToByteList(byteList);

            StreamObjectHeaderStart header;
            if ((int)this.StreamObjectType <= 0x3F && lengthOfItems <= 127)
            {
                header = new StreamObjectHeaderStart16bit(this.StreamObjectType, lengthOfItems);
            }
            else
            {
                header = new StreamObjectHeaderStart32bit(this.StreamObjectType, lengthOfItems);
            }

            byteList.InsertRange(0, header.SerializeToByteList());

            if (CompoundTypes.Contains(this.StreamObjectType))
            {
                if ((int)this.StreamObjectType <= 0x3F)
                {
                    byteList.AddRange(new StreamObjectHeaderEnd8bit((int)this.StreamObjectType).SerializeToByteList());
                }
                else
                {
                    byteList.AddRange(new StreamObjectHeaderEnd16bit((int)this.StreamObjectType).SerializeToByteList());
                }
            }

            return byteList;
        }

        /// <summary>
        /// Used to return the length of this element.
        /// </summary>
        /// <param name="header">Then instance of StreamObjectHeaderStart.</param>
        /// <param name="byteArray">The byte list</param>
        /// <param name="startIndex">The position where to start.</param>
        /// <returns>The element length</returns>
        public int DeserializeFromByteArray(StreamObjectHeaderStart header, byte[] byteArray, int startIndex)
        {
            this.StreamObjectType = header.Type;
            this.LengthOfItems = header.Length;

            if (header is StreamObjectHeaderStart32bit)
            {
                if (header.Length == 32767)
                {
                    this.LengthOfItems = (int)(header as StreamObjectHeaderStart32bit).LargeLength.DecodedValue;
                }
            }

            int index = startIndex;
            this.StreamObjectHeaderStart = header;
            this.DeserializeItemsFromByteArray(byteArray, ref index, this.LengthOfItems);

            if (CompoundTypes.Contains(this.StreamObjectType))
            {
                StreamObjectHeaderEnd end = null;
                BitReader bitReader = new BitReader(byteArray, index);
                int aField = bitReader.ReadInt32(2);
                if (aField == 0x1)
                {
                    end = BasicObject.Parse<StreamObjectHeaderEnd8bit>(byteArray, ref index);
                }
                if (aField == 0x3)
                {
                    end = BasicObject.Parse<StreamObjectHeaderEnd16bit>(byteArray, ref index);
                }

                if ((int)end.Type != (int)this.StreamObjectType)
                {
                    throw new StreamObjectParseErrorException(index, null, "Unexpected the stream header end value " + (int)this.StreamObjectType, null);
                }

                this.StreamObjectHeaderEnd = end;
            }

            // Capture all the type related requirements
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                new MsfsshttpbAdapterCapture().InvokeCaptureMethod(this.GetType(), this, SharedContext.Current.Site);
            }

            return index - startIndex;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected abstract int SerializeItemsToByteList(List<byte> byteList);

        /// <summary>
        /// De-serialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected abstract void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems);
    }

    /// <summary>
    /// Base class of data element.
    /// </summary>
    public abstract class DataElementData : IFSSHTTPBSerializable
    {
        /// <summary>
        /// De-serialize data element data from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array.</param>
        /// <param name="startIndex">The position where to start.</param>
        /// <returns>The length of the item.</returns>
        public abstract int DeserializeDataElementDataFromByteArray(byte[] byteArray, int startIndex);

        /// <summary>
        /// Serialize item to byte list.
        /// </summary>
        /// <returns>The byte list.</returns>
        public abstract List<byte> SerializeToByteList();
    }

    /// <summary>
    /// The base class of specialize knowledge data.
    /// </summary>
    public abstract class SpecializedKnowledgeData : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the SpecializedKnowledgeData class.
        /// </summary>
        /// <param name="headerType">The instance of StreamObjectTypeHeaderStart.</param>
        protected SpecializedKnowledgeData(StreamObjectTypeHeaderStart headerType)
            : base(headerType)
        {
        }
    }

    /// <summary>
    /// Base class of sub response data.
    /// </summary>
    public abstract class SubResponseData
    {
        /// <summary>
        /// The dictionary of sub response data type.
        /// </summary>
        public static readonly Dictionary<int, Type> SubResponseDataTypeMapping = new Dictionary<int, Type>
        {
            { 1, typeof(QueryAccessSubResponseData) },
            { 2, typeof(QueryChangesSubResponseData) },
            { 5, typeof(PutChangesSubResponseData) },
            { 11, typeof(AllocateExtendedGuidRangeSubResponseData) },
        };

        /// <summary>
        /// The reverse dictionary of sub response data type.
        /// </summary>
        public static readonly Dictionary<Type, RequestTypes> SubResponseDataTypeReverseMapping = new Dictionary<Type, RequestTypes>
        {
            { typeof(QueryAccessCellSubRequest), RequestTypes.QueryAccess },
            { typeof(QueryChangesCellSubRequest), RequestTypes.QueryChanges },
            { typeof(PutChangesCellSubRequest), RequestTypes.PutChanges },
            { typeof(AllocateExtendedGuidRangeCellSubRequest), RequestTypes.AllocateExtendedGuidRange },
        };

        /// <summary>
        /// Get current sub response data.
        /// </summary>
        /// <param name="type">The type that want to get.</param>
        /// <param name="byteArray">The byte array.</param>
        /// <param name="startIndex">The position where to start.</param>
        /// <returns>The instance of sub response data.</returns>
        public static SubResponseData GetCurrentSubResponseData(int type, byte[] byteArray, ref int startIndex)
        {
            int index = startIndex;
            if (!SubResponseDataTypeMapping.Keys.Contains(type))
            {
                throw new ResponseParseErrorException(-1, "Unexpected sub response type value" + type, null);
            }

            SubResponseData subResponseData = Activator.CreateInstance(SubResponseDataTypeMapping[type]) as SubResponseData;
            subResponseData.DeserializeSubResponseDataFromByteArray(byteArray, ref index);
            startIndex = index;
            return subResponseData;
        }

        /// <summary>
        /// De-serialize sub response data from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains sub response data.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        protected abstract void DeserializeSubResponseDataFromByteArray(byte[] byteArray, ref int currentIndex);
    }

    /// <summary>
    /// Base class for Response Error.
    /// </summary>
    public abstract class ErrorData : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ErrorData class.
        /// </summary>
        /// <param name="headerType">The instance of StreamObjectTypeHeaderStart.</param>
        protected ErrorData(StreamObjectTypeHeaderStart headerType)
            : base(headerType)
        {
        }

        /// <summary>
        /// Gets the Error detail information.
        /// </summary>
        public abstract string ErrorDetail { get; }
    }
}