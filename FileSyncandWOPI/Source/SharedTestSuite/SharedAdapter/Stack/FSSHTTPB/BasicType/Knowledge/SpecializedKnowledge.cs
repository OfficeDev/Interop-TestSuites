namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// This class specifies the SpecializedKnowledge.
    /// </summary>
    public class SpecializedKnowledge : StreamObject
    {
        /// <summary>
        /// A GUID value specifies a cell knowledge.
        /// </summary>
        public static readonly Guid CellKnowledgeGuid = new Guid("327A35F6-0761-4414-9686-51E900667A4D");

        /// <summary>
        /// A GUID value specifies a waterline knowledge.
        /// </summary>
        public static readonly Guid WaterlineKnowledgeGuid = new Guid("3A76E90E-8032-4D0C-B9DD-F3C65029433E");

        /// <summary>
        /// A GUID value specifies a fragment knowledge.
        /// </summary>
        public static readonly Guid FragmentKnowledgeGuid = new Guid("0ABE4F35-01DF-4134-A24A-7C79F0859844");

        /// <summary>
        /// A GUID value specifies a content tag knowledge.
        /// </summary>
        public static readonly Guid ContentTagKnowledgeGuid = new Guid("10091F13-C882-40FB-9886-6533F934C21D");

        /// <summary>
        /// A GUID value specifies a Version Token knowledge.
        /// </summary>
        public static readonly Guid VersionTokenKnowledgeGuid = new Guid("BF12E2C1-E64F-4959-8282-73B9A24A7C44");

        /// <summary>
        /// A mapping that maps the knowledge GUID value and the knowledge types.
        /// </summary
        public static readonly Dictionary<Guid, Type> KnowledgeTypeGuidMapping = new Dictionary<Guid, Type>
        {
            { CellKnowledgeGuid, typeof(CellKnowledge) },
            { WaterlineKnowledgeGuid, typeof(WaterlineKnowledge) },
            { FragmentKnowledgeGuid, typeof(FragmentKnowledge) },
            { ContentTagKnowledgeGuid, typeof(ContentTagKnowledge) },
            { VersionTokenKnowledgeGuid,typeof(VersionTokenKnowledge)},
        };

        /// <summary>
        /// A mapping that maps the stream object header and knowledge GUID value.
        /// </summary>
        public static readonly Dictionary<StreamObjectTypeHeaderStart, Guid> KnowledgeEnumGuidMapping = new Dictionary<StreamObjectTypeHeaderStart, Guid>
        {
            { StreamObjectTypeHeaderStart.CellKnowledge, CellKnowledgeGuid },
            { StreamObjectTypeHeaderStart.WaterlineKnowledge, WaterlineKnowledgeGuid },
            { StreamObjectTypeHeaderStart.FragmentKnowledge, FragmentKnowledgeGuid },
            { StreamObjectTypeHeaderStart.ContentTagKnowledge, ContentTagKnowledgeGuid },
            { StreamObjectTypeHeaderStart.VersionTokenKnowledge, VersionTokenKnowledgeGuid},
        };

        /// <summary>
        /// A specialized knowledge data.
        /// </summary>
        private SpecializedKnowledgeData specializedKnowledgeData;

        /// <summary>
        /// Initializes a new instance of the SpecializedKnowledge class.
        /// </summary>
        public SpecializedKnowledge()
            : base(StreamObjectTypeHeaderStart.SpecializedKnowledge)
        {
            this.GUID = Guid.Empty;
        }

        /// <summary>
        /// Gets or sets a GUID that specifies the type of specialized knowledge. 
        /// </summary>
        public Guid GUID { get; set; }

        /// <summary>
        /// Gets a specialized knowledge data.
        /// </summary>
        public SpecializedKnowledgeData SpecializedKnowledgeData
        {
            get { return this.specializedKnowledgeData; }
        }

        /// <summary>
        /// This method is used to deserialize the items of the specialized knowledge from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the specialized knowledge.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 16)
            {
                throw new KnowledgeParseErrorException(currentIndex, "Failed to check the stream object header length in the DeserializeSpecializedKnowledgeFromByteArray", null);
            }

            int index = currentIndex;
            byte[] temp = new byte[16];
            Array.Copy(byteArray, index, temp, 0, 16);
            this.GUID = new Guid(temp);
            if (!KnowledgeTypeGuidMapping.Keys.Contains(this.GUID))
            {
                throw new KnowledgeParseErrorException(index, string.Format("Failed to check the special knowledge guid value in the DeserializeSpecializedKnowledgeFromByteArray, the value {0 }is not defined", this.GUID), null);
            }

            index += 16;
            this.specializedKnowledgeData = Activator.CreateInstance(KnowledgeTypeGuidMapping[this.GUID]) as SpecializedKnowledgeData;

            StreamObjectHeaderStart specializedKnowledgeDataHeader;
            int headerLength = 0;
            if ((headerLength = StreamObjectHeaderStart.TryParse(byteArray, index, out specializedKnowledgeDataHeader)) == 0)
            {
                throw new KnowledgeParseErrorException(index, "Failed to parse the specialized knowledge data stream object header", null);
            }

            if (!KnowledgeEnumGuidMapping.Keys.Contains(specializedKnowledgeDataHeader.Type))
            {
                throw new KnowledgeParseErrorException(index, "Unexpected specialized knowledge data stream object header type " + specializedKnowledgeDataHeader.Type, null);
            }

            if (KnowledgeEnumGuidMapping[specializedKnowledgeDataHeader.Type] != this.GUID)
            {
                throw new KnowledgeParseErrorException(index, "Unmatched specialized knowledge data stream object header type and the specified guid value", null);
            }

            index += headerLength;
            try
            {
                index += this.specializedKnowledgeData.DeserializeFromByteArray(specializedKnowledgeDataHeader, byteArray, index);
            }
            catch (StreamObjectParseErrorException streamObjectException)
            {
                throw new KnowledgeParseErrorException(index, streamObjectException);
            }
            catch (BasicObjectParseErrorException basicObjectException)
            {
                throw new KnowledgeParseErrorException(index, basicObjectException);
            }

            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the specialized knowledge to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of the specialized knowledge.</param>
        /// <returns>Return the length in byte of the items in the specialized knowledge.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            byteList.AddRange(this.GUID.ToByteArray());
            byteList.AddRange(this.specializedKnowledgeData.SerializeToByteList());
            return 16;
        }
    }
}