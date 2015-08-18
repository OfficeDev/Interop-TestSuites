namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System.Collections.Generic;

    /// <summary>
    /// The folderMessages element contains the messages contained in a folder.
    /// folderMessages       = *2( PidTagFXDelProp MessageList )
    /// </summary>
    public class FolderMessages : SyntacticalBase
    {
        /// <summary>
        /// A list of message list.
        /// </summary>
        private List<MessageList> messageLists;

        /// <summary>
        /// A list of uint32 values which are with the back of PidTagFXDelProp.
        /// Represents a directive to a client to delete specific subObjects of the object in context. 
        /// </summary>
        private List<uint> fxdelPropList;

        /// <summary>
        /// A list of FXDelPropList and MessageList tuple.
        /// </summary>
        private List<Tuple<List<uint>, MessageList>> messageTupleList;

        /// <summary>
        /// Initializes a new instance of the FolderMessages class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public FolderMessages(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets FXDelPropList.
        /// </summary>
        public List<uint> FXDelPropList
        {
            get { return this.fxdelPropList; }
            set { this.fxdelPropList = value; }
        }

        /// <summary>
        /// Gets or sets the list of FXDelPropList and MessageList tuple.
        /// </summary>
        public List<Tuple<List<uint>, MessageList>> MessageTupleList
        {
            get { return this.messageTupleList; }
            set { this.messageTupleList = value; }
        }

        /// <summary>
        /// Gets or sets messageLists
        /// </summary>
        public List<MessageList> MessageLists
        {
            get { return this.messageLists; }
            set { this.messageLists = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized folderMessages
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized folderMessages, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return !stream.IsEndOfStream
                || stream.VerifyMetaProperty(MetaProperties.PidTagFXDelProp)
                || MessageList.Verify(stream);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            int count = 0;
            long lastPosi = stream.Position;
            this.fxdelPropList = new List<uint>();
            this.messageLists = new List<MessageList>();
            this.messageTupleList = new List<Tuple<List<uint>, MessageList>>();
            while (!stream.IsEndOfStream 
                && count < 2)
            {
                lastPosi = stream.Position;
                List<uint> metaProps = new List<uint>();
                while (stream.VerifyMetaProperty(MetaProperties.PidTagFXDelProp))
                {
                    stream.ReadMarker();
                    uint delProp = stream.ReadUInt32();
                    metaProps.Add(delProp);
                }

                if (MessageList.Verify(stream))
                {
                    MessageList msgList = new MessageList(stream);

                    this.messageLists.Add(msgList);
                    this.fxdelPropList.AddRange(metaProps);
                    this.messageTupleList.Add(new Tuple<List<uint>, MessageList>(
                    metaProps, msgList));
                }
                else
                {
                    stream.Position = lastPosi;
                    break;
                }

                count++;
            }
        }
    }
}