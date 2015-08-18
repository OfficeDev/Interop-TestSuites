namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// The ContentsSync element contains the result of the contents synchronization download operation.
    /// contentsSync         = [progressTotal]
    ///                 *( [ProgressPerMessage] messageChange )
    ///                 [deletions]
    ///                 [readStateChanges]
    ///                 state
    ///                 IncrSyncEnd
    /// </summary>
    public class ContentsSync : SyntacticalBase
    {
        /// <summary>
        /// The end marker of the contentsSync.
        /// </summary>
        public const Markers EndMarker = Markers.PidTagIncrSyncEnd;

        /// <summary>
        /// A progressTotal value contains the result of the contents synchronization download operation
        /// </summary>
        private ProgressTotal progressTotal;

        /// <summary>
        /// A readStateChanges value.
        /// </summary>
        private ReadStateChanges readStateChanges;

        /// <summary>
        /// A deletions value contains information about IDs of messaging objects 
        /// that had been deleted, expired, or moved out of the 
        /// synchronization scope since the last synchronization,
        /// as specified in the initial ICS state.
        /// </summary>
        private Deletions deletions;

        /// <summary>
        /// A state value contains the final ICS state of the synchronization download operation.
        /// </summary>
        private State state;

        /// <summary>
        /// A list of ProgressPerMessage and messageChange tuple.
        /// </summary>
        private List<Tuple<ProgressPerMessage, MessageChange>>
            messageChangeTuples;

        /// <summary>
        /// Initializes a new instance of the ContentsSync class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public ContentsSync(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the progressTotal.
        /// </summary>
        public ProgressTotal ProgressTotal
        {
            get { return this.progressTotal; }
            set { this.progressTotal = value; }
        }

        /// <summary>
        /// Gets or sets the readStateChanges.
        /// </summary>
        public ReadStateChanges ReadStateChanges
        {
            get { return this.readStateChanges; }
            set { this.readStateChanges = value; }
        }

        /// <summary>
        /// Gets or sets the deletions.
        /// </summary>
        public Deletions Deletions
        {
            get { return this.deletions; }
            set { this.deletions = value; }
        }

        /// <summary>
        /// Gets or sets the state.
        /// </summary>
        public State State
        {
            get { return this.state; }
            set { this.state = value; }
        }

        /// <summary>
        /// Gets or sets the messageChangeTuples.
        /// </summary>
        public List<Tuple<ProgressPerMessage, MessageChange>> MessageChangeTuples
        {
            get { return this.messageChangeTuples; }
            set { this.messageChangeTuples = value; }
        }

        /// <summary>
        /// Gets a value indicating whether there is Progress Information.
        /// </summary>
        public bool HasProgressInformation
        {
            get
            {
                if (this.progressTotal != null)
                {
                    return true;
                }
                else if (this.MessageChangeTuples != null && this.MessageChangeTuples.Count > 0)
                {
                    foreach (Tuple<ProgressPerMessage, MessageChange> t in this.MessageChangeTuples)
                    {
                        if (t.Item1 != null)
                        {
                            return true;
                        }
                    }
                }

                return false;
            }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized contentsSync.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized contentsSync, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return (ProgressTotal.Verify(stream)
                || ProgressPerMessage.Verify(stream)
                || MessageChange.Verify(stream)
                || Deletions.Verify(stream)
                || ReadStateChanges.Verify(stream)
                || State.Verify(stream))
                && stream.VerifyMarker(
EndMarker,
                     (int)stream.Length - MarkersHelper.PidTagLength - (int)stream.Position);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            this.messageChangeTuples = new List<Tuple<ProgressPerMessage, MessageChange>>();
            if (ProgressTotal.Verify(stream))
            {
                this.progressTotal = new ProgressTotal(stream);
            }

            while (ProgressPerMessage.Verify(stream)
                || MessageChange.Verify(stream))
            {
                ProgressPerMessage tmp1 = null;
                MessageChange tmp2 = null;
                if (ProgressPerMessage.Verify(stream))
                {
                    tmp1 = new ProgressPerMessage(stream);
                }

                tmp2 = MessageChange.DeserializeFrom(stream) as MessageChange;
                this.messageChangeTuples.Add(
                    new Tuple<ProgressPerMessage, MessageChange>(
                        tmp1, tmp2));
            }

            if (Deletions.Verify(stream))
            {
                this.deletions = new Deletions(stream);
            }

            if (ReadStateChanges.Verify(stream))
            {
                this.readStateChanges = new ReadStateChanges(stream);
            }

            this.state = new State(stream);
            if (!stream.ReadMarker(Markers.PidTagIncrSyncEnd))
            {
                AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
            }
        }
    }
}