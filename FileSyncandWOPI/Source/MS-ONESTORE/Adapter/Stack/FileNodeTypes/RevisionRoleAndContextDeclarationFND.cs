namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the RevisionRoleAndContextDeclarationFND structure.
    /// </summary>
    public class RevisionRoleAndContextDeclarationFND : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of base field.
        /// </summary>
        public RevisionRoleDeclarationFND Base { get; set; }
        /// <summary>
        /// Gets or sets the value of gctxid field.
        /// </summary>
        public ExtendedGUID gctxid { get; set; }

        /// <summary>
        /// This method is used to deserialize the RevisionRoleAndContextDeclarationFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the RevisionRoleAndContextDeclarationFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.Base = new RevisionRoleDeclarationFND();
            int len = this.Base.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.gctxid = new ExtendedGUID();
            len = this.gctxid.DoDeserializeFromByteArray(byteArray, index);
            index += len;

            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of RevisionRoleDeclarationFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of RevisionRoleDeclarationFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.Base.SerializeToByteList());
            byteList.AddRange(this.gctxid.SerializeToByteList());

            return byteList;
        }
    }
}
