namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to represent the RevisionRoleDeclarationFND structure.
    /// </summary>
    public class RevisionRoleDeclarationFND : FileNodeBase
    {
        /// <summary>
        /// Gets or sets the value of rid field.
        /// </summary>
        public ExtendedGUID rid { get; set; }

        /// <summary>
        /// Gets or sets the value of RevisionRole field.
        /// </summary>
        public byte[] RevisionRole { get; set; }

        /// <summary>
        /// This method is used to deserialize the RevisionRoleDeclarationFND object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the RevisionRoleDeclarationFND object.</returns>
        public override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.rid = new ExtendedGUID();
            int len = this.rid.DoDeserializeFromByteArray(byteArray, index);
            index += len;
            this.RevisionRole = new byte[4];
            Array.Copy(byteArray, index, this.RevisionRole, 0, 4);
            index += 4;
            
            return index - startIndex;
        }

        /// <summary>
        /// This method is used to convert the element of RevisionRoleDeclarationFND object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of RevisionRoleDeclarationFND.</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.rid.SerializeToByteList());
            byteList.AddRange(this.RevisionRole);

            return byteList;
        }
    }
}
