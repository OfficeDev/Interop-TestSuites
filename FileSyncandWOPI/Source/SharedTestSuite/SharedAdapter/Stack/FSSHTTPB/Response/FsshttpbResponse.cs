namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;

    /// <summary>
    /// A response message specifies the format used to contain sub-responses matching sub-requests. 
    /// </summary>
    public class FsshttpbResponse
    {
        /// <summary>
        /// Initializes a new instance of the FsshttpbResponse class.
        /// </summary>
        public FsshttpbResponse()
        {
            this.CellSubResponses = new List<FsshttpbSubResponse>();
        }

        /// <summary>
        /// Gets or sets Protocol Version (2bytes): An unsigned integer that specifies the protocol schema version number used in this request.This value MUST be 12.
        /// </summary>
        public ushort ProtocolVersion { get; set; }

        /// <summary>
        /// Gets or sets Minimum Version (2 bytes): An unsigned integer that specifies the oldest version of the protocol schema that this schema is compatible with. This value MUST be 11.
        /// </summary>
        public ushort MinimumVersion { get; set; }

        /// <summary>
        /// Gets or sets Signature (8 bytes): An unsigned integer that specifies a constant signature, to identify this as a response. This MUST be set to 0x9B069439F329CF9D.
        /// </summary>
        public ulong Signature { get; set; }

        /// <summary>
        /// Gets or sets Response Start (4 bytes): A 32-bit stream object header that specifies a response start.
        /// </summary>
        public StreamObjectHeaderStart32bit ResponseStart { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the request has failed and a response error MUST follow.
        /// </summary>
        public bool Status { get; set; }

        /// <summary>
        /// Gets or sets Reserved (7 bits): A 7-bit reserved and MUST be set to 0 and MUST be ignored.
        /// </summary>
        public byte Reserved { get; set; }

        /// <summary>
        /// Gets or sets Response Data (variable): A response error that specifies the error information if the request failed 
        /// </summary>
        public ResponseError ResponseError { get; set; }

        /// <summary>
        /// Gets or sets Data Element Package (variable): An optional data elements package structure that specifies data elements corresponding to the sub-responses. 
        /// If the sub-responses do not reference data elements or no data elements are available for the sub-response then this structure will not be present.
        /// </summary>
        public DataElementPackage DataElementPackage { get; set; }

        /// <summary>
        /// Gets or sets the cell sub responses.
        /// </summary>
        public List<FsshttpbSubResponse> CellSubResponses { get; set; }

        /// <summary>
        /// Gets or sets the response end.
        /// </summary>
        public StreamObjectHeaderEnd16bit ResponseEnd { get; set; }
        
        /// <summary>
        /// Deserialize response from byte array.
        /// </summary>
        /// <param name="byteArray">Server returned message.</param>
        /// <param name="startIndex">The index special where start.</param>
        /// <returns>The instance of CellResponse.</returns>
        public static FsshttpbResponse DeserializeResponseFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;

            FsshttpbResponse response = new FsshttpbResponse();

            response.ProtocolVersion = LittleEndianBitConverter.ToUInt16(byteArray, index);
            index += 2;

            response.MinimumVersion = LittleEndianBitConverter.ToUInt16(byteArray, index);
            index += 2;

            response.Signature = LittleEndianBitConverter.ToUInt64(byteArray, index);
            index += 8;

            int length = 0;
            StreamObjectHeaderStart streamObjectHeader;
            if ((length = StreamObjectHeaderStart.TryParse(byteArray, index, out streamObjectHeader)) == 0)
            {
                throw new ResponseParseErrorException(index, "Failed to parse the response header", null);
            }

            if (!(streamObjectHeader is StreamObjectHeaderStart32bit))
            {
                throw new ResponseParseErrorException(index, "Unexpected 16-bit response stream object header, expect 32-bit stream object header for Response", null);
            }

            if (streamObjectHeader.Type != StreamObjectTypeHeaderStart.FsshttpbResponse)
            {
                throw new ResponseParseErrorException(index, "Failed to extract the response header type, unexpected value " + streamObjectHeader.Type, null);
            }

            if (streamObjectHeader.Length != 1)
            {
                throw new ResponseParseErrorException(index, "Response object over-parse error", null);
            }

            index += length;
            response.ResponseStart = streamObjectHeader as StreamObjectHeaderStart32bit;

            response.Status = (byteArray[index] & 0x1) == 0x1 ? true : false;
            response.Reserved = (byte)(byteArray[index] >> 1);
            index += 1;

            try
            {
                if (response.Status)
                {
                    response.ResponseError = StreamObject.GetCurrent<ResponseError>(byteArray, ref index);
                    response.DataElementPackage = null;
                    response.CellSubResponses = null;
                }
                else
                {
                    DataElementPackage package;
                    if (StreamObject.TryGetCurrent<DataElementPackage>(byteArray, ref index, out package))
                    {
                        response.DataElementPackage = package;
                    }

                    response.CellSubResponses = new List<FsshttpbSubResponse>();
                    FsshttpbSubResponse subResponse;
                    while (StreamObject.TryGetCurrent<FsshttpbSubResponse>(byteArray, ref index, out subResponse))
                    {
                        response.CellSubResponses.Add(subResponse);
                    }
                }

                response.ResponseEnd = BasicObject.Parse<StreamObjectHeaderEnd16bit>(byteArray, ref index);
            }
            catch (StreamObjectParseErrorException streamObjectException)
            {
                throw new ResponseParseErrorException(index, streamObjectException);
            }
            catch (DataElementParseErrorException dataElementException)
            {
                throw new ResponseParseErrorException(index, dataElementException);
            }
            catch (KnowledgeParseErrorException knowledgeException)
            {
                throw new ResponseParseErrorException(index, knowledgeException);
            }

            if (index != byteArray.Length)
            {
                throw new ResponseParseErrorException(index, "Failed to pass the whole response, not reach the end of the byte array", null);
            }

            return response;
        }
    }
}