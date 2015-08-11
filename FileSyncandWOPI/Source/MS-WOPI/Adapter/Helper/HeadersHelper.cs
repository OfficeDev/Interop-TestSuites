namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Net;
    using System.Text;

    /// <summary>
    /// This class is used to build the various kinds of headers
    /// </summary>
    public class HeadersHelper
    {
        /// <summary>
        /// Prevents a default instance of the HeadersHelper class from being created
        /// </summary>
        private HeadersHelper()
        { 
        }

        /// <summary>
        /// This method is used to get the proof header from the resource URI.
        /// </summary>
        /// <param name="targetResourceUri">The value of the target resource URI.</param>
        /// <param name="currentTicks">The current header value.</param>
        /// <param name="isuseOldPublicKey">Whether used the old public key.</param>
        /// <returns>The value of the proof header.</returns>
        public static string GetProofHeaderValue(string targetResourceUri, long currentTicks, bool isuseOldPublicKey = false)
        {
            byte[] proofHeaderBytes = ConstructProofHeaderBytes(targetResourceUri, currentTicks);

            // encrypt the bytes by using RSA Asymmetric Encrypt algorithm, the public key can be get by "RSACryptoContext.PublicKeyString"
            byte[] encryptedBytes = new byte[0];
            if (isuseOldPublicKey)
            {
                encryptedBytes = RSACryptoContext.SignDataWithOldPublicKey(proofHeaderBytes);
            }
            else
            {
                encryptedBytes = RSACryptoContext.SignDataWithCurrentPublicKey(proofHeaderBytes);
            }

            return Convert.ToBase64String(encryptedBytes);
        }

        /// <summary>
        /// This method is used to get the common headers.
        /// </summary>
        /// <param name="targetResourceUri">The value of the target resource URI.</param>
        /// <returns>The common headers.</returns>
        public static WebHeaderCollection GetCommonHeaders(string targetResourceUri)
        {
            WebHeaderCollection getCommonHeaders = new WebHeaderCollection();
            string tokenValue = TokenAndRequestUrlHelper.GetTokenValueFromWOPIResourceUrl(targetResourceUri);

            // Setting the required headers:
            string authorizationValue = string.Format(@"Bearer {0}", tokenValue);
            getCommonHeaders.Add("Authorization", authorizationValue);
            long currentTicks = DateTime.UtcNow.Ticks;
            string proofHeaderValueOfCurrent = HeadersHelper.GetProofHeaderValue(targetResourceUri, currentTicks);
            getCommonHeaders.Add("X-WOPI-Proof", proofHeaderValueOfCurrent);
            string proofHeaderValueOfOld = HeadersHelper.GetProofHeaderValue(targetResourceUri, currentTicks, true);
            getCommonHeaders.Add("X-WOPI-ProofOld", proofHeaderValueOfOld);
            getCommonHeaders.Add("X-WOPI-TimeStamp", currentTicks.ToString());

            return getCommonHeaders;
        }

        /// <summary>
        /// This method is used to build the proof header.
        /// </summary>
        /// <param name="encodeString">The value of the encode string which is used to build header.</param>
        /// <returns>The header.</returns>
        private static byte[] EncodeStringValueForProofHeader(string encodeString)
        {
            if (string.IsNullOrEmpty(encodeString))
            {
                throw new ArgumentNullException("encodeString");
            }

            return Encoding.UTF8.GetBytes(encodeString);
        }

        /// <summary>
        /// This method is used to convert the integer value to Proof Header type.
        /// </summary>
        /// <param name="numericValue">The value which will be transformed.</param>
        /// <returns>The byte proof header collections.</returns>
        private static byte[] EncodeNumericValueForProofHeader(int numericValue)
        {
            if (numericValue <= 0)
            {
                throw new ArgumentException("The [numericValue] parameter must larger than 0.");
            }

            byte[] numericBytes = BitConverter.GetBytes(numericValue);
            return ConvertNetWorkOrderBytes(numericBytes);
        }

        /// <summary>
        /// This method is used to convert the long value to Proof Header type.
        /// </summary>
        /// <param name="numericValue">The value which will be transformed.</param>
        /// <returns>The byte proof header collections.</returns>
        private static byte[] EncodeNumericValueForProofHeader(long numericValue)
        {
            if (numericValue <= 0)
            {
                throw new ArgumentException("The [numericValue] parameter must larger than 0.");
            }

            byte[] numericBytes = BitConverter.GetBytes(numericValue);
            return ConvertNetWorkOrderBytes(numericBytes);
        }

        /// <summary>
        /// This method is used to reverse the sequence of the elements in the byte array.
        /// </summary>
        /// <param name="originalBytes">The byte array which will be reversed.</param>
        /// <returns>The byte array had been reversed.</returns>
        private static byte[] ConvertNetWorkOrderBytes(byte[] originalBytes)
        { 
           if (BitConverter.IsLittleEndian)
           {
               Array.Reverse(originalBytes);
           }

           return originalBytes;
        }

        /// <summary>
        /// This method is used to construct an byte array type for Proof Header.
        /// </summary>
        /// <param name="targetResourceUrl">The value of the target resource URL.</param>
        /// <param name="ticksOfTimeStamp">The current header value.</param>
        /// <returns>An byte array of the proof header.</returns>
        private static byte[] ConstructProofHeaderBytes(string targetResourceUrl, long ticksOfTimeStamp)
        {
            if (string.IsNullOrEmpty(targetResourceUrl))
            {
                throw new ArgumentNullException("targetResourceUrl");
            }

            if (0 == ticksOfTimeStamp)
            {
                throw new ArgumentException("The [ticksOfTimeStamp] parameter must larger than 0.");
            }

            // Encode Token
            string tokenValue = TokenAndRequestUrlHelper.GetTokenValueFromWOPIResourceUrl(targetResourceUrl);
            byte[] encodedTokenLength = EncodeNumericValueForProofHeader(tokenValue.Length);
            byte[] encodedTokenContent = EncodeStringValueForProofHeader(tokenValue);

            // Encode Url content
            targetResourceUrl = targetResourceUrl.ToUpper(CultureInfo.CurrentCulture);
            byte[] encodedUrlLength = EncodeNumericValueForProofHeader(targetResourceUrl.Length);
            byte[] encodedUrlContent = EncodeStringValueForProofHeader(targetResourceUrl);

            // Encode TimeStamp, the TimeStamp is a 64-bit integer, its length is 8 bytes.
            byte[] encodedTimeStampLength = EncodeNumericValueForProofHeader(8);
            byte[] encodedTimeStampContent = EncodeNumericValueForProofHeader(ticksOfTimeStamp);

            // Construct the proof header value: [token length] + [token content] + [URL length] + [url content] + [timestamp length] + [timestamp value]
            List<byte> proofHeaderBytesValue = new List<byte>();
            proofHeaderBytesValue.AddRange(encodedTokenLength);
            proofHeaderBytesValue.AddRange(encodedTokenContent);
            proofHeaderBytesValue.AddRange(encodedUrlLength);
            proofHeaderBytesValue.AddRange(encodedUrlContent);
            proofHeaderBytesValue.AddRange(encodedTimeStampLength);
            proofHeaderBytesValue.AddRange(encodedTimeStampContent);

            return proofHeaderBytesValue.ToArray();
        }
    }
}