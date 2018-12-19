namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Response Error
    /// </summary>
    public class ResponseError : StreamObject
    {
        /// <summary>
        /// GUID of cell error.
        /// </summary>
        public const string CellErrorGuid = "5A66A756-87CE-4290-A38B-C61C5BA05A67";

        /// <summary>
        /// GUID of protocol error.
        /// </summary>
        public const string ProtocolErrorGuid = "7AFEAEBF-033D-4828-9C31-3977AFE58249";

        /// <summary>
        /// GUID of win32 error.
        /// </summary>
        public const string Win32ErrorGuid = "32C39011-6E39-46C4-AB78-DB41929D679E";

        /// <summary>
        /// GUID of HRESULT error.
        /// </summary>
        public const string HresultErrorGuid = "8454C8F2-E401-405A-A198-A10B6991B56E";

        /// <summary>
        /// Initializes a new instance of the ResponseError class. 
        /// </summary>
        public ResponseError()
            : base(StreamObjectTypeHeaderStart.ResponseError)
        {
        }

        /// <summary>
        /// Gets or sets Error Type GUID.
        /// </summary>
        public Guid ErrorTypeGUID { get; set; }

        /// <summary>
        /// Gets or sets Error Data.
        /// </summary>
        public ErrorData ErrorData { get; set; }

        /// <summary>
        /// Gets or sets Chained Error.
        /// </summary>
        public ResponseError ChainedError { get; set; }

        /// <summary>
        ///  Gets or sets the error string supplemental info.
        /// </summary>
        public ErrorStringSupplementalInfo ErrorStringSupplementalInfo { get; set; }

        /// <summary>
        /// Used to get error data.
        /// </summary>
        /// <typeparam name="T">Type of error.</typeparam>
        /// <returns>Return the error data.</returns>
        public T GetErrorData<T>()
            where T : class
        {
            if (this.ErrorData is T)
            {
                return this.ErrorData as T;
            }

            throw new InvalidOperationException(string.Format("Unable to cast DataElementData to the type {0}, its actual type is {1}", typeof(T).Name, this.ErrorData.GetType().Name));
        }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;

            if (lengthOfItems != 16)
            {
                throw new StreamObjectParseErrorException(currentIndex, "ResponseError", "Stream object over-parse error", null);
            }

            byte[] guidarray = new byte[16];
            Array.Copy(byteArray, index, guidarray, 0, 16);
            this.ErrorTypeGUID = new Guid(guidarray);
            index += 16;

            switch (this.ErrorTypeGUID.ToString().ToUpper(CultureInfo.CurrentCulture))
            {
                case CellErrorGuid:
                    this.ErrorData = StreamObject.GetCurrent<CellError>(byteArray, ref index);
                    break;

                case ProtocolErrorGuid:
                    this.ErrorData = StreamObject.GetCurrent<ProtocolError>(byteArray, ref index);
                    break;

                case Win32ErrorGuid:
                    this.ErrorData = StreamObject.GetCurrent<Win32Error>(byteArray, ref index);
                    break;

                case HresultErrorGuid:
                    this.ErrorData = StreamObject.GetCurrent<HRESULTError>(byteArray, ref index);
                    break;

                default:
                    throw new StreamObjectParseErrorException(index - 16, "ResponseError", "Failed to extract the error Guid value, the value" + this.ErrorTypeGUID + "is not defined", null);
            }

            ErrorStringSupplementalInfo errorInfo;
            if (StreamObject.TryGetCurrent<ErrorStringSupplementalInfo>(byteArray, ref index, out errorInfo))
            {
                this.ErrorStringSupplementalInfo = errorInfo;
            }

            ResponseError chainedError;
            if (StreamObject.TryGetCurrent<ResponseError>(byteArray, ref index, out chainedError))
            {
                this.ChainedError = chainedError;
            }

            currentIndex = index;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            throw new NotImplementedException("The method ResponseError::SerializeItemsToByteList does not implement in the current stage.");
        }
    }

    /// <summary>
    /// Error string supplemental info.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class ErrorStringSupplementalInfo : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ErrorStringSupplementalInfo class. 
        /// </summary>
        public ErrorStringSupplementalInfo()
            : base(StreamObjectTypeHeaderStart.ErrorStringSupplementalInfo)
        {
        }

        /// <summary>
        /// Gets or sets a string item (section 2.2.1.4) that specifies the supplemental information of the error string for the error string supplemental info start.
        /// </summary>
        public StringItem ErrorString { get; set; }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            throw new NotImplementedException("The method CellError::SerializeItemsToByteList does not implement in the current stage.");
        }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains error message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.ErrorString = BasicObject.Parse<StringItem>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(index, "ErrorStringSupplementalInfo", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }
    }

    /// <summary>
    /// Cell Error
    /// </summary>
    public class CellError : ErrorData
    {
        /// <summary>
        /// Initializes a new instance of the CellError class. 
        /// </summary>
        public CellError() :
            base(StreamObjectTypeHeaderStart.CellError)
        {
        }

        /// <summary>
        /// Gets the Error detail information.
        /// </summary>
        public override string ErrorDetail
        {
            get
            {
                return this.ErrorCode.ToString();
            }
        }

        /// <summary>
        /// Gets or sets error code.
        /// </summary>
        public CellErrorCode ErrorCode { get; set; }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 4)
            {
                throw new StreamObjectParseErrorException(currentIndex, "CellError", "Stream object over-parse error", null);
            }

            this.ErrorCode = (CellErrorCode)LittleEndianBitConverter.ToInt32(byteArray, currentIndex);

            if (!Enum.IsDefined(typeof(CellErrorCode), this.ErrorCode))
            {
                throw new StreamObjectParseErrorException(currentIndex, "CellError", "Unexpected error code value " + this.ErrorCode, null);
            }

            currentIndex += 4;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(System.Collections.Generic.List<byte> byteList)
        {
            throw new NotImplementedException("The method CellError::SerializeItemsToByteList does not implement in the current stage.");
        }
    }

    /// <summary>
    /// Protocol Error
    /// </summary>
    public class ProtocolError : ErrorData
    {
        /// <summary>
        /// Initializes a new instance of the ProtocolError class. 
        /// </summary>
        public ProtocolError()
            : base(StreamObjectTypeHeaderStart.ProtocolError)
        {
        }

        /// <summary>
        /// Gets or sets error code.
        /// </summary>
        public ProtocolErrorCode ErrorCode { get; set; }

        /// <summary>
        /// Gets the Error detail information.
        /// </summary>
        public override string ErrorDetail
        {
            get
            {
                return this.ErrorCode.ToString();
            }
        }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 4)
            {
                throw new StreamObjectParseErrorException(currentIndex, "ProtocolError", "Stream object over-parse error", null);
            }

            this.ErrorCode = (ProtocolErrorCode)LittleEndianBitConverter.ToInt32(byteArray, currentIndex);

            if (!Enum.IsDefined(typeof(ProtocolErrorCode), this.ErrorCode))
            {
                throw new StreamObjectParseErrorException(currentIndex, "ProtocolError", "Unexpected error code value " + this.ErrorCode, null);
            }

            currentIndex += 4;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(System.Collections.Generic.List<byte> byteList)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Win32 Error
    /// </summary>
    public class Win32Error : ErrorData
    {
        /// <summary>
        /// Initializes a new instance of the Win32Error class. 
        /// </summary>
        public Win32Error()
            : base(StreamObjectTypeHeaderStart.Win32Error)
        {
        }

        /// <summary>
        /// Gets or sets error code.
        /// </summary>
        public int ErrorCode { get; set; }

        /// <summary>
        /// Gets the Error detail information.
        /// </summary>
        public override string ErrorDetail
        {
            get
            {
                return string.Format("Win32 Error: {0}", this.ErrorCode.ToString());
            }
        }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 4)
            {
                throw new StreamObjectParseErrorException(currentIndex, "Win32Error", "Stream object over-parse error", null);
            }

            this.ErrorCode = LittleEndianBitConverter.ToInt32(byteArray, currentIndex);

            currentIndex += 4;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(System.Collections.Generic.List<byte> byteList)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// HRESULT Error
    /// </summary>
    public class HRESULTError : ErrorData
    {
        /// <summary>
        /// Initializes a new instance of the HRESULTError class. 
        /// </summary>
        public HRESULTError()
            : base(StreamObjectTypeHeaderStart.HRESULTError)
        {
        }

        /// <summary>
        /// Gets or sets error code.
        /// </summary>
        public int ErrorCode { get; set; }

        /// <summary>
        /// Gets the Error detail information.
        /// </summary>
        public override string ErrorDetail
        {
            get
            {
                return this.ErrorCode.ToString();
            }
        }

        /// <summary>
        /// Deserialize items from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains response message.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        /// <param name="lengthOfItems">The length of items.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 4)
            {
                throw new StreamObjectParseErrorException(currentIndex, "HRESULTError", "Stream object over-parse error", null);
            }

            this.ErrorCode = LittleEndianBitConverter.ToInt32(byteArray, currentIndex);

            currentIndex += 4;
        }

        /// <summary>
        /// Serialize items to byte list.
        /// </summary>
        /// <param name="byteList">The byte list need to serialized.</param>
        /// <returns>The length in bytes for additional data if the current stream object has, otherwise return 0.</returns>
        protected override int SerializeItemsToByteList(System.Collections.Generic.List<byte> byteList)
        {
            throw new NotImplementedException();
        }
    }
}