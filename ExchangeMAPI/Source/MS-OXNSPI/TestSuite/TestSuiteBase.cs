namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The base test suite class defines common initialization method and cleanup method for all the five scenarios.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : Microsoft.Protocols.TestTools.TestClassBase
    {
        #region Private variables

        /// <summary>
        /// Indicate whether the cleanup step is necessary to delete the PidTagAddressBookMember property.
        /// </summary>
        private bool isRequireToDeleteAddressBookMember;

        /// <summary>
        /// Indicate whether the cleanup step is necessary to delete the PidTagAddressBookPublicDelegates property.
        /// </summary>
        private bool isRequireToDeleteAddressBookPublicDelegate;

        /// <summary>
        /// Indicate whether the entry ID used to modify object property is Ephemeral Entry ID .
        /// </summary>
        private bool isEphemeralEntryID;

        /// <summary>
        /// A Minimal Entry ID of a specific object to be modified.
        /// </summary>
        private uint midToBeModified;

        /// <summary>
        /// A valid value of PidTagAddressBookMember or PidTagAddressBookPublicDelegates to be deleted.
        /// </summary>
        private BinaryArray_r entryIdToBeDeleted;

        /// <summary>
        /// Return value of NSPI method.
        /// </summary>
        private ErrorCodeValue result;

        /// <summary>
        /// The instance of MS-OXNSPI protocol adapter.
        /// </summary>
        private IMS_OXNSPIAdapter protocolAdatper;

        /// <summary>
        /// The instance of the SUT Control Adapter.
        /// </summary>
        private IMS_OXNSPISUTControlAdapter sutControlAdapter;

        /// <summary>
        /// The transport used by the test suite.
        /// </summary>
        private string transport;

        #endregion

        /// <summary>
        /// Gets or sets a value indicating whether the PidTagAddressBookMember property needs to be deleted.
        /// </summary>
        public bool IsRequireToDeleteAddressBookMember
        {
            get { return this.isRequireToDeleteAddressBookMember; }
            set { this.isRequireToDeleteAddressBookMember = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the PidTagAddressBookPublicDelegates property needs to be deleted.
        /// </summary>
        public bool IsRequireToDeleteAddressBookPublicDelegate
        {
            get { return this.isRequireToDeleteAddressBookPublicDelegate; }
            set { this.isRequireToDeleteAddressBookPublicDelegate = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the entry ID used to modify object property is Ephemeral Entry ID.
        /// </summary>
        public bool IsEphemeralEntryID
        {
            get { return this.isEphemeralEntryID; }
            set { this.isEphemeralEntryID = value; }
        }

        /// <summary>
        /// Gets or sets the value of midToBeModified.
        /// </summary>
        public uint MidToBeModified
        {
            get { return this.midToBeModified; }
            set { this.midToBeModified = value; }
        }

        /// <summary>
        /// Gets or sets the value of transport.
        /// </summary>
        public string Transport
        {
            get { return this.transport; }
            set { this.transport = value; }
        }

        /// <summary>
        /// Gets or sets the value of entryIdToBeDeleted.
        /// </summary>
        public BinaryArray_r EntryIdToBeDeleted
        {
            get { return this.entryIdToBeDeleted; }
            set { this.entryIdToBeDeleted = value; }
        }

        /// <summary>
        /// Gets or sets the return value of NSPI method.
        /// </summary>
        public ErrorCodeValue Result
        {
            get { return this.result; }
            set { this.result = value; }
        }

        /// <summary>
        /// Gets or sets the instance of MS-OXNSPI protocol adapter.
        /// </summary>
        public IMS_OXNSPIAdapter ProtocolAdatper
        {
            get { return this.protocolAdatper; }
            set { this.protocolAdatper = value; }
        }

        /// <summary>
        /// Gets or sets the instance of the SUT control adapter.
        /// </summary>
        public IMS_OXNSPISUTControlAdapter SutControlAdapter
        {
            get { return this.sutControlAdapter; }
            set { this.sutControlAdapter = value; }
        }

        #region Test case initialization and dispose
        /// <summary>
        ///  Initialize the test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.ProtocolAdatper = Site.GetAdapter<IMS_OXNSPIAdapter>();
            this.SutControlAdapter = this.Site.GetAdapter<IMS_OXNSPISUTControlAdapter>();
            this.IsRequireToDeleteAddressBookMember = false;
            this.IsRequireToDeleteAddressBookPublicDelegate = false;
            this.IsEphemeralEntryID = false;
            this.Transport = Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower(System.Globalization.CultureInfo.CurrentCulture);
            this.MidToBeModified = 0;
            this.EntryIdToBeDeleted = new BinaryArray_r
            {
                CValues = 1,
                Lpbin = new Binary_r[1]
            };
        }

        /// <summary>
        ///  Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            bool transportIsMAPI = Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower(CultureInfo.InvariantCulture) == "mapi_http";
            if (bool.Parse(Common.GetConfigurationPropertyValue("MS-OXNSPI_Supported", this.Site)) && (!transportIsMAPI || (transportIsMAPI && Common.IsRequirementEnabled(2003, this.Site))))
            {
                if (this.IsRequireToDeleteAddressBookMember)
                {
                    uint flagsOfModLinkAtt = (uint)NspiModLinkAtFlag.fDelete;
                    uint propTagOfModLinkAtt = (uint)AulProp.PidTagAddressBookMember;
                    ErrorCodeValue result = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, this.midToBeModified, this.entryIdToBeDeleted);
                    ErrorCodeValue expectedResult = ErrorCodeValue.Success;
                    if (!Common.IsRequirementEnabled(1340, this.Site) && this.IsEphemeralEntryID == true)
                    {
                        expectedResult = ErrorCodeValue.GeneralFailure; // The Exchange 2013 returns "GeneralFailure" when the Ephemeral Entry ID is used to modify the PidTagAddressBookMember value.
                    }

                    Site.Assert.AreEqual<ErrorCodeValue>(expectedResult, result, "NspiModLinkAtt method should return Success in Exchange 2010 and Exchange 2013.");
                }

                if (this.IsRequireToDeleteAddressBookPublicDelegate)
                {
                    uint flagsOfModLinkAtt = (uint)NspiModLinkAtFlag.fDelete;
                    uint propTagOfModLinkAtt = (uint)AulProp.PidTagAddressBookPublicDelegates;
                    ErrorCodeValue result = this.ProtocolAdatper.NspiModLinkAtt(flagsOfModLinkAtt, propTagOfModLinkAtt, this.midToBeModified, this.entryIdToBeDeleted);
                    ErrorCodeValue expectedResult = ErrorCodeValue.Success;
                    if (!Common.IsRequirementEnabled(1340, this.Site))
                    {
                        expectedResult = ErrorCodeValue.GeneralFailure; // The Exchange 2013 returns "GeneralFailure" when the Ephemeral Entry ID is used to modify the PidTagAddressBookMember value.
                    }

                    Site.Assert.AreEqual<ErrorCodeValue>(expectedResult, result, "NspiModLinkAtt method should return Success in Exchange 2010, and Exchange 2013 will return GeneralFailure.");
                }

                this.ProtocolAdatper.Reset();
            }

            base.TestCleanup();
        }
        #endregion Test case initialization and dispose

        /// <summary>
        /// Disable the test case if MAPIHTTP transport is selected but not supported by current test environment.
        /// </summary>
        protected void CheckMAPIHTTPTransportSupported()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower(CultureInfo.InvariantCulture) == "mapi_http" && !Common.IsRequirementEnabled(2003, this.Site))
            {
                Site.Assume.Inconclusive("This case can not run with MAPIHTTP transport with current SUT version. MAPIHTTP is supported from Exchange Server 2013 SP1.");
            }
        }

        /// <summary>
        /// Disable the test case if the protocol doesn't support current product version.
        /// </summary>
        protected void CheckProductSupported()
        {
            bool isSupported = bool.Parse(Common.GetConfigurationPropertyValue("MS-OXNSPI_Supported", this.Site));
            Site.Assume.IsTrue(isSupported, "The MS-OXNSPI is not supported when the MS-OXNSPI_Supported property is false in Should/May PTFconfig file.");
        }
    }
}