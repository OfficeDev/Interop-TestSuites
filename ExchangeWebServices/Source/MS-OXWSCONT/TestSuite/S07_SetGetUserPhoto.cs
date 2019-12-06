namespace Microsoft.Protocols.TestSuites.MS_OXWSCONT
{
    using System.IO;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to set and get user photo in server.
    /// </summary>
    [TestClass]
    public class S07_SetGetUserPhoto : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="context">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            TestClassBase.Initialize(context);
        }

        /// <summary>
        /// Clean up the test class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        /// <summary>
        /// This test case is intended to validate the successful response of SetUserPhoto operation and GetUserPhoto operation.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S07_TC01_SetUserPhotoSuccess()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1275114, this.Site), "Implementation does support the SetUserPhoto operation.");

            #region Step 1: Call SetUserPhoto operation to set a photo to specific user.
            string emailAddress = string.Format("{0}@{1}", Common.GetConfigurationPropertyValue("ContactUserName", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            SetUserPhotoType setUserPhotoRequest = new SetUserPhotoType();
            setUserPhotoRequest.Email = emailAddress;

            using (FileStream imageStream = new FileStream("UserPhoto.jpg", FileMode.Open, FileAccess.ReadWrite))
            {
                byte[] buffer = new byte[imageStream.Length];

                imageStream.Read(buffer, 0, (int)imageStream.Length);

                string imagContent = System.Convert.ToBase64String(buffer);
                setUserPhotoRequest.Content = imagContent;
            }

            SetUserPhotoResponseMessageType setUserPhotoResponse = this.CONTAdapter.SetUserPhoto(setUserPhotoRequest);

            Site.Assert.IsNotNull(setUserPhotoResponse, "SetUserPhoto operation success.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302081");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302081
            this.Site.CaptureRequirementIfIsNotNull(
                setUserPhotoResponse,
                302081,
                @"[In SetUserPhoto] [The protocol client sends a SetUserPhotoSoapIn request WSDL message] and the protocol server responds with a SetUserPhotoSoapOut response WSDL message");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275114");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275114
            this.Site.CaptureRequirementIfIsNotNull(
                setUserPhotoResponse,
                1275114,
                @"[In Appendix C: Product Behavior] Implementation does support the SetUserPhoto operation. (Exchange 2016 and above follow this behavior.)");
            #endregion

            #region Step 2: Call GetUserPhoto operation to get the photo which is set by step above.
            GetUserPhotoType getUserPhotoRequest = new GetUserPhotoType();
            getUserPhotoRequest.Email = emailAddress;
            getUserPhotoRequest.SizeRequested = UserPhotoSizeType.HR96x96;

            GetUserPhotoResponseMessageType getUserPhotoResponse = this.CONTAdapter.GetUserPhoto(getUserPhotoRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302024");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302024
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getUserPhotoResponse.ResponseClass,
                302024,
                @"[In GetUserPhotoSoapOut] A successful GetUserPhoto WSDL operation request returns a GetUserPhotoResponse element with the ResponseClass attribute set to ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302025");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302025
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                getUserPhotoResponse.ResponseCode,
                302025,
                @"[In GetUserPhotoSoapOut] [A successful GetUserPhoto WSDL operation request returns a GetUserPhotoResponse element ] The ResponseCode element of the GetUserPhotoResponse element is set to ""No Error"".");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the UserPhotoSizeType enum value. 
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S07_TC02_GetUserPhotoSizeTypeEnumValue()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1275114, this.Site), "Implementation does not support the SetUserPhoto operation.");

            #region Step 1: Call SetUserPhoto operation to set a photo to specific user.
            string emailAddress = string.Format("{0}@{1}", Common.GetConfigurationPropertyValue("ContactUserName", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            SetUserPhotoType setUserPhotoRequest = new SetUserPhotoType();
            setUserPhotoRequest.Email = emailAddress;

            using (FileStream imageStream = new FileStream("UserPhoto.jpg", FileMode.Open, FileAccess.ReadWrite))
            {
                byte[] buffer = new byte[imageStream.Length];

                imageStream.Read(buffer, 0, (int)imageStream.Length);

                string imagContent = System.Convert.ToBase64String(buffer);
                setUserPhotoRequest.Content = imagContent;
            }

            SetUserPhotoResponseMessageType setUserPhotoResponse = this.CONTAdapter.SetUserPhoto(setUserPhotoRequest);

            Site.Assert.IsNotNull(setUserPhotoResponse, "SetUserPhoto operation success.");

            #endregion

            #region Step 2: Call GetUserPhoto operation to get a photo with specified size:HR48x48.
            GetUserPhotoType getUserPhotoRequest = new GetUserPhotoType();
            getUserPhotoRequest.Email = emailAddress;
            getUserPhotoRequest.SizeRequested = UserPhotoSizeType.HR48x48;

            GetUserPhotoResponseMessageType getUserPhotoResponse;
            getUserPhotoResponse = this.CONTAdapter.GetUserPhoto(getUserPhotoRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302068");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302068
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getUserPhotoResponse.ResponseClass,
                302068,
                @"[In UserPhotoSizeType] HR48x48: Specifies that the image is 48 pixels high and 48 pixels wide.");

            #endregion

            #region Step 3: Call GetUserPhoto operation to get a photo with specified size:HR64x64.
            getUserPhotoRequest.SizeRequested = UserPhotoSizeType.HR64x64;

            getUserPhotoResponse = this.CONTAdapter.GetUserPhoto(getUserPhotoRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302069");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302069
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getUserPhotoResponse.ResponseClass,
                302069,
                @"[In UserPhotoSizeType] HR64x64: Specifies that the image is 64 pixels high and 64 pixels wide.");

            #endregion

            #region Step 4: Call GetUserPhoto operation to get a photo with specified size:HR96x96.
            getUserPhotoRequest.SizeRequested = UserPhotoSizeType.HR96x96;

            getUserPhotoResponse = this.CONTAdapter.GetUserPhoto(getUserPhotoRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302070");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302070
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getUserPhotoResponse.ResponseClass,
                302070,
                @"[In UserPhotoSizeType] HR96x96: Specifies that the image is 96 pixels high and 96 pixels wide.");

            #endregion

            #region Step 6: Call GetUserPhoto operation to get a photo with specified size:HR120x120.
            getUserPhotoRequest.SizeRequested = UserPhotoSizeType.HR120x120;

            getUserPhotoResponse = this.CONTAdapter.GetUserPhoto(getUserPhotoRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302071");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302071
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getUserPhotoResponse.ResponseClass,
                302071,
                @"[In UserPhotoSizeType] HR120x120: Specifies that the image is 120 pixels high and 120 pixels wide.");

            #endregion

            #region Step 7: Call GetUserPhoto operation to get a photo with specified size:HR240x240.
            getUserPhotoRequest.SizeRequested = UserPhotoSizeType.HR240x240;

            getUserPhotoResponse = this.CONTAdapter.GetUserPhoto(getUserPhotoRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302072");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302072
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getUserPhotoResponse.ResponseClass,
                302072,
                @"[In UserPhotoSizeType] HR240x240: Specifies that the image is 240 pixels high and 240 pixels wide.");

            #endregion

            #region Step 8: Call GetUserPhoto operation to get a photo with specified size:HR360x360.
            getUserPhotoRequest.SizeRequested = UserPhotoSizeType.HR360x360;

            getUserPhotoResponse = this.CONTAdapter.GetUserPhoto(getUserPhotoRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302073");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302073
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getUserPhotoResponse.ResponseClass,
                302073,
                @"[In UserPhotoSizeType] HR360x360: Specifies that the image is 360 pixels high and 360 pixels wide.");

            #endregion

            #region Step 9: Call GetUserPhoto operation to get a photo with specified size:HR432x432.
            getUserPhotoRequest.SizeRequested = UserPhotoSizeType.HR432x432;

            getUserPhotoResponse = this.CONTAdapter.GetUserPhoto(getUserPhotoRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302074");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302074
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getUserPhotoResponse.ResponseClass,
                302074,
                @"[In UserPhotoSizeType] HR432x432: Specifies that the image is 432 pixels high and 432 pixels wide.");

            #endregion

            #region Step 10: Call GetUserPhoto operation to get a photo with specified size:HR504x504.
            getUserPhotoRequest.SizeRequested = UserPhotoSizeType.HR504x504;

            getUserPhotoResponse = this.CONTAdapter.GetUserPhoto(getUserPhotoRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302075");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302075
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getUserPhotoResponse.ResponseClass,
                302075,
                @"[In UserPhotoSizeType] HR504x504: Specifies that the image is 504 pixels high and 504 pixels wide.");

            #endregion

            #region Step 11: Call GetUserPhoto operation to get a photo with specified size:HR648x648.
            getUserPhotoRequest.SizeRequested = UserPhotoSizeType.HR648x648;

            getUserPhotoResponse = this.CONTAdapter.GetUserPhoto(getUserPhotoRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302076");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302076
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getUserPhotoResponse.ResponseClass,
                302076,
                @"[In UserPhotoSizeType] HR648x648: Specifies that the image is 648 pixels high and 648 pixels wide.");

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the whether the photo has changed.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S07_TC03_ChangeUserPhoto()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1275114, this.Site), "Implementation does not support the SetUserPhoto operation.");

            #region Step 1: Call SetUserPhoto operation to set a photo to specific user.
            string emailAddress = string.Format("{0}@{1}", Common.GetConfigurationPropertyValue("ContactUserName", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            SetUserPhotoType setUserPhotoRequest = new SetUserPhotoType();
            setUserPhotoRequest.Email = emailAddress;

            using (FileStream imageStream = new FileStream("UserPhoto.jpg", FileMode.Open, FileAccess.ReadWrite))
            {
                byte[] buffer = new byte[imageStream.Length];

                imageStream.Read(buffer, 0, (int)imageStream.Length);

                string imagContent = System.Convert.ToBase64String(buffer);
                setUserPhotoRequest.Content = imagContent;
            }

            SetUserPhotoResponseMessageType setUserPhotoResponse = this.CONTAdapter.SetUserPhoto(setUserPhotoRequest);

            Site.Assert.IsNotNull(setUserPhotoResponse, "SetUserPhoto operation success.");
            #endregion

            #region Step 2: Call GetUserPhoto operation to get the photo which is set by step above.
            GetUserPhotoType getUserPhotoRequest = new GetUserPhotoType();
            getUserPhotoRequest.Email = emailAddress;
            getUserPhotoRequest.SizeRequested = UserPhotoSizeType.HR96x96;

            GetUserPhotoResponseMessageType getUserPhotoResponse;
            getUserPhotoResponse = this.CONTAdapter.GetUserPhoto(getUserPhotoRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R30205601");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R30205601
            this.Site.CaptureRequirementIfIsTrue(
                getUserPhotoResponse.HasChanged,
                30205601,
                @"[In GetUserPhotoResponseMessageType] HasChanged element: If the value is true, the photo has changed.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302058");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302058
            this.Site.CaptureRequirementIfIsNotNull(
                getUserPhotoResponse.PictureData,
                302058,
                @"[In GetUserPhotoResponseMessageType] PictureData element: Specifies the binary data for the picture.");

            #endregion
        }
    }
}
