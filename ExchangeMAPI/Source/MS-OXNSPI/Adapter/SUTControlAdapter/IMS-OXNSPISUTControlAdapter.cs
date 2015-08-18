namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT control adapter interface which is used in the test suite to carry out various operations related to SUT settings.
    /// </summary>
    public interface IMS_OXNSPISUTControlAdapter : IAdapter
    {
        /// <summary>
        /// A method used to get the number of Address Book objects contained in the default Global Address List.
        /// </summary>
        /// <returns>The number of Address Book objects contained in the default Global Address List.</returns>
        [MethodHelp("Enter the number of Address Book objects, which have a valid mailbox alias and are not hidden on the Global Address List. "
            + "For example, on Windows Server 2008 R2, these objects can be found in the Active Directory Service Interfaces Editor (ADSI Edit). "
            + "1. Open ADSI Edit, select \"Action\" and click \"Connect to...\" to connect to \"Default naming context\". "
            + "2. Expand to \"ADSI Edi\\Default naming context\\DC=Domain\". "
            + "3. Count all the objects that have properties with "
            + "\"mailNickName\" exists and has real value (not equal to null), and "
            + "\"msExchHideFromAddressLists\" exists and is not set to true.")]
        uint GetNumberOfAddressBookObject();
    }
}