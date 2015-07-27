//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_CPSWS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-CPSWS adapter.
    /// </summary>
    public interface IMS_CPSWSAdapter : IAdapter
    {
        /// <summary>
        /// A method used to get the claim types.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names which are used to retrieve the claim types.</param>
        /// <returns>A return value represents a list of claim types.</returns>
        ArrayOfString ClaimTypes(ArrayOfString providerNames);

        /// <summary>
        /// A method used to get the claim value types.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names which are used to retrieve the claim value types.</param>
        /// <returns>A return value represents a list of claim value types.</returns>
        ArrayOfString ClaimValueTypes(ArrayOfString providerNames);

        /// <summary>
        /// A method used to get the entity types.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names which are used to retrieve the entity types.</param>
        /// <returns>A return value represents a list of entity types.</returns>
        ArrayOfString EntityTypes(ArrayOfString providerNames);

        /// <summary>
        /// A method used to retrieve a claims provider hierarchy tree from a claims provider.
        /// </summary>
        /// <param name="providerName">A parameter represents a provider name.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the output claims provider hierarchy tree.</param>
        /// <param name="hierarchyNodeID">A parameter represents the identifier of the node to be used as root of the returned claims provider hierarchy tree.</param>
        /// <param name="numberOfLevels">A parameter represents the maximum number of levels that can be returned by the protocol server in any of the output claims provider hierarchy tree.</param>
        /// <returns>A return value represents a claims provider hierarchy tree.</returns>
        SPProviderHierarchyTree GetHierarchy(string providerName, SPPrincipalType principalType, string hierarchyNodeID, int numberOfLevels);

        /// <summary>
        /// A method used to retrieve a list of claims provider hierarchy trees from a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the output claims provider hierarchy tree.</param>
        /// <param name="numberOfLevels">A parameter represents the maximum number of levels that can be returned by the protocol server in any of the output claims provider hierarchy tree.</param>
        /// <returns>A return value represents a list of claims provider hierarchy trees.</returns>
        SPProviderHierarchyTree[] GetHierarchyAll(ArrayOfString providerNames, SPPrincipalType principalType, int numberOfLevels);

        /// <summary>
        /// A method used to retrieve schema for the current hierarchy provider.
        /// </summary>
        /// <returns>A return value represents the hierarchy claims provider schema.</returns>
        SPProviderSchema HierarchyProviderSchema();

        /// <summary>
        /// A method used to retrieve a list of claims provider schemas from a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <returns>A return value represents a list of claims provider schemas</returns>
        SPProviderSchema[] ProviderSchemas(ArrayOfString providerNames);

        /// <summary>
        /// A method used to resolve an input string to picker entities using a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the result.</param>
        /// <param name="resolveInput">A parameter represents the input to be resolved.</param>
        /// <returns>A return value represents a list of picker entities.</returns>
        PickerEntity[] Resolve(ArrayOfString providerNames, SPPrincipalType principalType, string resolveInput);

        /// <summary>
        /// A method used to resolve an input claim to picker entities using a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the result</param>
        /// <param name="resolveInput">A parameter represents the SPClaim to be resolved.</param>
        /// <returns>A return value represents a list of picker entities.</returns>
        PickerEntity[] ResolveClaim(ArrayOfString providerNames, SPPrincipalType principalType, SPClaim resolveInput);

        /// <summary>
        /// A method used to resolve a list of strings to picker entities using a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the result.</param>
        /// <param name="resolveInput">A parameter represents a list of input strings to be resolved.</param>
        /// <returns>A return value represents a list of picker entities.</returns>
        PickerEntity[] ResolveMultiple(ArrayOfString providerNames, SPPrincipalType principalType, ArrayOfString resolveInput);

        /// <summary>
        /// A method used to resolve a list of claims to picker entities using a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the result.</param>
        /// <param name="resolveInput">A parameter represents a list of claims to be resolved.</param>
        /// <returns>A return value represents a list of picker entities.</returns>
        PickerEntity[] ResolveMultipleClaim(ArrayOfString providerNames, SPPrincipalType principalType, SPClaim[] resolveInput);

        /// <summary>
        /// A method used to perform a search for entities on a list of claims providers.
        /// </summary>
        /// <param name="providerSearchArguments">A parameter represents the search arguments.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the output claims provider hierarchy tree.</param>
        /// <param name="searchPattern">A parameter represents the search string.</param>
        /// <returns>A return value represents a list of claims provider hierarchy trees.</returns>
        SPProviderHierarchyTree[] Search(SPProviderSearchArguments[] providerSearchArguments, SPPrincipalType principalType, string searchPattern);

        /// <summary>
        /// A method used to perform a search for entities on a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the output claims provider hierarchy tree.</param>
        /// <param name="searchPattern">A parameter represents the search string.</param>
        /// <param name="maxCount">A parameter represents the maximum number of matched entities to be returned in total across all the output claims provider hierarchy trees.</param>
        /// <returns>A return value represents a list of claims provider hierarchy trees.</returns>
        SPProviderHierarchyTree[] SearchAll(ArrayOfString providerNames, SPPrincipalType principalType, string searchPattern, int maxCount);
    }
}
