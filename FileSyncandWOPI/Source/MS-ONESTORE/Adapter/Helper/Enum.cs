namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    /// <summary>
    /// The enum of the FileNodeID values.
    /// </summary>
    public enum FileNodeIDValues : uint
    {
        /// <summary>
        ///  Indicate the ObjectSpaceManifestRootFND
        /// </summary>
        ObjectSpaceManifestRootFND = 0x004,
        /// <summary>
        /// Indicate the ObjectSpaceManifestListReferenceFND
        /// </summary>
        ObjectSpaceManifestListReferenceFND = 0x008,
        /// <summary>
        /// Indicate the ObjectSpaceManifestListStartFND 
        /// </summary>
        ObjectSpaceManifestListStartFND = 0x00C,
        /// <summary>
        /// Indicate the RevisionManifestListReferenceFND  
        /// </summary>
        RevisionManifestListReferenceFND = 0x010,
        /// <summary>
        /// Indicate the RevisionManifestListStartFND   
        /// </summary>
        RevisionManifestListStartFND = 0x014,
        /// <summary>
        /// Indicate the RevisionManifestStart4FND    
        /// </summary>
        RevisionManifestStart4FND = 0x01B,
        /// <summary>
        /// Indicate the RevisionManifestEndFND    
        /// </summary>
        RevisionManifestEndFND = 0x01C,
        /// <summary>
        /// Indicate the RevisionManifestStart6FND     
        /// </summary>
        RevisionManifestStart6FND = 0x01E,
        /// <summary>
        /// Indicate the RevisionManifestStart7FND      
        /// </summary>
        RevisionManifestStart7FND = 0x01F,
        /// <summary>
        /// Indicate the GlobalIdTableStartFNDX       
        /// </summary>
        GlobalIdTableStartFNDX = 0x021,
        /// <summary>
        /// Indicate the GlobalIdTableStart2FND       
        /// </summary>
        GlobalIdTableStart2FND = 0x022,
        /// <summary>
        /// Indicate the GlobalIdTableEntryFNDX        
        /// </summary>
        GlobalIdTableEntryFNDX = 0x024,
        /// <summary>
        /// Indicate the GlobalIdTableEntry2FNDX         
        /// </summary>
        GlobalIdTableEntry2FNDX = 0x025,
        /// <summary>
        /// Indicate the GlobalIdTableEntry3FNDX          
        /// </summary>
        GlobalIdTableEntry3FNDX = 0x026,
        /// <summary>
        /// Indicate the GlobalIdTableEndFNDX          
        /// </summary>
        GlobalIdTableEndFNDX = 0x028,
        /// <summary>
        /// Indicate the ObjectDeclarationWithRefCountFNDX           
        /// </summary>
        ObjectDeclarationWithRefCountFNDX = 0x02D,
        /// <summary>
        /// Indicate the ObjectDeclarationWithRefCount2FNDX            
        /// </summary>
        ObjectDeclarationWithRefCount2FNDX = 0x02E,
        /// <summary>
        /// Indicate the ObjectRevisionWithRefCountFNDX             
        /// </summary>
        ObjectRevisionWithRefCountFNDX = 0x041,
        /// <summary>
        /// Indicate the ObjectRevisionWithRefCount2FNDX              
        /// </summary>
        ObjectRevisionWithRefCount2FNDX = 0x042,
        /// <summary>
        /// Indicate the RootObjectReference2FNDX               
        /// </summary>
        RootObjectReference2FNDX = 0x059,
        /// <summary>
        /// Indicate the RootObjectReference3FND                
        /// </summary>
        RootObjectReference3FND = 0x05A,
        /// <summary>
        /// Indicate the RevisionRoleDeclarationFND                 
        /// </summary>
        RevisionRoleDeclarationFND = 0x05C,
        /// <summary>
        /// Indicate the RevisionRoleAndContextDeclarationFND                  
        /// </summary>
        RevisionRoleAndContextDeclarationFND = 0x05D,
        /// <summary>
        /// Indicate the ObjectDeclarationFileData3RefCountFND                   
        /// </summary>
        ObjectDeclarationFileData3RefCountFND = 0x072,
        /// <summary>
        /// Indicate the ObjectDeclarationFileData3LargeRefCountFND                    
        /// </summary>
        ObjectDeclarationFileData3LargeRefCountFND = 0x073,
        /// <summary>
        /// Indicate the ObjectDataEncryptionKeyV2FNDX                     
        /// </summary>
        ObjectDataEncryptionKeyV2FNDX = 0x07C,
        /// <summary>
        /// Indicate the ObjectInfoDependencyOverridesFND                      
        /// </summary>
        ObjectInfoDependencyOverridesFND = 0x084,
        /// <summary>
        /// Indicate the DataSignatureGroupDefinitionFND                       
        /// </summary>
        DataSignatureGroupDefinitionFND = 0x08C,
        /// <summary>
        /// Indicate the FileDataStoreListReferenceFND                        
        /// </summary>
        FileDataStoreListReferenceFND = 0x090,
        /// <summary>
        /// Indicate the FileDataStoreObjectReferenceFND                         
        /// </summary>
        FileDataStoreObjectReferenceFND = 0x094,
        /// <summary>
        /// Indicate the ObjectDeclaration2RefCountFND                          
        /// </summary>
        ObjectDeclaration2RefCountFND = 0x0A4,
        /// <summary>
        /// Indicate the ObjectDeclaration2LargeRefCountFND                           
        /// </summary>
        ObjectDeclaration2LargeRefCountFND = 0x0A5,
        /// <summary>
        /// Indicate the ObjectGroupListReferenceFND                            
        /// </summary>
        ObjectGroupListReferenceFND = 0x0B0,
        /// <summary>
        /// Indicate the ObjectGroupStartFND                             
        /// </summary>
        ObjectGroupStartFND = 0x0B4,
        /// <summary>
        /// Indicate the ObjectGroupEndFND                             
        /// </summary>
        ObjectGroupEndFND = 0x0B8,
        /// <summary>
        /// Indicate the HashedChunkDescriptor2FND                              
        /// </summary>
        HashedChunkDescriptor2FND = 0x0C2,
        /// <summary>
        /// Indicate the ReadOnlyObjectDeclaration2RefCountFND                               
        /// </summary>
        ReadOnlyObjectDeclaration2RefCountFND = 0x0C4,
        /// <summary>
        /// Indicate the ReadOnlyObjectDeclaration2LargeRefCountFND                                
        /// </summary>
        ReadOnlyObjectDeclaration2LargeRefCountFND = 0x0C5,
        /// <summary>
        /// Indicate the ChunkTerminatorFND                                
        /// </summary>
        ChunkTerminatorFND = 0x0FF,
    }
}
