CREATE TABLE MessageIDinProgress (
	[RecID] BigInt IDENTITY(1,1) PRIMARY KEY NOT NULL,
	[MessageID] CHAR(255) NOT NULL,
    [CreationDateTime] DATETIME NOT NULL);
    
ALTER TABLE MessageIDinProgress ADD CONSTRAINT     
    Pk_MessageIDinProgress PRIMARY KEY NONCLUSTERED
    ( RecID ) ;
      
GO


