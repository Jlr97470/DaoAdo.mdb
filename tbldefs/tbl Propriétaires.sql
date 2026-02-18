CREATE TABLE [tbl Propriétaires] (
  [Numéro Propriétaire] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Titre] VARCHAR (20),
  [Nom Propriétaire] VARCHAR (30),
  [Prénom Propriétaire] VARCHAR (20),
  [Adresse 1] VARCHAR (50),
  [Adresse 2] VARCHAR (50),
  [Code Ville] LONG CONSTRAINT [{0940E38E-80E8-4E7F-BAAC-D3BBDD308695}] REFERENCES [tbl Villes] ([Code Ville]) ON UPDATE CASCADE  ON DELETE CASCADE ,
  [Téléphone personnel] VARCHAR (25),
  [Téléphone mobile] VARCHAR (25),
  [Téléphone bureau] VARCHAR (25),
  [Email] VARCHAR (100),
  [Site web] VARCHAR (50),
  [Remarques] LONGTEXT,
  [Date de création] DATETIME,
  [Date de modification] DATETIME,
  [Archiver] BIT
)
