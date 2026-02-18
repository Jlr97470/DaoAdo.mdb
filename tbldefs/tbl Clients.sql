CREATE TABLE [tbl Clients] (
  [Numéro Client] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Titre] VARCHAR (20),
  [Nom Client] VARCHAR (30),
  [Prénom Client] VARCHAR (20),
  [Adresse 1] VARCHAR (50),
  [Adresse 2] VARCHAR (50),
  [Code Ville] LONG,
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
