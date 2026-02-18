CREATE TABLE [tbl Recherches] (
  [Numéro Recherche] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Numéro Client] LONG,
  [Type de transaction] VARCHAR (4),
  [Type de bien] VARCHAR (4),
  [Pièces] BYTE,
  [Chambres] SHORT,
  [Surface habitable] SHORT,
  [Surface du terrain] LONG,
  [Prix minimum] CURRENCY,
  [Prix maximum] CURRENCY,
  [Remarques] LONGTEXT,
  [Statut] VARCHAR (20)
)
