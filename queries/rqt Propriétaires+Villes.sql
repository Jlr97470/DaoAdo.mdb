SELECT
  [tbl Propriétaires].[Numéro Propriétaire],
  [tbl Propriétaires].Titre,
  [tbl Propriétaires].[Nom Propriétaire],
  [tbl Propriétaires].[Prénom Propriétaire],
  [tbl Propriétaires].[Adresse 1],
  [tbl Propriétaires].[Adresse 2],
  [tbl Propriétaires].[Code Ville],
  [tbl Propriétaires].[Téléphone personnel],
  [tbl Propriétaires].[Téléphone mobile],
  [tbl Propriétaires].[Téléphone bureau],
  [tbl Propriétaires].Email,
  [tbl Propriétaires].[Site web],
  [tbl Propriétaires].Remarques,
  [tbl Propriétaires].[Date de création],
  [tbl Propriétaires].[Date de modification],
  [tbl Propriétaires].Archiver,
  [tbl Villes].[Code postal],
  [tbl Villes].[Nom Ville],
  [tbl Villes].Pays
FROM
  [tbl Villes]
  INNER JOIN [tbl Propriétaires] ON [tbl Villes].[Code Ville] = [tbl Propriétaires].[Code Ville]
ORDER BY
  [tbl Propriétaires].[Nom Propriétaire],
  [tbl Propriétaires].[Prénom Propriétaire];
