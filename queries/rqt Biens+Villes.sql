SELECT
  [tbl Biens].[Numéro Bien],
  [tbl Biens].[Numéro Propriétaire],
  [tbl Biens].[Nom Bien],
  [tbl Biens].[Adresse 1],
  [tbl Biens].[Adresse 2],
  [tbl Biens].[Code Ville],
  [tbl Biens].[Code Transaction],
  [tbl Biens].[Code Bien],
  [tbl Biens].[Code Catégorie],
  [tbl Biens].Pièces,
  [tbl Biens].Chambres,
  [tbl Biens].Surface,
  [tbl Biens].[Surface Terrain],
  [tbl Biens].Descriptif,
  [tbl Biens].Montant,
  [tbl Biens].Statut,
  [tbl Biens].[Date de création],
  [tbl Biens].[Date de modification],
  [tbl Villes].[Code postal],
  [tbl Villes].[Nom Ville],
  [tbl Villes].Pays,
  [tbl Propriétaires].[Nom Propriétaire],
  [tbl Propriétaires].[Prénom Propriétaire],
  [tbl Propriétaires].Titre,
  [tbl Villes_1].[Code postal] AS [Code postal Propriétaire],
  [tbl Villes_1].[Nom Ville] AS [Nom Ville Propriétaire]
FROM
  [tbl Villes]
  INNER JOIN (
    (
      [tbl Propriétaires]
      LEFT JOIN [tbl Villes] AS [tbl Villes_1] ON [tbl Propriétaires].[Code Ville] = [tbl Villes_1].[Code Ville]
    )
    INNER JOIN [tbl Biens] ON [tbl Propriétaires].[Numéro Propriétaire] = [tbl Biens].[Numéro Propriétaire]
  ) ON [tbl Villes].[Code Ville] = [tbl Biens].[Code Ville];
