PARAMETERS Département Text (255); DELETE [tbl Villes Copie].[Code Ville],
[tbl Villes Copie].[Nom Ville],
[tbl Villes Copie].[Code postal],
[tbl Villes Copie].Pays
FROM
  [tbl Villes Copie]
WHERE
  (
    (
      (
        [tbl Villes Copie].[Code postal]
      ) Like [Département] & "%"
    )
  );
