PARAMETERS [Date de début] DateTime,
[Date de fin] DateTime;
SELECT
  [tbl Visites].[Numéro Bien],
  [tbl Visites].[Numéro Client],
  [tbl Visites].[Date Visite],
  [tbl Visites].[Prochain contact],
  [tbl Visites].Intérêt,
  [tbl Visites].Remarques
FROM
  [tbl Visites]
WHERE
  (
    (
      (
        [tbl Visites].[Prochain contact]
      ) Between [Date de début]
      And [Date de fin]
    )
  )
ORDER BY
  [tbl Visites].[Prochain contact];
