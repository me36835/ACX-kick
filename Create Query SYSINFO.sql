CREATE QUERY [SYSINFO] AS
SELECT DISTINCT sys.Name, sys.Type
FROM MSysObjects AS sys LEFT JOIN GIT ON git.mdl = sys.Name
WHERE sys.Type IN (-32768, -32764, -32761)
  AND (git.vrs = 0 OR git.vrs IS NULL)
ORDER BY sys.Name DESC;

