--- DLM
SELECT
  g.id                    AS game_id,
  g.name                  AS name,
  CASE 
    WHEN g.type = 'INSTANT' THEN 1
    ELSE 0
  END                     AS is_instant,
  CASE
    WHEN g.parent IS NOT NULL THEN g.parent
    ELSE g.id
  END                     AS parent_id,
  SUM(s.sales)            AS sales,
  SUM(s.sales) / g.price  AS qty,
  COUNT(DISTINCT t.id)    AS tcount
FROM
  sales AS s
  INNER JOIN games AS g
    ON s.game_id = g.id
  INNER JOIN terminals AS t
    ON s.terminal_id = t.id
WHERE
  s.date BETWEEN '2016-02-01' AND '2016-02-29'
  AND t.agent_id <> 225
GROUP BY g.id
HAVING qty > 0
ORDER BY is_instant, parent_id, g.id

SELECT
  CASE 
    WHEN g.type = 'INSTANT' THEN 1
    ELSE 0
  END                     AS is_instant,
  SUM(s.sales)            AS sales,
  COUNT(DISTINCT t.id)    AS tcount
FROM
  sales AS s
  INNER JOIN games AS g
    ON s.game_id = g.id
  INNER JOIN terminals AS t
    ON s.terminal_id = t.id
WHERE
  s.date BETWEEN '2016-02-01' AND '2016-02-29'
  AND t.agent_id <> 225
  AND s.sales > 0
GROUP BY is_instant
ORDER BY is_instant

SELECT
  SUM(s.sales),
  COUNT(DISTINCT t.id)
FROM
  sales AS s
  INNER JOIN terminals AS t
    ON s.terminal_id = t.id
  INNER JOIN games AS g
    ON s.game_id = g.id
WHERE
  s.date BETWEEN '2016-02-01' AND '2016-02-29'
  AND t.agent_id <> 225
  AND s.sales > 0
  AND g.type <> 'INSTANT'


SELECT
  SUM(s.sales),
  COUNT(DISTINCT t.id)
FROM
  sales AS s
  INNER JOIN terminals AS t
    ON s.terminal_id = t.id
  INNER JOIN games AS g
    ON s.game_id = g.id
WHERE
  s.date BETWEEN '2016-02-01' AND '2016-02-29'
  AND t.agent_id = 256
  AND s.sales > 0



--- DLM
SELECT
  g.id                    AS game_id,
  g.name                  AS name,
  g.price                 AS price,
  CASE 
    WHEN g.type = 'INSTANT' THEN 1
    ELSE 0
  END                     AS is_instant,
  CASE
    WHEN g.parent IS NOT NULL THEN g.parent
    ELSE g.id
  END                     AS parent_id,
  SUM(s.sales)            AS sales,
  SUM(s.sales) / g.price  AS qty,
  COUNT(DISTINCT t.id)    AS tcount
FROM
  sales AS s
  INNER JOIN games AS g
    ON s.game_id = g.id
  INNER JOIN terminals AS t
    ON s.terminal_id = t.id
WHERE
  s.date BETWEEN '2016-02-01' AND '2016-02-29'
  AND t.agent_id <> 225
GROUP BY g.id
HAVING qty > 0
ORDER BY is_instant, parent_id, g.id

--- marketing
SELECT
  g.id                    AS game_id,
  g.name                  AS name,
  g.price                 AS price,
  CASE 
    WHEN g.type = 'INSTANT' THEN 1
    ELSE 0
  END                     AS is_instant,
  CASE
    WHEN g.parent IS NOT NULL THEN g.parent
    ELSE g.id
  END                     AS parent_id,
  SUM(s.sales)            AS sales,
  SUM(s.sales) / g.price  AS qty,
  COUNT(DISTINCT t.id)    AS tcount
FROM
  sales AS s
  INNER JOIN games AS g
    ON s.game_id = g.id
  INNER JOIN terminals AS t
    ON s.terminal_id = t.id
WHERE
  s.date BETWEEN '2016-02-01' AND '2016-02-29'
  AND t.agent_id = 225
GROUP BY g.id
HAVING qty > 0
ORDER BY is_instant, parent_id, g.id
