SELECT
  s.fdom,
  s.game_type,
  SUM(s.sales) total_sales,
  SUM(s.qty)  total_qty
FROM (
  SELECT
    date(s.date, 'start of month')
                  AS fdom,
    g.id          AS game_id,
    g.name        AS game_name,
    CASE 
      WHEN g.parent IS NOT NULL THEN g.parent
      ELSE g.id
    END           AS parent_id,
    g2.type       AS game_type,
    s.sales       AS sales,
    s.sales/g.price
                  AS  qty
  FROM
    games AS g
    INNER JOIN games AS g2
      ON g2.id = parent_id
    INNER JOIN sales AS s
      ON g.id = s.game_id
---  WHERE
---    s.date BETWEEN '2016-01-01' AND '2016-02-29'
) AS s
GROUP BY s.fdom, s.game_type
HAVING total_sales > 0.0
ORDER BY s.fdom, s.game_type
