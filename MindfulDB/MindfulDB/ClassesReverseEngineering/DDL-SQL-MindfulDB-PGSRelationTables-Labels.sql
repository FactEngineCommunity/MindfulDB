WITH RECURSIVE pk_constraints AS (
    SELECT
        referencing.name AS referencing_table,
        TRIM(SUBSTR(referencing.sql, INSTR(referencing.sql, 'CONSTRAINT') + LENGTH('CONSTRAINT'), INSTR(referencing.sql, 'PRIMARY KEY') - INSTR(referencing.sql, 'CONSTRAINT') - LENGTH('CONSTRAINT'))) AS primary_key_constraint,
        SUBSTR(referencing.sql, INSTR(referencing.sql, 'PRIMARY KEY') + LENGTH('PRIMARY KEY') + 1) AS remaining_sql
    FROM
        sqlite_master AS referencing
    WHERE
        referencing.type = 'table'
        AND referencing.sql LIKE '%PRIMARY KEY%'
        AND referencing.sql LIKE '%/* { Label:"%'

    UNION ALL

    SELECT
        referencing_table,
        TRIM(SUBSTR(remaining_sql, INSTR(remaining_sql, 'CONSTRAINT') + LENGTH('CONSTRAINT'), INSTR(remaining_sql, 'PRIMARY KEY') - INSTR(remaining_sql, 'CONSTRAINT') - LENGTH('CONSTRAINT'))),
        SUBSTR(remaining_sql, INSTR(remaining_sql, 'PRIMARY KEY') + LENGTH('PRIMARY KEY') + 1)
    FROM
        pk_constraints
    WHERE
        remaining_sql LIKE '%PRIMARY KEY%'
        AND remaining_sql REGEXP  '^(?!.*FOREIGN KEY.*{ Label).*{ Label'

)
SELECT
    referencing_table AS [Table],
    primary_key_constraint,
    CASE
        WHEN remaining_sql REGEXP '^(?!.*FOREIGN KEY.*{ Label).*{ Label' THEN
            TRIM(
                REPLACE(
                    SUBSTR(
                        remaining_sql,
                        INSTR(remaining_sql, '/* { Label:') + LENGTH('/* { Label:'),
                        INSTR(remaining_sql, '} */') - INSTR(remaining_sql, '/* { Label:') - LENGTH('/* { Label:')
                    ),
                    '"',
                    ''
                )
            )
        ELSE
            NULL
    END AS label,
    remaining_sql
FROM
    pk_constraints
    WHERE remaining_sql REGEXP '^(?!.*FOREIGN KEY.*{ Label).*{ Label'