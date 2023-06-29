WITH RECURSIVE fk_constraints AS (
    SELECT
        referencing.name AS referencing_table,
        TRIM(SUBSTR(referencing.sql, INSTR(referencing.sql, 'FOREIGN KEY') + LENGTH('FOREIGN KEY') + 1, INSTR(referencing.sql, 'REFERENCES') - INSTR(referencing.sql, 'FOREIGN KEY') - LENGTH('FOREIGN KEY') - 1)) AS foreign_key_constraint,
        SUBSTR(referencing.sql, INSTR(referencing.sql, 'REFERENCES') + LENGTH('REFERENCES') + 1) AS remaining_sql
    FROM
        sqlite_master AS referencing
    WHERE
        referencing.type = 'table'
        AND referencing.sql LIKE '%FOREIGN KEY%'

    UNION ALL

    SELECT
        referencing_table,
        TRIM(SUBSTR(remaining_sql, INSTR(remaining_sql, 'FOREIGN KEY') + LENGTH('FOREIGN KEY') + 1, INSTR(remaining_sql, 'REFERENCES') - INSTR(remaining_sql, 'FOREIGN KEY') - LENGTH('FOREIGN KEY') - 1)),
        SUBSTR(remaining_sql, INSTR(remaining_sql, 'REFERENCES') + LENGTH('REFERENCES') + 1)
    FROM
        fk_constraints
    WHERE
        remaining_sql LIKE '%FOREIGN KEY%'
)
SELECT
    referencing_table,
    foreign_key_constraint,
    TRIM(SUBSTR(remaining_sql, 0, INSTR(remaining_sql, '('))) AS referenced_table,
    CASE
        WHEN remaining_sql LIKE '%/* { Label:"%' THEN
    TRIM(
        SUBSTR(
            remaining_sql,
            INSTR(remaining_sql, 'Label:"') + LENGTH('Label:"'),
            INSTR(SUBSTR(remaining_sql, INSTR(remaining_sql, 'Label:"') + LENGTH('Label:"')), '"') - 1
        )
    ) 
        ELSE
            NULL
    END AS label
FROM
    fk_constraints
WHERE
    foreign_key_constraint NOT LIKE '%REFERENCES%';