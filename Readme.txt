08.12.2021
I need to be able to select the data based on the number of movements.
To Do:
- how do I select that first (in DB).
- how do I add it as a querry.
13.12.2021
- How to do the querry for number of movements / select the number of movements. 
Analysing Querry 
                SELECT
                    mo.`TotalMovements`,
                    t.`TestDateTime`,
                    `'{0}`' as janus
                FROM
                    `MovementTbl` mo
                INNER JOIN `TestTbl` t ON
                    t.`TestId` = mo.`TestId`
                WHERE mo.`TotalMovements` <> 0
                    AND t.`TestDateTime` > #{1:yyyy-MM-dd HH:mm}#
                ;
