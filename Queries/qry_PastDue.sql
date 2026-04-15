SELECT q.*, addWorkdays(
        [RequestDateReal],
        IIf([UnitName]="U6",5,
            IIf([UnitName]="U17" Or [UnitName]="U21",7,3)
        )
    ) AS DueDate, countWorkdays(
        addWorkdays(
            [RequestDateReal],
            IIf([UnitName]="U6",5,
                IIf([UnitName]="U17" Or [UnitName]="U21",7,3)
            )
        ),
        Date()
    ) AS WorkdaysPastDue
FROM qry_Reports AS q
WHERE [responseDate] Is Null
    AND [UnitName] <> "U7"
    AND Date() > addWorkdays(
        [RequestDateReal],
        IIf([UnitName]="U6",5,
            IIf([UnitName]="U17" Or [UnitName]="U21",7,3)
        )
    );

