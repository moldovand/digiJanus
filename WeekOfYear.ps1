function Get-ISO8601Week {
     Param(
     [datetime]$DT = (Get-Date)
     )
     <#
     First create an integer(0/1) from the boolean,
     "Is the integer DayOfWeek value greater than zero?".
     Then Multiply it with 4 or 6 (weekrule = 0 or 2) minus the integer DayOfWeek value.
     This turns every day (except Sunday) into Thursday.
     Then return the ISO8601 WeekNumber.
     #>
     $Cult = Get-Culture; $DT = Get-Date($DT)
     $WeekRule = $Cult.DateTimeFormat.CalendarWeekRule.value__
     $FirstDayOfWeek = $Cult.DateTimeFormat.FirstDayOfWeek.value__
     $WeekRuleDay = [int]($DT.DayOfWeek.Value__ -ge $FirstDayOfWeek ) * ( (6 - $WeekRule) - $DT.DayOfWeek.Value__ )
     $Cult.Calendar.GetWeekOfYear(($DT).AddDays($WeekRuleDay), $WeekRule, $FirstDayOfWeek)
}

Get-ISO8601Week '2012-11-30'
