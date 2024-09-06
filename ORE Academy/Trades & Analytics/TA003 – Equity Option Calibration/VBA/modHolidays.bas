Attribute VB_Name = "modHolidays"
Option Explicit

'This file is inspired by QuantLib C++ code, a free-software/open-source library _
for financial quantitative analysts and developers - http://quantlib.org/ _
This program is distributed in the hope that it will be useful, but WITHOUT _
ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS _
FOR A PARTICULAR PURPOSE.  See the license for more details.

'Return the list of holidays between 2 years
Function getListHolidays(startYear As Long, _
                         endYear As Long) As Collection
                    
    Dim holidaysColl As New Collection
    Dim year As Long
    
    For year = startYear To endYear
        Call addYearHolidaysToCollection(year, holidaysColl)
    Next year
    
    Set getListHolidays = holidaysColl
    
End Function
                    
'Add the holidays date to the collection for a given year
Sub addYearHolidaysToCollection(year As Long, _
                                ByRef holidaysColl As Collection)
        
    'New Year Day
    holidaysColl.Add DateSerial(year, 1, 1)
    
    'Easter Friday Monday
    Dim easterMondayDate As Date: easterMondayDate = getYearEasterMonday(year)
    holidaysColl.Add easterMondayDate - 3
    holidaysColl.Add easterMondayDate
     
    'Labour Day
    holidaysColl.Add DateSerial(year, 5, 1)
    
    'Christmas
    holidaysColl.Add DateSerial(year, 12, 25)
    
    'Day of Goodwill
    holidaysColl.Add DateSerial(year, 12, 26)
    
End Sub

'Return the Easter Day for the given year
Function getYearEasterMonday(year As Long) As Date
    
    'Day of the year that corresponds to the Easter monday _
    Each line corresponds to a deceny from 2000 to 2199 (first line is 2000 to 2009)
    Dim easterMonday() As Variant: easterMonday = Array(115, 106, 91, 111, 103, 87, 107, 99, 84, 103, _
                                                         95, 115, 100, 91, 111, 96, 88, 107, 92, 112, _
                                                        104, 95, 108, 100, 92, 111, 96, 88, 108, 92, _
                                                        112, 104, 89, 108, 100, 85, 105, 96, 116, 101, _
                                                         93, 112, 97, 89, 109, 100, 85, 105, 97, 109, _
                                                        101, 93, 113, 97, 89, 109, 94, 113, 105, 90, _
                                                        110, 101, 86, 106, 98, 89, 102, 94, 114, 105, _
                                                         90, 110, 102, 86, 106, 98, 111, 102, 94, 114, _
                                                         99, 90, 110, 95, 87, 106, 91, 111, 103, 94, _
                                                        107, 99, 91, 103, 95, 115, 107, 91, 111, 103, _
                                                         88, 108, 100, 85, 105, 96, 109, 101, 93, 112, _
                                                         97, 89, 109, 93, 113, 105, 90, 109, 101, 86, _
                                                        106, 97, 89, 102, 94, 113, 105, 90, 110, 101, _
                                                         86, 106, 98, 110, 102, 94, 114, 98, 90, 110, _
                                                         95, 86, 106, 91, 111, 102, 94, 107, 99, 90, _
                                                        103, 95, 115, 106, 91, 111, 103, 87, 107, 99, _
                                                         84, 103, 95, 115, 100, 91, 111, 96, 88, 107, _
                                                         92, 112, 104, 95, 108, 100, 92, 111, 96, 88, _
                                                        108, 92, 112, 104, 89, 108, 100, 85, 105, 96, _
                                                        116, 101, 93, 112, 97, 89, 109, 100, 85, 105)
    
    Dim easterDay As Integer: easterDay = easterMonday(year - 2000)
    Dim easterDate As Date: easterDate = DateSerial(year, 1, 1) + easterDay - 1
    
    getYearEasterMonday = easterDate
    
End Function

'Create the list of holiday dates for our example (from 2021 to 2081 given our maximum tenor is 60Years)
Sub main()
    
    Dim holidaysColl As Collection: Set holidaysColl = getListHolidays(2021, 2081)
    Dim nbDates As Long: nbDates = holidaysColl.Count
    Dim holidayDatesRange As Range: Set holidayDatesRange = Range("rngListHolidayDatesHeader").Offset(1, 0).Resize(nbDates, 1)
    Dim holidayDatesList() As Variant
    Dim datei As Long
    
    'Convert collection to 2D array
    ReDim holidayDatesList(1 To nbDates, 1 To 1)
    For datei = 1 To nbDates
        holidayDatesList(datei, 1) = holidaysColl.Item(datei)
    Next datei
    
    'Write the dates to Excel
    holidayDatesRange = holidayDatesList
    
End Sub
