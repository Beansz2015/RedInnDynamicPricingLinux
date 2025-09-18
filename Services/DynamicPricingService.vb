﻿Imports System.Net.Http
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports System.Text
Imports System.IO
Imports Microsoft.Extensions.Configuration  ' NEW - replaces System.Configuration
Imports System.Net.Mail
Imports System.Net
Imports TimeZoneConverter

Public Class DynamicPricingService
    Private ReadOnly httpClient As New HttpClient()
    Private ReadOnly googleSheetsService As GoogleSheetsService
    Private ReadOnly configuration As IConfiguration ' NEW

    ' Configuration settings - NEW way to read config
    Private ReadOnly apiKey As String
    Private ReadOnly propertyId As String
    Private ReadOnly apiBaseUrl As String

    ' Email Settings
    Private ReadOnly smtpHost As String
    Private ReadOnly smtpPort As Integer
    Private ReadOnly smtpUsername As String
    Private ReadOnly smtpPassword As String
    Private ReadOnly emailFromAddress As String
    Private ReadOnly emailFromName As String
    Private ReadOnly emailToAddress As String

    ' Room capacity configuration with temporary closures
    Private ReadOnly baseDormBeds As Integer = 20
    Private ReadOnly basePrivateRooms As Integer = 3
    Private ReadOnly baseEnsuiteRooms As Integer = 2

    ' Calculate effective capacity minus temporarily closed units
    Private ReadOnly totalDormBeds As Integer
    Private ReadOnly totalPrivateRooms As Integer
    Private ReadOnly totalEnsuiteRooms As Integer

    ' UPDATED CONSTRUCTOR - Accept IConfiguration
    Public Sub New(config As IConfiguration)
        configuration = config

        ' Load configuration values
        apiKey = configuration("LittleHotelier:ApiKey")
        propertyId = configuration("LittleHotelier:PropertyId")
        apiBaseUrl = configuration("LittleHotelier:ApiBaseUrl")

        ' Email Settings
        smtpHost = configuration("Email:SmtpHost")
        smtpPort = CInt(configuration("Email:SmtpPort"))
        smtpUsername = configuration("Email:Username")
        smtpPassword = configuration("Email:Password")
        emailFromAddress = configuration("Email:FromAddress")
        emailFromName = configuration("Email:FromName")
        emailToAddress = configuration("Email:ToAddress")

        ' Calculate room capacities
        totalDormBeds = baseDormBeds - CInt(If(configuration("RoomCapacity:TemporaryClosedDormBeds"), "0"))
        totalPrivateRooms = basePrivateRooms - CInt(If(configuration("RoomCapacity:TemporaryClosedPrivateRooms"), "0"))
        totalEnsuiteRooms = baseEnsuiteRooms - CInt(If(configuration("RoomCapacity:TemporaryClosedEnsuiteRooms"), "0"))

        ' Initialize Google Sheets service
        Try
            googleSheetsService = New GoogleSheetsService(configuration)
            Console.WriteLine("Google Sheets service initialized successfully")
        Catch ex As Exception
            Console.WriteLine($"Warning: Google Sheets service initialization failed: {ex.Message}")
            googleSheetsService = Nothing
        End Try
    End Sub


    ' Business timezone helper
    Public Function GetBusinessDateTime() As DateTime
        Try
            Dim timeZoneInfo As TimeZoneInfo

            ' Use cross-platform timezone handling
            Try
                ' Try Windows timezone ID first
                timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Singapore Standard Time")
            Catch
                ' Fallback to IANA timezone ID for Linux
                timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Asia/Singapore")
            End Try

            Return TimeZoneInfo.ConvertTime(DateTime.UtcNow, timeZoneInfo)
        Catch ex As Exception
            Console.WriteLine($"Warning: Timezone conversion failed, using manual UTC+8: {ex.Message}")
            ' Manual fallback to Singapore time
            Return DateTime.UtcNow.AddHours(8)
        End Try
    End Function

    Private Function GetBusinessToday() As DateTime
        Return GetBusinessDateTime().Date
    End Function

    ' NEW: Check if current time is in email quiet period (using Google Sheets)
    Private Function IsEmailQuietPeriod() As Boolean
        Try
            If googleSheetsService IsNot Nothing AndAlso googleSheetsService.IsGoogleSheetsEnabled() Then
                Dim quietPeriods = googleSheetsService.GetQuietPeriods()
                Dim now = GetBusinessDateTime()
                Dim currentDay = now.DayOfWeek
                Dim currentTime = now.TimeOfDay

                For Each period In quietPeriods
                    If period.Enabled AndAlso period.DayOfWeek = currentDay Then
                        ' Handle periods that span midnight
                        If period.EndTime < period.StartTime Then
                            ' Period spans midnight (e.g., 22:00 to 06:00)
                            If currentTime >= period.StartTime OrElse currentTime < period.EndTime Then
                                Return True
                            End If
                        Else
                            ' Normal period (e.g., 08:00 to 16:00)
                            If currentTime >= period.StartTime AndAlso currentTime < period.EndTime Then
                                Return True
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            Console.WriteLine($"Error checking quiet periods from Google Sheets: {ex.Message}")
        End Try

        ' Fallback to hardcoded periods if Google Sheets fails
        Return IsEmailQuietPeriodFallback()
    End Function

    ' NEW: Get description of current quiet period (using Google Sheets)
    Private Function GetQuietPeriodDescription() As String
        Try
            If googleSheetsService IsNot Nothing AndAlso googleSheetsService.IsGoogleSheetsEnabled() Then
                Dim quietPeriods = googleSheetsService.GetQuietPeriods()
                Dim now = GetBusinessDateTime()
                Dim currentDay = now.DayOfWeek
                Dim currentTime = now.TimeOfDay

                For Each period In quietPeriods
                    If period.Enabled AndAlso period.DayOfWeek = currentDay Then
                        Dim isInPeriod As Boolean = False

                        ' Handle periods that span midnight
                        If period.EndTime < period.StartTime Then
                            isInPeriod = (currentTime >= period.StartTime OrElse currentTime < period.EndTime)
                        Else
                            isInPeriod = (currentTime >= period.StartTime AndAlso currentTime < period.EndTime)
                        End If

                        If isInPeriod Then
                            Return $"{period.Name} ({period.StartTime:hh\:mm}-{period.EndTime:hh\:mm})"
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            Console.WriteLine($"Error getting quiet period description from Google Sheets: {ex.Message}")
        End Try

        ' Fallback to hardcoded description
        Return GetQuietPeriodDescriptionFallback()
    End Function

    ' FALLBACK: Hardcoded quiet period check (your original logic)
    Private Function IsEmailQuietPeriodFallback() As Boolean
        Dim now = GetBusinessDateTime()
        Dim dayOfWeek = now.DayOfWeek
        Dim timeOfDay = now.TimeOfDay

        ' 1. 12am - 8am every day (00:00 - 08:00)
        If timeOfDay >= New TimeSpan(0, 0, 0) AndAlso timeOfDay < New TimeSpan(8, 0, 0) Then
            Return True
        End If

        ' 2. 4pm - 12am every Saturday (16:00 - 23:59)
        If dayOfWeek = DayOfWeek.Saturday AndAlso timeOfDay >= New TimeSpan(16, 0, 0) Then
            Return True
        End If

        ' 3. 8am - 4pm every Sunday (08:00 - 16:00)
        If dayOfWeek = DayOfWeek.Sunday AndAlso timeOfDay >= New TimeSpan(8, 0, 0) AndAlso timeOfDay < New TimeSpan(16, 0, 0) Then
            Return True
        End If

        Return False
    End Function

    ' FALLBACK: Hardcoded quiet period description
    Private Function GetQuietPeriodDescriptionFallback() As String
        Dim now = GetBusinessDateTime()
        Dim dayOfWeek = now.DayOfWeek
        Dim timeOfDay = now.TimeOfDay

        If timeOfDay >= New TimeSpan(0, 0, 0) AndAlso timeOfDay < New TimeSpan(8, 0, 0) Then
            Return "Daily quiet period (12am - 8am)"
        ElseIf dayOfWeek = DayOfWeek.Saturday AndAlso timeOfDay >= New TimeSpan(16, 0, 0) Then
            Return "Saturday evening quiet period (4pm - 12am)"
        ElseIf dayOfWeek = DayOfWeek.Sunday AndAlso timeOfDay >= New TimeSpan(8, 0, 0) AndAlso timeOfDay < New TimeSpan(16, 0, 0) Then
            Return "Sunday daytime quiet period (8am - 4pm)"
        Else
            Return "Unknown quiet period"
        End If
    End Function

    Public Async Function RunDynamicPricingCheck() As Task
        Dim errorToReport As String = Nothing
        Dim propertyData As LittleHotelierResponse = Nothing

        Try
            Console.WriteLine($"Starting dynamic pricing check at {DateTime.Now}")

            ' Get availability and store the API response
            Dim availabilityResult = Await GetRoomAvailabilityWithResponseAsync()

            If availabilityResult.AvailabilityData.Count = 0 Then
                Console.WriteLine("No availability data retrieved. Exiting.")
                Return
            End If

            ' Calculate new rates and detect changes using API data
            Dim rateChanges = CalculateRateChanges(availabilityResult.AvailabilityData, availabilityResult.PropertyData)

            ' Send email notifications if there are changes
            If rateChanges.Any() Then
                Await SendEmailNotificationAsync(rateChanges)
                Console.WriteLine($"Processed {rateChanges.Count} rate changes")
            Else
                Console.WriteLine("No rate changes detected")
            End If

            Console.WriteLine("Dynamic pricing check completed successfully")

        Catch ex As Exception
            Console.WriteLine($"Error in dynamic pricing check: {ex.Message}")
            Console.WriteLine($"Stack trace: {ex.StackTrace}")
            errorToReport = ex.Message
        End Try

        ' Send error notification if there was an error
        If errorToReport IsNot Nothing Then
            Try
                Await SendEmailErrorNotificationAsync(errorToReport)
            Catch notificationEx As Exception
                Console.WriteLine($"Failed to send email error notification: {notificationEx.Message}")
            End Try
        End If
    End Function

    Public Async Function GetRoomAvailabilityAsync() As Task(Of Dictionary(Of String, RoomAvailability))
        Dim availabilityData As New Dictionary(Of String, RoomAvailability)

        Try
            Dim startDate = GetBusinessDateTime().ToString("yyyy-MM-dd")
            Dim url = $"{apiBaseUrl}properties/{propertyId}/rates.json?start_date={startDate}"

            Console.WriteLine($"API URL: {url}")

            httpClient.DefaultRequestHeaders.Clear()
            httpClient.DefaultRequestHeaders.Add("Accept", "application/json")
            httpClient.DefaultRequestHeaders.Add("User-Agent", "RedInnDynamicPricing/1.0")

            Dim response = Await httpClient.GetAsync(url)

            If response.IsSuccessStatusCode Then
                Dim jsonContent = Await response.Content.ReadAsStringAsync()
                Console.WriteLine("Successfully retrieved API data")
                Console.WriteLine($"Response length: {jsonContent.Length} characters")

                Dim apiDataArray = JsonConvert.DeserializeObject(Of List(Of LittleHotelierResponse))(jsonContent)

                If apiDataArray.Count > 0 Then
                    Dim propertyData = apiDataArray(0)
                    availabilityData = ParseAvailabilityByDate(propertyData)
                    Console.WriteLine($"Parsed availability for {availabilityData.Count} dates")
                End If
            Else
                Dim errorContent = Await response.Content.ReadAsStringAsync()
                Console.WriteLine($"API Error: {response.StatusCode} - {errorContent}")
                Console.WriteLine($"Request URL: {url}")
            End If

        Catch ex As Exception
            Console.WriteLine($"Error fetching room availability: {ex.Message}")
            Throw
        End Try

        Return availabilityData
    End Function

    Private Async Function GetRoomAvailabilityWithResponseAsync() As Task(Of (AvailabilityData As Dictionary(Of String, RoomAvailability), PropertyData As LittleHotelierResponse))
        Dim availabilityData As New Dictionary(Of String, RoomAvailability)
        Dim propertyData As LittleHotelierResponse = Nothing

        Try
            Dim startDate = GetBusinessDateTime().ToString("yyyy-MM-dd")
            Dim url = $"{apiBaseUrl}properties/{propertyId}/rates.json?start_date={startDate}"

            httpClient.DefaultRequestHeaders.Clear()
            httpClient.DefaultRequestHeaders.Add("Accept", "application/json")
            httpClient.DefaultRequestHeaders.Add("User-Agent", "RedInnDynamicPricing/1.0")

            Dim response = Await httpClient.GetAsync(url)

            If response.IsSuccessStatusCode Then
                Dim jsonContent = Await response.Content.ReadAsStringAsync()
                Dim apiDataArray = JsonConvert.DeserializeObject(Of List(Of LittleHotelierResponse))(jsonContent)

                If apiDataArray.Count > 0 Then
                    propertyData = apiDataArray(0)
                    availabilityData = ParseAvailabilityByDate(propertyData)
                End If
            End If
        Catch ex As Exception
            Console.WriteLine($"Error fetching room availability: {ex.Message}")
            Throw
        End Try

        Return (availabilityData, propertyData)
    End Function

    Private Function ParseAvailabilityByDate(propertyData As LittleHotelierResponse) As Dictionary(Of String, RoomAvailability)
        Dim availabilityByDate As New Dictionary(Of String, RoomAvailability)

        ' Get only 3 days: Today, Day +1, Day >+2
        Dim targetDates As New List(Of String)
        For i As Integer = 0 To 2
            targetDates.Add(GetBusinessDateTime().AddDays(i).ToString("yyyy-MM-dd"))
        Next

        ' Process each target date
        For Each targetDate In targetDates
            Dim availability As New RoomAvailability With {
                .CheckDate = targetDate,
                .DormBedsAvailable = 0,
                .PrivateRoomsAvailable = 0,
                .PrivateEnsuitesAvailable = 0
            }

            ' Sum up availability by room category
            For Each ratePlan In propertyData.rate_plans
                Dim dateEntry = ratePlan.rate_plan_dates.FirstOrDefault(Function(d) d.date = targetDate)

                If dateEntry IsNot Nothing Then
                    Select Case ratePlan.name.ToLower()
                        Case "2 bed mixed dorm", "4 bed female dorm", "4 bed mixed dorm", "6 bed mixed dorm"
                            availability.DormBedsAvailable += dateEntry.available
                        Case "superior double (shared bathroom)"
                            availability.PrivateRoomsAvailable += dateEntry.available
                        Case "superior queen ensuite"
                            availability.PrivateEnsuitesAvailable += dateEntry.available
                    End Select
                End If
            Next

            availabilityByDate(targetDate) = availability

            ' Display availability and calculated rates
            Dim calculatedRates = GetCurrentRates(availability, targetDate)
            Dim daysAhead = DateDiff(DateInterval.Day, GetBusinessToday(), DateTime.Parse(targetDate))
            Dim dayLabel = If(daysAhead = 0, "Today", If(daysAhead = 1, "Day +1", "Day >+2"))

            Dim dayType = IsWeekendOrHoliday(targetDate)
            Dim dayTypeIndicator = If(dayType = "Regular", "", $" [{dayType}]")

            Console.WriteLine($"Date: {targetDate} ({dayLabel}){dayTypeIndicator}")
            Console.WriteLine($"  Dorms: {availability.DormBedsAvailable}/{totalDormBeds} available - RM{calculatedRates.DormRegularRate}/RM{calculatedRates.DormWalkInRate}")
            Console.WriteLine($"  Private: {availability.PrivateRoomsAvailable}/{totalPrivateRooms} available - RM{calculatedRates.PrivateRegularRate}/RM{calculatedRates.PrivateWalkInRate}")
            Console.WriteLine($"  Ensuite: {availability.PrivateEnsuitesAvailable}/{totalEnsuiteRooms} available - RM{calculatedRates.EnsuiteRegularRate}/RM{calculatedRates.EnsuiteWalkInRate}")
        Next

        Return availabilityByDate
    End Function

    Private Function GetApiRatesFromResponse(propertyData As LittleHotelierResponse, targetDate As String) As PreviousRate
        Dim apiRates As New PreviousRate With {
            .DormRegularRate = 0,
            .DormWalkInRate = 0,
            .PrivateRegularRate = 0,
            .PrivateWalkInRate = 0,
            .EnsuiteRegularRate = 0,
            .EnsuiteWalkInRate = 0
        }

        ' Add null checks for propertyData and rate_plans
        If propertyData Is Nothing OrElse propertyData.rate_plans Is Nothing Then
            Return apiRates
        End If

        For Each ratePlan In propertyData.rate_plans
            ' Check if ratePlan and its properties are not null
            If ratePlan Is Nothing OrElse ratePlan.rate_plan_dates Is Nothing OrElse String.IsNullOrEmpty(ratePlan.name) Then
                Continue For
            End If

            Dim dateEntry = ratePlan.rate_plan_dates.FirstOrDefault(Function(d) d IsNot Nothing AndAlso d.date = targetDate)

            ' Check if dateEntry was found and is not null
            If dateEntry Is Nothing Then
                Continue For
            End If

            Select Case ratePlan.name.ToLower()
                Case "2 bed mixed dorm", "4 bed female dorm", "4 bed mixed dorm", "6 bed mixed dorm"
                    ' Only set if not already set (first match wins)
                    If apiRates.DormRegularRate = 0 Then
                        apiRates.DormRegularRate = CDbl(dateEntry.rate)
                    End If
                Case "superior double (shared bathroom)"
                    If apiRates.PrivateRegularRate = 0 Then
                        apiRates.PrivateRegularRate = CDbl(dateEntry.rate)
                    End If
                Case "superior queen ensuite"
                    If apiRates.EnsuiteRegularRate = 0 Then
                        apiRates.EnsuiteRegularRate = CDbl(dateEntry.rate)
                    End If
            End Select
        Next

        Return apiRates
    End Function

    Private Function GetCurrentRates(avail As RoomAvailability, dateStr As String) As PreviousRate
        ' Get base rates (existing logic)
        Dim daysAhead = DateDiff(DateInterval.Day, GetBusinessToday(), DateTime.Parse(dateStr))
        Dim dayPrefix As String = If(daysAhead = 0, "Today", If(daysAhead = 1, "Day1_", "Day2Plus_"))

        ' Get base rates
        Dim dormRates = GetRoomTypeRates("Dorm", dayPrefix, avail.DormBedsAvailable)
        Dim privateRates = GetRoomTypeRates("Private", dayPrefix, avail.PrivateRoomsAvailable)
        Dim ensuiteRates = GetRoomTypeRates("Ensuite", dayPrefix, avail.PrivateEnsuitesAvailable)

        ' Apply surcharges
        Dim dormSurcharge = CalculateRateWithSurcharges(dormRates.RegularRate, dateStr)
        Dim privateSurcharge = CalculateRateWithSurcharges(privateRates.RegularRate, dateStr)
        Dim ensuiteSurcharge = CalculateRateWithSurcharges(ensuiteRates.RegularRate, dateStr)

        Return New PreviousRate With {
        .DormRegularRate = dormSurcharge.FinalRate,
        .DormWalkInRate = dormRates.WalkInRate,  ' Walk-in rates don't get surcharge
        .PrivateRegularRate = privateSurcharge.FinalRate,
        .PrivateWalkInRate = privateRates.WalkInRate,
        .EnsuiteRegularRate = ensuiteSurcharge.FinalRate,
        .EnsuiteWalkInRate = ensuiteRates.WalkInRate
    }
    End Function


    Private Function GetRoomTypeRates(roomType As String, dayPrefix As String, available As Integer) As (RegularRate As Double, WalkInRate As Double)
        Dim configKey As String = ""

        Select Case roomType.ToLower()
            Case "dorm"
                If available >= 8 Then
                    configKey = $"{roomType}{dayPrefix}8Plus"
                ElseIf available >= 4 Then
                    configKey = $"{roomType}{dayPrefix}4to7"
                ElseIf available >= 2 Then
                    configKey = $"{roomType}{dayPrefix}2to3"
                Else
                    configKey = $"{roomType}{dayPrefix}1"
                End If
            Case "private"
                If available >= 3 Then
                    configKey = $"{roomType}{dayPrefix}3Rooms"
                ElseIf available >= 2 Then
                    configKey = $"{roomType}{dayPrefix}2Rooms"
                Else
                    configKey = $"{roomType}{dayPrefix}1Room"
                End If
            Case "ensuite"
                If available >= 2 Then
                    configKey = $"{roomType}{dayPrefix}2Rooms"
                Else
                    configKey = $"{roomType}{dayPrefix}1Room"
                End If
        End Select

        ' Try Google Sheets first, fallback to App.config
        Try
            If googleSheetsService IsNot Nothing AndAlso googleSheetsService.IsGoogleSheetsEnabled() Then
                Dim googleRates = googleSheetsService.GetRateData()
                If googleRates.ContainsKey(configKey) Then
                    Return googleRates(configKey)
                Else
                    Console.WriteLine($"Rate {configKey} not found in Google Sheets, using App.config fallback")
                End If
            End If
        Catch ex As Exception
            Console.WriteLine($"Google Sheets error for {configKey}, using App.config fallback: {ex.Message}")
        End Try

        ' Fallback to appsettings.json
        Dim rateString = configuration($"FallbackRates:{configKey}")

        If Not String.IsNullOrEmpty(rateString) Then
            Dim rates = rateString.Split(","c)
            If rates.Length = 2 Then
                Console.WriteLine($"Using App.config fallback rate for {configKey}")
                Return (Double.Parse(rates(0)), Double.Parse(rates(1)))
            End If
        End If

        ' Default rates if nothing found
        Console.WriteLine($"Using default rates for {configKey}")
        Return (50, 40)
    End Function

    Private Function CalculateRateChanges(availabilityData As Dictionary(Of String, RoomAvailability), propertyData As LittleHotelierResponse) As List(Of RateChange)
        Dim changes As New List(Of RateChange)

        For Each kvp In availabilityData
            Dim dateStr = kvp.Key
            Dim availability = kvp.Value
            Dim daysAhead = DateDiff(DateInterval.Day, GetBusinessToday(), DateTime.Parse(dateStr))

            ' Only apply dynamic pricing for <15 days
            If daysAhead < 15 Then
                Dim calculatedRates = GetCurrentRates(availability, dateStr)
                Dim apiCurrentRates = GetApiRatesFromResponse(propertyData, dateStr)

                ' Check Dorm REGULAR rate changes only
                If calculatedRates.DormRegularRate <> apiCurrentRates.DormRegularRate Then
                    changes.Add(New RateChange With {
                        .CheckDate = dateStr,
                        .RoomType = GetAvailabilityDescription("Dorm", availability.DormBedsAvailable),
                        .OldRegularRate = apiCurrentRates.DormRegularRate,
                        .NewRegularRate = calculatedRates.DormRegularRate,
                        .OldWalkInRate = -1, ' Not used anymore
                        .NewWalkInRate = calculatedRates.DormWalkInRate, ' Show current walk-in rate
                        .AvailableUnits = availability.DormBedsAvailable,
                        .DaysAhead = daysAhead
                    })

                    LogPriceChangeToAnalytics(changes.Last(), availability.DormBedsAvailable, totalDormBeds, "Dorm")
                End If

                ' Check Private REGULAR rate changes only
                If calculatedRates.PrivateRegularRate <> apiCurrentRates.PrivateRegularRate Then
                    changes.Add(New RateChange With {
                        .CheckDate = dateStr,
                        .RoomType = GetAvailabilityDescription("Private", availability.PrivateRoomsAvailable),
                        .OldRegularRate = apiCurrentRates.PrivateRegularRate,
                        .NewRegularRate = calculatedRates.PrivateRegularRate,
                        .OldWalkInRate = -1, ' Not used anymore
                        .NewWalkInRate = calculatedRates.PrivateWalkInRate, ' Show current walk-in rate
                        .AvailableUnits = availability.PrivateRoomsAvailable,
                        .DaysAhead = daysAhead
                    })

                    LogPriceChangeToAnalytics(changes.Last(), availability.PrivateRoomsAvailable, totalPrivateRooms, "Private")
                End If

                ' Check Ensuite REGULAR rate changes only
                If calculatedRates.EnsuiteRegularRate <> apiCurrentRates.EnsuiteRegularRate Then
                    changes.Add(New RateChange With {
                        .CheckDate = dateStr,
                        .RoomType = GetAvailabilityDescription("Ensuite", availability.PrivateEnsuitesAvailable),
                        .OldRegularRate = apiCurrentRates.EnsuiteRegularRate,
                        .NewRegularRate = calculatedRates.EnsuiteRegularRate,
                        .OldWalkInRate = -1, ' Not used anymore
                        .NewWalkInRate = calculatedRates.EnsuiteWalkInRate, ' Show current walk-in rate
                        .AvailableUnits = availability.PrivateEnsuitesAvailable,
                        .DaysAhead = daysAhead
                    })

                    LogPriceChangeToAnalytics(changes.Last(), availability.PrivateEnsuitesAvailable, totalEnsuiteRooms, "Ensuite")
                End If
            End If
        Next

        Return changes
    End Function

    Private Function GetAvailabilityDescription(roomType As String, available As Integer) As String
        Select Case roomType.ToLower()
            Case "dorm"
                Return "Dorm Beds"
            Case "private"
                Return "Private Rooms - Shared Bath"
            Case "ensuite"
                Return "Queen Ensuite"
        End Select
        Return roomType
    End Function


    ' UPDATED: Rate change notifications with quiet period check
    Public Async Function SendEmailNotificationAsync(changes As List(Of RateChange)) As Task
        Try
            ' Check if we're in a quiet period
            If IsEmailQuietPeriod() Then
                Console.WriteLine($"📧 Suppressed rate change email during {GetQuietPeriodDescription()}")
                Console.WriteLine($"   Would have sent notification for {changes.Count} rate changes")
                Return
            End If

            Dim subject = "🏨 Red Inn Court - Rate Update"
            Dim body = BuildRateChangeEmailBody(changes)

            Await SendEmailAsync(subject, body, False)
            Console.WriteLine("✅ Email notification sent successfully")

        Catch ex As Exception
            Console.WriteLine($"❌ Error sending email notification: {ex.Message}")
        End Try
    End Function

    ' UPDATED: Error notifications - always send (critical)
    Public Async Function SendEmailErrorNotificationAsync(errorMessage As String) As Task
        Try
            ' Check if we're in a quiet period
            If IsEmailQuietPeriod() Then
                Console.WriteLine($"⚠️  Sending critical error email despite {GetQuietPeriodDescription()}")
            End If

            Dim subject = "🚨 Dynamic Pricing Error Alert"
            Dim body = $"<h2>🚨 DYNAMIC PRICING ERROR 🚨</h2>" &
                      $"<div style='padding: 15px; background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 5px; color: #721c24;'>" &
                      $"<p><strong>❌ Error:</strong> {errorMessage}</p>" &
                      $"<p><strong>⏰ Time:</strong> {DateTime.Now:yyyy-MM-dd HH:mm}</p>" &
                      $"</div>" &
                      $"<p>Please check the application logs for more details.</p>"

            Await SendEmailAsync(subject, body, True)
            Console.WriteLine("✅ Error email notification sent successfully")

        Catch ex As Exception
            Console.WriteLine($"❌ Error sending email error notification: {ex.Message}")
        End Try
    End Function

    Private Function BuildRateChangeEmailBody(changes As List(Of RateChange)) As String
        Dim body As New StringBuilder()
        body.AppendLine($"<h2>🏨 RED INN COURT - RATE UPDATE 🏨</h2>")
        body.AppendLine($"<p><strong>📅 {DateTime.Now:yyyy-MM-dd HH:mm}</strong></p>")
        body.AppendLine($"<br>")

        For Each change In changes
            Dim dayLabel = If(change.DaysAhead = 0, "Today", If(change.DaysAhead = 1, "Tomorrow (Day +1)", $"Day >+2 ({change.DaysAhead} days ahead)"))

            body.AppendLine($"<div style='margin-bottom: 20px; padding: 15px; border-left: 4px solid #007bff; background-color: #f8f9fa;'>")
            body.AppendLine($"<h3>📅 {change.CheckDate} ({dayLabel})</h3>")
            body.AppendLine($"<p><strong>🏠 Room Type:</strong> {change.RoomType}</p>")
            body.AppendLine($"<p><strong>🛏️ Available Units:</strong> {change.AvailableUnits}</p>")

            If change.OldRegularRate = 0 Then
                ' First time setting rates (API had 0)
                body.AppendLine($"<p><strong>💰 New Rates:</strong></p>")
                body.AppendLine($"<p>• Regular: <span style='color: #28a745; font-size: 1.1em;'>RM{change.NewRegularRate}</span></p>")
                body.AppendLine($"<p>• Walk-in: <span style='color: #6c757d; font-size: 1.0em;'>RM{change.NewWalkInRate}</span></p>")
            Else
                ' Rate change detected
                body.AppendLine($"<p><strong>💰 Rate Change:</strong></p>")
                body.AppendLine($"<p>• Regular: RM{change.OldRegularRate} → <span style='color: #28a745; font-size: 1.1em;'>RM{change.NewRegularRate}</span></p>")
                body.AppendLine($"<p>• Walk-in: <span style='color: #6c757d; font-size: 1.0em;'>RM{change.NewWalkInRate}</span> <em>(Calculated rate)</em></p>")
            End If
            body.AppendLine($"</div>")
        Next

        body.AppendLine($"<br><hr>")
        body.AppendLine($"<p><em>Generated by Dynamic Pricing Bot 🤖</em></p>")
        body.AppendLine($"<p style='font-size: 0.9em; color: #6c757d;'><em>Note: Rates sourced from Google Sheets with App.config fallback. Walk-in rates calculated automatically.</em></p>")

        Return body.ToString()
    End Function

    Private Async Function SendEmailAsync(subject As String, body As String, isError As Boolean) As Task
        Try
            Using smtpClient As New SmtpClient(smtpHost, smtpPort)
                smtpClient.EnableSsl = True
                smtpClient.Credentials = New NetworkCredential(smtpUsername, smtpPassword)

                Using message As New MailMessage()
                    message.From = New MailAddress(emailFromAddress, emailFromName)
                    message.To.Add(emailToAddress)
                    message.Subject = subject
                    message.Body = body
                    message.IsBodyHtml = True

                    If isError Then
                        message.Priority = MailPriority.High
                    End If

                    Await smtpClient.SendMailAsync(message)
                End Using
            End Using

        Catch ex As Exception
            Console.WriteLine($"Error sending email: {ex.Message}")
            Throw
        End Try
    End Function

    ' UPDATED: Log price changes to analytics with quiet period check
    Private Sub LogPriceChangeToAnalytics(change As RateChange, available As Integer, total As Integer, category As String)
        If googleSheetsService Is Nothing Then Return

        Try
            ' Check if we're in a quiet period - suppress analytics logging as well
            If IsEmailQuietPeriod() Then
                Console.WriteLine($"📊 Suppressed {category} analytics logging during {GetQuietPeriodDescription()}")
                Console.WriteLine($"   Would have logged rate change: {change.OldRegularRate} → {change.NewRegularRate}")
                Return
            End If

            Dim entry As New PriceHistoryEntry With {
            .Timestamp = GetBusinessDateTime(),
            .CheckInDate = change.CheckDate,
            .RoomType = change.RoomType,
            .RoomCategory = category,
            .AvailableUnits = available,
            .TotalCapacity = total,
            .OccupancyRate = If(total > 0, CDbl(total - available) / total, 0),
            .DaysAhead = change.DaysAhead,
            .OldRate = change.OldRegularRate,
            .NewRate = change.NewRegularRate,
            .RateChange = change.NewRegularRate - change.OldRegularRate,
            .PercentChange = If(change.OldRegularRate > 0, (change.NewRegularRate - change.OldRegularRate) / change.OldRegularRate, 0),
            .PriceTier = GetPriceTier(category, available),
            .DayType = If(change.DaysAhead = 0, "Today", If(change.DaysAhead = 1, "Day+1", "Day>+2")),
            .BusinessHour = GetBusinessDateTime().Hour,
            .ConfigSource = If(googleSheetsService.IsGoogleSheetsEnabled(), "GoogleSheets", "AppConfig")
        }

            Dim success = googleSheetsService.LogPriceChangeToAnalytics(entry)
            If success Then
                Console.WriteLine($"✓ Logged {category} rate change to analytics")
            End If

        Catch ex As Exception
            Console.WriteLine($"Failed to log price change to analytics: {ex.Message}")
        End Try
    End Sub


    Private Function GetPriceTier(category As String, available As Integer) As String
        Select Case category.ToLower()
            Case "dorm"
                If available >= 8 Then Return "8Plus"
                If available >= 4 Then Return "4to7"
                If available >= 2 Then Return "2to3"
                Return "1"
            Case "private"
                If available >= 3 Then Return "3Rooms"
                If available >= 2 Then Return "2Rooms"
                Return "1Room"
            Case "ensuite"
                If available >= 2 Then Return "2Rooms"
                Return "1Room"
            Case Else
                Return "Unknown"
        End Select
    End Function

    ' NEW: Calculate surcharges based on Google Sheets config
    Private Function CalculateRateWithSurcharges(baseRate As Double, targetDate As String) As SurchargeResult
        Dim result As New SurchargeResult With {
        .BaseRate = baseRate,
        .WeekendSurcharge = 0,
        .PublicHolidaySurcharge = 0,
        .FinalRate = baseRate
    }

        Try
            If googleSheetsService Is Nothing OrElse Not googleSheetsService.IsGoogleSheetsEnabled() Then
                Return result
            End If

            ' Check if surcharges are enabled
            Dim surchargesEnabled = String.Equals(configuration("GoogleSheets:EnableSurcharges"), "true", StringComparison.OrdinalIgnoreCase)
            If Not surchargesEnabled Then
                Return result
            End If

            Dim checkDate = DateTime.Parse(targetDate)
            Dim surchargeConfigs = googleSheetsService.GetSurchargeConfig()
            Dim publicHolidays = googleSheetsService.GetPublicHolidays()

            ' Check for weekend surcharge (Friday = 5, Saturday = 6)
            If (checkDate.DayOfWeek = DayOfWeek.Friday OrElse checkDate.DayOfWeek = DayOfWeek.Saturday) AndAlso
           surchargeConfigs.ContainsKey("Weekend") AndAlso surchargeConfigs("Weekend").Enabled Then

                result.WeekendSurcharge = CalculateSurchargeAmount(baseRate, surchargeConfigs("Weekend"))
                result.AppliedSurcharges.Add($"Weekend ({surchargeConfigs("Weekend").Amount}%)")
            End If

            ' Check for public holiday surcharge
            Dim isPublicHoliday = publicHolidays.Any(Function(h) h.Enabled AndAlso h.HolidayDate.Date = checkDate.Date)
            If isPublicHoliday AndAlso surchargeConfigs.ContainsKey("PublicHoliday") AndAlso surchargeConfigs("PublicHoliday").Enabled Then
                result.PublicHolidaySurcharge = CalculateSurchargeAmount(baseRate, surchargeConfigs("PublicHoliday"))
                Dim holiday = publicHolidays.First(Function(h) h.HolidayDate.Date = checkDate.Date)
                result.AppliedSurcharges.Add($"Public Holiday - {holiday.Description} ({surchargeConfigs("PublicHoliday").Amount}%)")
            End If

            ' Calculate final rate
            result.FinalRate = result.BaseRate + result.WeekendSurcharge + result.PublicHolidaySurcharge

        Catch ex As Exception
            Console.WriteLine($"Error calculating surcharges for {targetDate}: {ex.Message}")
        End Try

        Return result
    End Function

    Private Function CalculateSurchargeAmount(baseRate As Double, surchargeConfig As SurchargeConfig) As Double
        Select Case surchargeConfig.Type.ToLower()
            Case "percentage"
                Return baseRate * (surchargeConfig.Amount / 100)
            Case "fixed"
                Return surchargeConfig.Amount
            Case Else
                Return 0
        End Select
    End Function

    Private Function IsWeekendOrHoliday(targetDate As String) As String
        Try
            Dim checkDate = DateTime.Parse(targetDate)
            Dim publicHolidays = If(googleSheetsService?.GetPublicHolidays(), New List(Of PublicHoliday))

            Dim isWeekend = checkDate.DayOfWeek = DayOfWeek.Friday OrElse checkDate.DayOfWeek = DayOfWeek.Saturday
            Dim isHoliday = publicHolidays.Any(Function(h) h.Enabled AndAlso h.HolidayDate.Date = checkDate.Date)

            If isHoliday AndAlso isWeekend Then
                Return "Weekend + Holiday"
            ElseIf isHoliday Then
                Return "Public Holiday"
            ElseIf isWeekend Then
                Return "Weekend"
            Else
                Return "Regular"
            End If
        Catch
            Return "Regular"
        End Try
    End Function


End Class
