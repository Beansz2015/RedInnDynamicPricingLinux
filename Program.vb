Imports System
Imports Microsoft.Extensions.Configuration

Module Program
    Sub Main(args As String())
        Try
            Console.WriteLine("=== Red Inn Court Dynamic Pricing Service (.NET 8) ===")
            Console.WriteLine($"Started at: {DateTime.Now}")
            Console.WriteLine()

            ' Read environment variable - FIXED: Use different variable name
            Dim env As String = System.Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT")
            If String.IsNullOrEmpty(env) Then env = "Production"

            ' Build configuration with environment-specific support - FIXED
            Dim builder = New ConfigurationBuilder() _
    .SetBasePath(AppContext.BaseDirectory) _
    .AddJsonFile("appsettings.json", optional:=False, reloadOnChange:=True) _
    .AddJsonFile($"appsettings.{env}.json", optional:=True, reloadOnChange:=True) _
    .AddEnvironmentVariables()

            Dim configuration = builder.Build()

            ' Debug output to verify config loading - IMPORTANT
            Console.WriteLine("=== CONFIGURATION DEBUG ===")
            Console.WriteLine($"Environment: {env}")
            Console.WriteLine("Loaded configuration files:")
            Console.WriteLine(" ✓ appsettings.json")
            Console.WriteLine($" ✓ appsettings.{env}.json")

            ' CRITICAL: Log the actual credentials path being used
            Dim credentialsPath = configuration("GoogleSheets:CredentialsPath")
            Console.WriteLine($"Google Credentials Path: {credentialsPath}")
            Console.WriteLine("=== END CONFIG DEBUG ===")
            Console.WriteLine()

            ' Test configuration reading
            Console.WriteLine("✅ Configuration loaded successfully")
            Console.WriteLine($"Property ID: {configuration("LittleHotelier:PropertyId")}")
            Console.WriteLine($"Email enabled: {Not String.IsNullOrEmpty(configuration("Email:SmtpHost"))}")

            Console.WriteLine()
            Console.WriteLine("=== TIMEZONE DEBUG ===")
            Console.WriteLine($"🕐 Local timezone: {TimeZoneInfo.Local.DisplayName}")
            Console.WriteLine($"🕐 Local time: {DateTime.Now}")
            Console.WriteLine($"🕐 UTC time: {DateTime.UtcNow}")
            Console.WriteLine($"🕐 Environment TZ: {Environment.GetEnvironmentVariable("TZ")}")

            ' Test your GetBusinessDateTime method
            Try
                Dim pricingServiceTest As New DynamicPricingService(configuration)
                Dim businessTime = pricingServiceTest.GetBusinessDateTime() ' You'll need to make this method public temporarily
                Console.WriteLine($"🕐 Business time: {businessTime}")
            Catch
                Console.WriteLine($"🕐 Business time: Cannot test (method private)")
            End Try

            Console.WriteLine("=== END TIMEZONE DEBUG ===")
            Console.WriteLine()
            Console.WriteLine()

            ' Create and test the main service
            Console.WriteLine("Testing complete DynamicPricingService...")
            Try
                Dim pricingService As New DynamicPricingService(configuration)
                Console.WriteLine("✅ DynamicPricingService created successfully!")

                ' Test the complete pricing check workflow
                Console.WriteLine("🔄 Running dynamic pricing check...")
                Dim task = pricingService.RunDynamicPricingCheck()
                task.Wait()

                Console.WriteLine("✅ Dynamic pricing check completed!")

            Catch ex As Exception
                Console.WriteLine($"❌ DynamicPricingService error: {ex.Message}")
                Console.WriteLine($"Stack trace: {ex.StackTrace}")
            End Try

            Console.WriteLine()
            Console.WriteLine("✅ Day 3 complete migration test finished!")
            Console.WriteLine("Press any key to exit...")
            Console.ReadKey()

        Catch ex As Exception
            Console.WriteLine($"❌ Fatal error: {ex.Message}")
            Console.WriteLine($"Stack trace: {ex.StackTrace}")
            Console.WriteLine("Press any key to exit...")
            Console.ReadKey()
        End Try
    End Sub
End Module
