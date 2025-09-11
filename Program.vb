Imports System
Imports Microsoft.Extensions.Configuration

Module Program
    Sub Main(args As String())
        Try
            Console.WriteLine("=== Red Inn Court Dynamic Pricing Service (.NET 8) ===")
            Console.WriteLine($"Started at: {DateTime.Now}")
            Console.WriteLine()

            ' Build configuration
            Dim configuration = New ConfigurationBuilder() _
                .SetBasePath(AppContext.BaseDirectory) _
                .AddJsonFile("appsettings.json", optional:=False) _
                .Build()

            ' Test configuration reading
            Console.WriteLine("✅ Configuration loaded successfully")
            Console.WriteLine($"Property ID: {configuration("LittleHotelier:PropertyId")}")
            Console.WriteLine($"Email enabled: {Not String.IsNullOrEmpty(configuration("Email:SmtpHost"))}")
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
