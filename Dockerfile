# Use official .NET 8 runtime image
FROM mcr.microsoft.com/dotnet/runtime:8.0

# Set working directory
WORKDIR /app

# Copy published application files
COPY bin/Release/net8.0/publish/ ./

# Create directories for external dependencies
RUN mkdir -p /app/credentials /app/logs

# CRITICAL FIX: Install timezone data package
RUN apt-get update && apt-get install -y tzdata && rm -rf /var/lib/apt/lists/*

# Set timezone to Singapore (your business timezone)
ENV TZ=Asia/Singapore
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone

# Create non-root user for security
RUN adduser --disabled-password --gecos '' appuser && \
    chown -R appuser:appuser /app
USER appuser

# Set entry point
ENTRYPOINT ["dotnet", "RedInnDynamicPricingLinux.dll"]
