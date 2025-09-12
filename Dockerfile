# Use official .NET 8 runtime image
FROM mcr.microsoft.com/dotnet/runtime:8.0

# Set working directory
WORKDIR /app

# Copy published application files
COPY bin/Release/net8.0/publish/ ./

# CRITICAL: Install timezone data FIRST
RUN apt-get update && apt-get install -y tzdata && rm -rf /var/lib/apt/lists/*

# Set timezone properly with explicit export
ENV TZ=Asia/Singapore
RUN ln -snf /usr/share/zoneinfo/Asia/Singapore /etc/localtime && \
    echo "Asia/Singapore" > /etc/timezone

# Verify timezone setup during build
RUN date && echo "Timezone set to: $(cat /etc/timezone)"

# Create directories and user
RUN mkdir -p /app/credentials /app/logs
RUN adduser --disabled-password --gecos '' appuser && \
    chown -R appuser:appuser /app

# Switch to non-root user
USER appuser

# Ensure TZ is available to the application
ENV TZ=Asia/Singapore

ENTRYPOINT ["dotnet", "RedInnDynamicPricingLinux.dll"]
