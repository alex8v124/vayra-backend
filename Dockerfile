# Etapa 1: Build con imagen oficial de Maven + Java 21
FROM maven:3.9.6-eclipse-temurin-21-alpine AS build
WORKDIR /app
COPY pom.xml .
COPY src ./src
RUN mvn clean package -DskipTests

# Etapa 2: Runtime con Java 21 + Python3 + Pandas + Excel libraries para Storechecks
FROM eclipse-temurin:21-jre-alpine
WORKDIR /app

# Instalar Python3 y paquetes optimizados Alpine (pandas, numpy, openpyxl) + xlsxwriter
RUN apk add --no-cache python3 py3-pip py3-pandas py3-numpy py3-openpyxl && \
    ln -sf /usr/bin/python3 /usr/bin/python && \
    pip3 install --break-system-packages --no-cache-dir xlsxwriter

COPY --from=build /app/target/*.jar app.jar
COPY scripts /app/scripts
COPY headless_storecheck.py /app/headless_storecheck.py

# Configuración por defecto de OpenTelemetry / Grafana Cloud
ENV OTEL_EXPORTER_OTLP_ENDPOINT="https://otlp-gateway-prod-us-east-3.grafana.net/otlp"
ENV OTEL_EXPORTER_OTLP_HEADERS="Authorization=Basic%20MTcyNDY3MTpnbGNfZXlKdklqb2lNVGcwTlRrNU55SXNJbTRpT2lKMGIyTnJaVzR0ZG1GNWNtRWlMQ0pySWpvaU1WRkVWRmt4UldvMmFHUXdNekpNWldRMmRtbEpPVEl4SWl3aWJTSTZleUp5SWpvaWNISnZaQzExY3kxbFlYTjBMVE1pZlgwPQ=="
ENV OTEL_SERVICE_NAME="xplora-backend"

EXPOSE 8080
ENTRYPOINT ["java", "-jar", "app.jar"]
