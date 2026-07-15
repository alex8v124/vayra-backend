# start-monitoring.ps1
Write-Host "Iniciando monitoreo Xplora (Prometheus & Grafana) con Podman..." -ForegroundColor Cyan

# 1. Crear red de podman si no existe
podman network exists xplora-net
if ($LASTEXITCODE -ne 0) {
    podman network create xplora-net
}

# 2. Detener y eliminar contenedores previos si existen
podman rm -f xplora-prometheus 2>$null
podman rm -f xplora-grafana 2>$null

# 3. Obtener ruta absoluta (convertir a formato POSIX para montajes en Podman WSL/Linux si es necesario)
$pwdPath = (Get-Item .).FullName.Replace('\', '/')
if ($pwdPath -match '^([a-zA-Z]):/(.*)') {
    $drive = $matches[1].ToLower()
    $rest = $matches[2]
    $pwdMount = "/mnt/$drive/$rest"
} else {
    $pwdMount = $pwdPath
}

Write-Host "Ruta de montaje detectada: $pwdMount" -ForegroundColor Gray

# 4. Levantar Prometheus
podman run -d --name xplora-prometheus --network xplora-net `
    -p 9090:9090 `
    -v "${pwdMount}/prometheus/prometheus.yml:/etc/prometheus/prometheus.yml:ro" `
    --add-host host.containers.internal:host-gateway `
    --add-host host.docker.internal:host-gateway `
    docker.io/prom/prometheus:latest `
    --config.file=/etc/prometheus/prometheus.yml `
    --storage.tsdb.path=/prometheus

# 5. Levantar Grafana
podman run -d --name xplora-grafana --network xplora-net `
    -p 3000:3000 `
    -e GF_SECURITY_ADMIN_USER=admin `
    -e GF_SECURITY_ADMIN_PASSWORD=admin `
    -e GF_USERS_ALLOW_SIGN_UP=false `
    -v "${pwdMount}/grafana/provisioning:/etc/grafana/provisioning:ro" `
    docker.io/grafana/grafana:latest

Write-Host "`n¡Contenedores en ejecución exitosamente con Podman!" -ForegroundColor Green
Write-Host " -> Prometheus UI: http://localhost:9090" -ForegroundColor Yellow
Write-Host " -> Grafana UI:    http://localhost:3000 (Usuario: admin | Clave: admin)" -ForegroundColor Yellow
