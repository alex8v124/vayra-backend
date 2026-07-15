# stop-monitoring.ps1
Write-Host "Deteniendo contenedores de monitoreo Podman..." -ForegroundColor Cyan
podman rm -f xplora-prometheus 2>$null
podman rm -f xplora-grafana 2>$null
Write-Host "Contenedores de Prometheus y Grafana detenidos y eliminados." -ForegroundColor Green
