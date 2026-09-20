# English Word Challenge

Static English vocabulary practice app served by Nginx and deployed to Azure Container Apps.

## Cost profile

- Azure Container Apps Consumption
- `minReplicas: 0` and `maxReplicas: 1`
- 0.25 vCPU and 0.5 GiB memory
- No Log Analytics workspace
- Public image in GitHub Container Registry, so no Azure Container Registry is required

## Deploy

The GitHub Actions workflow builds `ghcr.io/elytischeng/english-word-challenge:latest`.
After the package is public, deploy the infrastructure with Azure CLI:

```powershell
az deployment sub create `
  --name english-word-challenge `
  --location eastus2 `
  --template-file .\infra\main.bicep `
  --parameters location=eastus2
```

The deployment output contains the HTTPS application URL.
