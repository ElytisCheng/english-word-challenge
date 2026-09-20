targetScope = 'subscription'

@description('Azure region for the resource group and Container Apps environment.')
param location string = 'eastus2'

@description('Resource group name.')
param resourceGroupName string = 'rg-english-word-challenge'

@description('Container Apps environment name.')
param environmentName string = 'cae-english-word-challenge'

@description('Container App name.')
param appName string = 'ca-english-word-challenge'

@description('Public container image to deploy.')
param image string = 'ghcr.io/elytischeng/english-word-challenge:latest'

resource resourceGroup 'Microsoft.Resources/resourceGroups@2024-03-01' = {
  name: resourceGroupName
  location: location
}

module containerApp 'container-app.bicep' = {
  name: 'container-app'
  scope: resourceGroup
  params: {
    location: location
    environmentName: environmentName
    appName: appName
    image: image
  }
}

output appUrl string = containerApp.outputs.appUrl

