// Azure Bicep Infrastructure Deployment for ReadMeeting.ai (DGF)
// Region: West Europe / France Central
targetScope = 'resourceGroup'

@description('Environnement de déploiement (dev, test, prod)')
@allowed([
  'dev'
  'test'
  'prod'
])
param environmentName string = 'prod'

@description('Localisation des ressources Azure')
param location string = resourceGroup().location

@description('Nom de base des ressources pour la DGF')
param baseName string = 'dgf-readmeeting'

var uniqueSuffix = uniqueString(resourceGroup().id)
var sqlServerName = '${baseName}-sql-${uniqueSuffix}'
var sqlDbName = 'ReadMeetingDb'
var appServiceName = '${baseName}-api-${environmentName}'
var appServicePlanName = '${baseName}-plan-${environmentName}'
var searchServiceName = '${baseName}-search-${uniqueSuffix}'
var serviceBusNamespaceName = '${baseName}-sb-${uniqueSuffix}'

// 1. Azure App Service Plan (Linux)
resource appServicePlan 'Microsoft.Web/serverfarms@2022-09-01' = {
  name: appServicePlanName
  location: location
  sku: {
    name: 'P1v3'
    tier: 'PremiumV3'
  }
  properties: {
    reserved: true // Linux
  }
}

// 2. Azure App Service (Backend .NET 8 / 9 Web API)
resource appService 'Microsoft.Web/sites@2022-09-01' = {
  name: appServiceName
  location: location
  properties: {
    serverFarmId: appServicePlan.id
    siteConfig: {
      linuxFxVersion: 'DOTNETCORE|8.0'
      alwaysOn: true
      http20Enabled: true
      minTlsVersion: '1.2'
    }
    httpsOnly: true
  }
  identity: {
    type: 'SystemAssigned'
  }
}

// 3. Azure SQL Database (ACID & Sécurité des données institutionnelles)
resource sqlServer 'Microsoft.Sql/servers@2022-11-01-preview' = {
  name: sqlServerName
  location: location
  properties: {
    version: '12.0'
    minimalTlsVersion: '1.2'
    publicNetworkAccess: 'Enabled'
    administrators: {
      administratorType: 'ActiveDirectory'
      azureADOnlyAuthentication: true
    }
  }
}

resource sqlDatabase 'Microsoft.Sql/servers/databases@2022-11-01-preview' = {
  parent: sqlServer
  name: sqlDbName
  location: location
  sku: {
    name: 'GP_Gen5'
    tier: 'GeneralPurpose'
    capacity: 2
  }
}

// 4. Azure AI Search (Indexation vectorielle des réunions & comptes rendus)
resource searchService 'Microsoft.Search/searchServices@2023-11-01' = {
  name: searchServiceName
  location: location
  sku: {
    name: 'standard'
  }
  properties: {
    replicaCount: 1
    partitionCount: 1
    semanticSearch: 'free'
  }
}

// 5. Azure Service Bus (Traitement asynchrone des transcriptions volumineuses)
resource serviceBusNamespace 'Microsoft.ServiceBus/namespaces@2022-10-01-preview' = {
  name: serviceBusNamespaceName
  location: location
  sku: {
    name: 'Standard'
    tier: 'Standard'
  }
}

output appServiceEndpoint string = 'https://${appService.properties.defaultHostName}'
output sqlServerFqdn string = sqlServer.properties.fullyQualifiedDomainName
output searchEndpoint string = 'https://${searchService.name}.search.windows.net'
