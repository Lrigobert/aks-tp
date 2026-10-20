// =============================================================
//  PROJET AKS - TP Azure / Kubernetes
//  Auteur  : TP étudiant
//  Région  : francecentral
//  Fichier : main.bicep
// =============================================================

// ---------------------------------------------------------------
// PARAMÈTRES — valeurs injectées depuis main.parameters.json
// ---------------------------------------------------------------

@description('Nom du cluster AKS')
param clusterName string = 'aks-tp-cluster'

@description('Région Azure pour toutes les ressources')
param location string = 'francecentral'

@description('Taille de VM du pool de nœuds (B2s = 2 vCPU, 4 Go RAM)')
param nodeVmSize string = 'Standard_B2s'

@description('Nombre de nœuds dans le pool (1 suffit pour un TP)')
@minValue(1)
@maxValue(3)
param nodeCount int = 1

@description('Taille du disque OS en Go (30 Go minimum recommandé)')
param osDiskSizeGB int = 30

@description('Nom du workspace Log Analytics')
param logAnalyticsWorkspaceName string = 'law-aks-tp'

@description('SKU du workspace Log Analytics (PerGB2018 = facturation à l usage)')
param logAnalyticsSku string = 'PerGB2018'

@description('Rétention des logs en jours (30 jours = gratuit dans la limite du free tier)')
@minValue(30)
@maxValue(90)
param logRetentionDays int = 30

// ---------------------------------------------------------------
// RESSOURCE 1 : LOG ANALYTICS WORKSPACE
// Rôle : stocke tous les logs et métriques du cluster AKS
// Container Insights y envoie les données collectées par l agent OMS
// ---------------------------------------------------------------
resource logAnalyticsWorkspace 'Microsoft.OperationalInsights/workspaces@2022-10-01' = {
  name: logAnalyticsWorkspaceName
  location: location
  properties: {
    sku: {
      name: logAnalyticsSku
    }
    retentionInDays: logRetentionDays
    features: {
      enableLogAccessUsingOnlyResourcePermissions: true
    }
    publicNetworkAccessForIngestion: 'Enabled'
    publicNetworkAccessForQuery: 'Enabled'
  }
}

// ---------------------------------------------------------------
// RESSOURCE 2 : CLUSTER AKS
// Rôle : cluster Kubernetes managé avec 1 nœud Standard_B2s
// - Identity SystemAssigned : Azure gère automatiquement le service principal
// - addonProfiles.omsagent : active Container Insights nativement
// - networkPlugin kubenet : réseau simple, adapté aux abonnements étudiants
// ---------------------------------------------------------------
resource aksCluster 'Microsoft.ContainerService/managedClusters@2023-07-01' = {
  name: clusterName
  location: location

  // Identité managée par le système — pas besoin de créer un SP manuellement
  identity: {
    type: 'SystemAssigned'
  }

  properties: {
    // Préfixe DNS pour l API server du cluster (https://<dnsPrefix>.hcp.francecentral.azmk8s.io)
    dnsPrefix: clusterName

    // Version Kubernetes (vide = dernière version stable recommandée par Azure)

    // ----- POOL DE NŒUDS SYSTÈME -----
    agentPoolProfiles: [
      {
        name: 'agentpool'         // Nom du pool (lettres minuscules uniquement)
        count: nodeCount          // 1 nœud pour économiser les crédits étudiants
        vmSize: nodeVmSize        // Standard_B2s : burstable, économique
        osType: 'Linux'           // Linux requis pour AKS
        osDiskSizeGB: osDiskSizeGB
        mode: 'System'            // Pool système = indispensable, héberge les pods système
        type: 'VirtualMachineScaleSets'  // VMSS obligatoire pour les clusters AKS modernes

        // Active l agent OMS directement sur chaque nœud du pool
        // L agent collecte : logs des conteneurs, métriques CPU/RAM, événements K8s
        nodeLabels: {
          environment: 'tp'
          managedby: 'bicep'
        }
      }
    ]

    // ----- PROFIL RÉSEAU -----
    networkProfile: {
      networkPlugin: 'kubenet'        // Réseau simple basé sur des routes IP
      loadBalancerSku: 'standard'     // Load Balancer Standard requis pour AKS en production
      outboundType: 'loadBalancer'    // Trafic sortant via le Load Balancer
    }

    // ----- ADDONS — Container Insights -----
    // omsagent déploie automatiquement un DaemonSet sur chaque nœud
    // Il collecte : stdout/stderr des conteneurs, métriques système, inventaire des pods
    addonProfiles: {
      omsagent: {
        enabled: true
        config: {
          // ID du workspace où les données seront envoyées
          logAnalyticsWorkspaceResourceID: logAnalyticsWorkspace.id
        }
      }
    }

    // ----- PROFIL AUTO-SCALER (désactivé pour économiser les crédits) -----
    // enableRBAC : activé par défaut sur les clusters AKS modernes
    enableRBAC: true

    // Désactive les features payantes non nécessaires pour le TP
    disableLocalAccounts: false
    publicNetworkAccess: 'Enabled'
  }
}

// ---------------------------------------------------------------
// OUTPUTS — valeurs retournées après le déploiement
// Utiles pour les scripts de post-déploiement et vérifications
// ---------------------------------------------------------------

@description('Nom du cluster AKS déployé')
output clusterName string = aksCluster.name

@description('ID de ressource complet du cluster AKS')
output clusterResourceId string = aksCluster.id

@description('ID du workspace Log Analytics')
output logAnalyticsWorkspaceId string = logAnalyticsWorkspace.id

@description('Nom du workspace Log Analytics')
output logAnalyticsWorkspaceName string = logAnalyticsWorkspace.name

@description('Object ID de l identité managée du cluster (pour les assignations RBAC)')
output clusterPrincipalId string = aksCluster.identity.principalId

@description('FQDN de l API server Kubernetes')
output clusterFqdn string = aksCluster.properties.fqdn
