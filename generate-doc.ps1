# =============================================================
#  GÉNÉRATEUR DE DOCUMENT WORD - TP AKS Azure
#  Exécuter : cd C:\Users\ldard\aks-tp ; .\generate-doc.ps1
# =============================================================

$OutputPath = "C:\Users\ldard\aks-tp\TP_AKS_Azure_Kubernetes.docx"

try {
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    $Doc = $Word.Documents.Add()
    $Selection = $Word.Selection

    # ---- Styles helper ----
    function Add-Heading1($text) {
        $Selection.Style = $Doc.Styles["Titre 1"]
        $Selection.TypeText($text)
        $Selection.TypeParagraph()
    }
    function Add-Heading2($text) {
        $Selection.Style = $Doc.Styles["Titre 2"]
        $Selection.TypeText($text)
        $Selection.TypeParagraph()
    }
    function Add-Heading3($text) {
        $Selection.Style = $Doc.Styles["Titre 3"]
        $Selection.TypeText($text)
        $Selection.TypeParagraph()
    }
    function Add-Normal($text) {
        $Selection.Style = $Doc.Styles["Normal"]
        $Selection.TypeText($text)
        $Selection.TypeParagraph()
    }
    function Add-Code($text) {
        $Selection.Style = $Doc.Styles["Normal"]
        $Selection.Font.Name = "Courier New"
        $Selection.Font.Size = 9
        $Selection.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorDarkBlue
        $para = $Selection.Paragraphs.Add()
        $para.Range.Shading.BackgroundPatternColor = 15921906  # Light gray
        $Selection.TypeText($text)
        $Selection.TypeParagraph()
        $Selection.Font.Name = "Calibri"
        $Selection.Font.Size = 11
        $Selection.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorAutomatic
    }
    function Add-PageBreak() {
        $Selection.InsertBreak(7)  # wdPageBreak
    }

    # ================================================================
    # PAGE DE GARDE
    # ================================================================
    $Selection.Style = $Doc.Styles["Normal"]
    $Selection.ParagraphFormat.Alignment = 1  # Center

    # Espacement haut
    for ($i = 0; $i -lt 6; $i++) { $Selection.TypeParagraph() }

    $Selection.Font.Size = 28
    $Selection.Font.Bold = $true
    $Selection.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorDarkBlue
    $Selection.TypeText("Déploiement d'un Cluster AKS")
    $Selection.TypeParagraph()
    $Selection.TypeText("sur Microsoft Azure")
    $Selection.TypeParagraph()
    $Selection.TypeParagraph()

    $Selection.Font.Size = 18
    $Selection.Font.Bold = $false
    $Selection.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorBlue
    $Selection.TypeText("Infrastructure as Code avec Bicep")
    $Selection.TypeParagraph()
    $Selection.TypeText("& Monitoring Container Insights")
    $Selection.TypeParagraph()
    $Selection.TypeParagraph()

    $Selection.Font.Size = 14
    $Selection.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorBlack
    $Selection.TypeText("Technologies : Azure AKS • Bicep • Azure Monitor • Kubernetes")
    $Selection.TypeParagraph()
    $Selection.TypeParagraph()

    $Selection.Font.Size = 13
    $Selection.Font.Bold = $true
    $Selection.TypeText("Date : 25 Avril 2026")
    $Selection.TypeParagraph()
    $Selection.TypeText("Resource Group : rg-aks-tp")
    $Selection.TypeParagraph()
    $Selection.TypeText("Région : francecentral (Paris)")
    $Selection.TypeParagraph()

    $Selection.Font.Bold = $false
    $Selection.ParagraphFormat.Alignment = 0  # Left
    Add-PageBreak

    # ================================================================
    # ÉTAPE 1 — PRÉSENTATION DU PROJET
    # ================================================================
    $Selection.Style = $Doc.Styles["Titre 1"]
    $Selection.TypeText("Étape 1 — Présentation du Projet")
    $Selection.TypeParagraph()

    Add-Heading2 "Objectif Global"
    Add-Normal "Ce projet consiste à déployer un cluster Kubernetes managé (AKS — Azure Kubernetes Service) sur Microsoft Azure en utilisant l'Infrastructure as Code (IaC) via le langage Bicep. L'ensemble de l'infrastructure est défini dans des fichiers déclaratifs, versionnables et reproductibles."
    Add-Normal "Le monitoring complet du cluster est assuré par Azure Monitor et Container Insights, qui collectent en temps réel les métriques de performance, les logs des conteneurs et les événements Kubernetes."

    Add-Heading2 "Rôle de chaque Composant"
    Add-Normal "AKS (Azure Kubernetes Service) : Cluster Kubernetes entièrement managé par Azure. Azure gère le control plane (API server, etcd, scheduler), l'utilisateur gère uniquement les nodes (VM Workers)."
    Add-Normal "Bicep : Langage d'Infrastructure as Code natif Azure. Il compile en ARM JSON et permet de décrire l'infrastructure de manière déclarative, lisible et maintenable."
    Add-Normal "Log Analytics Workspace : Base de données centralisée qui stocke tous les logs, métriques et événements collectés par Azure Monitor."
    Add-Normal "Azure Monitor + Container Insights : Solution de monitoring spécialisée Kubernetes qui collecte les métriques CPU/RAM des pods et nodes, les logs stdout/stderr des conteneurs, et les événements du cluster."
    Add-Normal "OMS Agent : DaemonSet Kubernetes (un pod par node) déployé automatiquement par AKS pour collecter et transmettre les données au Log Analytics Workspace."

    Add-Heading2 "Workflow Complet"
    Add-Code "1. Écriture du fichier Bicep (infrastructure déclarée en code)"
    Add-Code "2. az deployment group create → Bicep compilé en ARM et envoyé à Azure"
    Add-Code "3. Azure provisionne : Log Analytics Workspace + Cluster AKS"
    Add-Code "4. AKS déploie automatiquement l'OMS Agent (DaemonSet) sur chaque node"
    Add-Code "5. L'OMS Agent collecte logs et métriques → envoie vers Log Analytics"
    Add-Code "6. Container Insights affiche les dashboards dans le portail Azure"
    Add-Code "7. Des alertes peuvent être configurées sur les métriques collectées"

    Add-PageBreak

    # ================================================================
    # ÉTAPE 2 — FICHIER MAIN.BICEP
    # ================================================================
    $Selection.Style = $Doc.Styles["Titre 1"]
    $Selection.TypeText("Étape 2 — Fichier main.bicep")
    $Selection.TypeParagraph()

    Add-Heading2 "Code complet"

    $bicepCode = @"
// =============================================================
//  PROJET AKS - TP Azure / Kubernetes  |  Région : francecentral
// =============================================================

@description('Nom du cluster AKS')
param clusterName string = 'aks-tp-cluster'

@description('Région Azure pour toutes les ressources')
param location string = 'francecentral'

@description('Taille de VM du pool de noeuds (B2s = 2 vCPU, 4 Go RAM)')
param nodeVmSize string = 'Standard_B2s'

@description('Nombre de noeuds dans le pool')
@minValue(1)
@maxValue(3)
param nodeCount int = 1

@description('Taille du disque OS en Go')
param osDiskSizeGB int = 30

@description('Nom du workspace Log Analytics')
param logAnalyticsWorkspaceName string = 'law-aks-tp'

@description('SKU du workspace Log Analytics')
param logAnalyticsSku string = 'PerGB2018'

@description('Retention des logs en jours')
@minValue(30)
@maxValue(90)
param logRetentionDays int = 30

@description('Version de Kubernetes (vide = version stable par defaut)')
param kubernetesVersion string = ''

// RESSOURCE 1 : LOG ANALYTICS WORKSPACE
resource logAnalyticsWorkspace 'Microsoft.OperationalInsights/workspaces@2022-10-01' = {
  name: logAnalyticsWorkspaceName
  location: location
  properties: {
    sku: { name: logAnalyticsSku }
    retentionInDays: logRetentionDays
    features: { enableLogAccessUsingOnlyResourcePermissions: true }
    publicNetworkAccessForIngestion: 'Enabled'
    publicNetworkAccessForQuery: 'Enabled'
  }
}

// RESSOURCE 2 : CLUSTER AKS
resource aksCluster 'Microsoft.ContainerService/managedClusters@2023-07-01' = {
  name: clusterName
  location: location
  identity: { type: 'SystemAssigned' }
  properties: {
    dnsPrefix: clusterName
    kubernetesVersion: kubernetesVersion == '' ? null : kubernetesVersion
    agentPoolProfiles: [
      {
        name: 'agentpool'
        count: nodeCount
        vmSize: nodeVmSize
        osType: 'Linux'
        osDiskSizeGB: osDiskSizeGB
        mode: 'System'
        type: 'VirtualMachineScaleSets'
        nodeLabels: { environment: 'tp', managedby: 'bicep' }
      }
    ]
    networkProfile: {
      networkPlugin: 'kubenet'
      loadBalancerSku: 'standard'
      outboundType: 'loadBalancer'
    }
    addonProfiles: {
      omsagent: {
        enabled: true
        config: {
          logAnalyticsWorkspaceResourceID: logAnalyticsWorkspace.id
        }
      }
    }
    enableRBAC: true
    disableLocalAccounts: false
    publicNetworkAccess: 'Enabled'
  }
}

// OUTPUTS
output clusterName string = aksCluster.name
output clusterResourceId string = aksCluster.id
output logAnalyticsWorkspaceId string = logAnalyticsWorkspace.id
output clusterPrincipalId string = aksCluster.identity.principalId
output clusterFqdn string = aksCluster.properties.fqdn
"@
    Add-Code $bicepCode

    Add-Heading2 "Explication des Blocs"
    Add-Normal "Bloc Paramètres : Tous les paramètres sont en haut du fichier avec des décorateurs @description() et des validations @minValue/@maxValue. Les valeurs par défaut permettent un déploiement sans paramètres externes."
    Add-Normal "Bloc Log Analytics Workspace : Crée le workspace de logs avant le cluster AKS (Bicep gère automatiquement l'ordre grâce à la dépendance implicite via logAnalyticsWorkspace.id)."
    Add-Normal "Bloc Cluster AKS : Définit le cluster avec une identité SystemAssigned (Azure crée et gère le service principal), un seul node pool de type System, et l'addon omsagent qui active Container Insights."
    Add-Normal "Bloc Outputs : Retourne les informations importantes après le déploiement (ID du cluster, FQDN de l'API server, ID du workspace)."

    Add-PageBreak

    # ================================================================
    # ÉTAPE 3 — PARAMÈTRES
    # ================================================================
    $Selection.Style = $Doc.Styles["Titre 1"]
    $Selection.TypeText("Étape 3 — Fichier main.parameters.json")
    $Selection.TypeParagraph()

    $paramsCode = @"
{
  "`$schema": "https://schema.management.azure.com/schemas/2019-04-01/deploymentParameters.json#",
  "contentVersion": "1.0.0.0",
  "parameters": {
    "clusterName":              { "value": "aks-tp-cluster" },
    "location":                 { "value": "francecentral" },
    "nodeVmSize":               { "value": "Standard_B2s" },
    "nodeCount":                { "value": 1 },
    "osDiskSizeGB":             { "value": 30 },
    "logAnalyticsWorkspaceName":{ "value": "law-aks-tp" },
    "logAnalyticsSku":          { "value": "PerGB2018" },
    "logRetentionDays":         { "value": 30 },
    "kubernetesVersion":        { "value": "" }
  }
}
"@
    Add-Code $paramsCode

    Add-Heading2 "Explication des Paramètres"
    Add-Normal "clusterName : 'aks-tp-cluster' — Nom du cluster visible dans le portail Azure et dans les commandes kubectl."
    Add-Normal "location : 'francecentral' — Région Paris. Doit correspondre à la région du resource group rg-aks-tp."
    Add-Normal "nodeVmSize : 'Standard_B2s' — VM burstable économique (2 vCPU, 4 Go RAM). Environ 30€/mois en continu. Arrêtez le cluster après le TP pour éviter les frais."
    Add-Normal "nodeCount : 1 — Un seul nœud suffit pour un TP. Minimise les coûts et la consommation de quota."
    Add-Normal "logAnalyticsSku : 'PerGB2018' — Facturation à l'usage. Les 5 premiers Go/mois sont gratuits, largement suffisant pour un TP."
    Add-Normal "logRetentionDays : 30 — 30 jours de rétention inclus dans le tier gratuit. Ne pas augmenter pour rester dans la limite gratuite."
    Add-Normal "kubernetesVersion : '' — Chaîne vide = Azure choisit la dernière version stable. Recommandé pour les TPs."

    Add-PageBreak

    # ================================================================
    # ÉTAPE 4 — DÉPLOIEMENT
    # ================================================================
    $Selection.Style = $Doc.Styles["Titre 1"]
    $Selection.TypeText("Étape 4 — Guide de Déploiement PowerShell")
    $Selection.TypeParagraph()

    Add-Heading2 "Prérequis"
    Add-Normal "Azure CLI installé et accessible dans PowerShell (az --version)"
    Add-Normal "kubectl installé (az aks install-cli si absent)"
    Add-Normal "Les fichiers main.bicep et main.parameters.json dans le même dossier"

    Add-Heading2 "Commandes de Déploiement"

    Add-Heading3 "1. Connexion à Azure"
    Add-Code "az login"
    Add-Normal "Une fenêtre de navigateur s'ouvre. Connectez-vous avec votre compte étudiant Azure."

    Add-Heading3 "2. Vérification de l'abonnement"
    Add-Code "az account show"
    Add-Code "az account list --output table  # Si plusieurs abonnements"
    Add-Code "az account set --subscription <ID>  # Pour changer d'abonnement"

    Add-Heading3 "3. Déploiement Bicep (5-10 minutes)"
    Add-Code "cd C:\Users\ldard\aks-tp"
    Add-Code "az deployment group create ``"
    Add-Code "    --resource-group rg-aks-tp ``"
    Add-Code "    --template-file main.bicep ``"
    Add-Code "    --parameters @main.parameters.json ``"
    Add-Code "    --name deploiement-aks-tp ``"
    Add-Code "    --output table"

    Add-Heading3 "4. Récupération des credentials kubectl"
    Add-Code "az aks get-credentials --resource-group rg-aks-tp --name aks-tp-cluster"

    Add-Heading3 "5. Vérification du cluster"
    Add-Code "kubectl get nodes"
    Add-Code "kubectl get pods --all-namespaces"
    Add-Code "kubectl get services --all-namespaces"

    Add-Heading3 "6. Déploiement de l'app de test"
    Add-Code "kubectl apply -f test-nginx.yaml"
    Add-Code "kubectl get pods -n default"
    Add-Code "kubectl get service nginx-test-service  # Attendre l'IP externe"

    Add-PageBreak

    # ================================================================
    # ÉTAPE 5 — MONITORING
    # ================================================================
    $Selection.Style = $Doc.Styles["Titre 1"]
    $Selection.TypeText("Étape 5 — Vérification du Monitoring")
    $Selection.TypeParagraph()

    Add-Heading2 "Vérifier l'Agent OMS avec kubectl"
    Add-Code "# Vérifier que l'agent OMS est bien déployé sur chaque node"
    Add-Code "kubectl get daemonset -n kube-system"
    Add-Code ""
    Add-Code "# Résultat attendu :"
    Add-Code "NAME                       DESIRED   CURRENT   READY   UP-TO-DATE   AVAILABLE"
    Add-Code "omsagent                   1         1         1       1            1"
    Add-Code ""
    Add-Code "# Vérifier les pods OMS"
    Add-Code "kubectl get pods -n kube-system -l component=oms-agent"
    Add-Code ""
    Add-Code "# Voir les logs de l'agent OMS (debug)"
    Add-Code "kubectl logs -n kube-system -l component=oms-agent"

    Add-Heading2 "Accéder à Container Insights dans le portail Azure"
    Add-Normal "Chemin exact dans le portail Azure (portal.azure.com) :"
    Add-Normal "1. Portail Azure → Moniteur (dans la barre de gauche)"
    Add-Normal "2. Conteneurs → Insights (ou directement : Moniteur > Conteneurs)"
    Add-Normal "3. Onglet 'Clusters' → Sélectionner 'aks-tp-cluster'"
    Add-Normal "4. Navigation dans les onglets : Cluster / Nœuds / Contrôleurs / Conteneurs"
    Add-Normal ""
    Add-Normal "Alternative directe : Portail Azure → Groupes de ressources → rg-aks-tp → aks-tp-cluster → Surveillance → Container Insights"

    Add-Heading2 "Ce qu'on voit dans Azure Monitor"
    Add-Normal "Onglet Cluster : CPU et mémoire totaux du cluster, nombre de pods actifs, santé globale."
    Add-Normal "Onglet Nœuds : Métriques par node (CPU %, RAM %, disque, réseau). Idéal pour détecter un node surchargé."
    Add-Normal "Onglet Contrôleurs : Vue par Deployment/DaemonSet. Permet de voir l'état de chaque objet Kubernetes."
    Add-Normal "Onglet Conteneurs : Liste tous les conteneurs en cours d'exécution avec leur consommation. Cliquez sur un conteneur → 'Voir les journaux en direct' pour voir les logs stdout en temps réel."

    Add-PageBreak

    # ================================================================
    # ÉTAPE 6 — ERREURS COURANTES
    # ================================================================
    $Selection.Style = $Doc.Styles["Titre 1"]
    $Selection.TypeText("Étape 6 — Erreurs Courantes et Solutions")
    $Selection.TypeParagraph()

    Add-Heading2 "1. Quota SKU insuffisant"
    Add-Code "Erreur : OperationNotAllowed - The subscription does not have"
    Add-Code "         enough quota for Standard_B2s in francecentral"
    Add-Normal "Solution : Vérifier les quotas disponibles :"
    Add-Code "az vm list-usage --location francecentral --query ""[?contains(name.value,'standardBSFamily')]"" -o table"
    Add-Normal "Si quota insuffisant, essayer une autre région :"
    Add-Code "# Modifier location dans main.parameters.json :"
    Add-Code "{ ""location"": { ""value"": ""westeurope"" } }"
    Add-Normal "Ou demander une augmentation de quota dans : Portail Azure → Abonnements → Utilisation + Quotas"

    Add-Heading2 "2. Mauvaise région (conflit resource group)"
    Add-Code "Erreur : The resource group 'rg-aks-tp' is in location 'francecentral'."
    Add-Code "         The template deployment is targeting location 'westeurope'."
    Add-Normal "Solution : Le paramètre location doit correspondre à la région du resource group. Vérifier :"
    Add-Code "az group show --name rg-aks-tp --query location"
    Add-Normal "Puis aligner le paramètre location dans main.parameters.json avec cette valeur."

    Add-Heading2 "3. Azure CLI non reconnu dans PowerShell"
    Add-Code "Erreur : 'az' is not recognized as the name of a cmdlet"
    Add-Normal "Solution : Réinstaller Azure CLI depuis https://docs.microsoft.com/cli/azure/install-azure-cli-windows"
    Add-Normal "Puis relancer PowerShell en tant qu'administrateur et vérifier :"
    Add-Code "az --version"
    Add-Normal "Si toujours absent, ajouter manuellement au PATH :"
    Add-Code '$env:PATH += ";C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin"'

    Add-Heading2 "4. Log Analytics Workspace déjà existant"
    Add-Code "Erreur : WorkspaceAlreadyExists - A workspace with the name 'law-aks-tp'"
    Add-Code "         already exists in the subscription."
    Add-Normal "Solution A : Renommer le workspace dans les paramètres :"
    Add-Code '{ "logAnalyticsWorkspaceName": { "value": "law-aks-tp-2" } }'
    Add-Normal "Solution B : Supprimer l'ancien workspace :"
    Add-Code "az monitor log-analytics workspace delete --resource-group rg-aks-tp --workspace-name law-aks-tp"

    Add-Heading2 "5. Cluster déjà existant"
    Add-Code "Erreur : ManagedClustersAlreadyExists"
    Add-Normal "Solution : Le déploiement Bicep est idempotent. Si le cluster existe déjà, il sera mis à jour. Si vous voulez repartir à zéro :"
    Add-Code "az aks delete --resource-group rg-aks-tp --name aks-tp-cluster --yes --no-wait"

    Add-PageBreak

    # ================================================================
    # ÉTAPE 7 — BONNES PRATIQUES
    # ================================================================
    $Selection.Style = $Doc.Styles["Titre 1"]
    $Selection.TypeText("Étape 7 — Bonnes Pratiques")
    $Selection.TypeParagraph()

    Add-Heading2 "Organisation des Fichiers"
    Add-Code "aks-tp/"
    Add-Code "├── main.bicep               # Template principal"
    Add-Code "├── main.parameters.json     # Paramètres (versionnable sans secrets)"
    Add-Code "├── deploy.ps1               # Script de déploiement"
    Add-Code "├── test-nginx.yaml          # Manifeste Kubernetes de test"
    Add-Code "└── README.md                # Documentation du projet"

    Add-Normal "Ne jamais versionner de secrets dans les fichiers de paramètres. Pour les mots de passe et certificats, utiliser Azure Key Vault."

    Add-Heading2 "Sécurité : Identity et RBAC"
    Add-Normal "SystemAssigned Identity : Le template utilise une identité managée par le système. Azure crée automatiquement un service principal lié au cycle de vie du cluster — pas besoin de gérer des credentials manuellement."
    Add-Normal "RBAC activé : enableRBAC: true dans le template garantit que l'accès au cluster est contrôlé par des rôles Azure AD."
    Add-Normal "Bonnes pratiques supplémentaires pour la production : Désactiver les comptes locaux (disableLocalAccounts: true), activer Azure AD integration pour l'authentification, utiliser des Network Policies pour isoler les pods."

    Add-Heading2 "Optimisation des Coûts"
    Add-Normal "Arrêt du cluster après le TP (pour éviter la facturation des VMs) :"
    Add-Code "az aks stop --resource-group rg-aks-tp --name aks-tp-cluster"
    Add-Code "az aks start --resource-group rg-aks-tp --name aks-tp-cluster  # Pour redémarrer"
    Add-Normal "Suppression complète après le TP :"
    Add-Code "az group delete --name rg-aks-tp --yes --no-wait"
    Add-Normal "Note : La suppression du resource group supprime TOUTES les ressources (cluster, workspace, IP publiques). Attention si d'autres ressources importantes y sont stockées."

    Add-Heading2 "Alertes de Monitoring à Configurer"
    Add-Normal "Dans Azure Monitor, configurer ces alertes de base :"
    Add-Normal "- CPU Node > 80% pendant 5 minutes → Alerte Warning"
    Add-Normal "- Mémoire Node > 85% pendant 5 minutes → Alerte Critical"
    Add-Normal "- Pod en état CrashLoopBackOff → Alerte immédiate"
    Add-Normal "- Nombre de pods prêts < replicas demandées → Alerte Warning"
    Add-Normal ""
    Add-Normal "Chemin : Portail Azure → Moniteur → Alertes → Créer une règle d'alerte → Sélectionner aks-tp-cluster comme ressource."

    Add-PageBreak

    # ================================================================
    # SECTION FINALE — PRÉSENTATION ORALE
    # ================================================================
    $Selection.Style = $Doc.Styles["Titre 1"]
    $Selection.TypeText("Comment Présenter ce Projet à l'Oral")
    $Selection.TypeParagraph()

    Add-Heading2 "Structure de Présentation Recommandée (10-15 minutes)"

    Add-Heading3 "Introduction (1-2 min)"
    Add-Normal "Présenter le contexte : 'Ce projet déploie une infrastructure Kubernetes complète sur Azure en utilisant uniquement du code — aucune action manuelle dans le portail Azure.'"
    Add-Normal "Poser la problématique : 'Comment déployer une infrastructure reproductible, versionnable et supervisée ?'"

    Add-Heading3 "Architecture (2-3 min)"
    Add-Normal "Expliquer le diagramme : AKS ← OMS Agent → Log Analytics → Container Insights"
    Add-Normal "Souligner l'intérêt de l'IaC : 'Ce même cluster peut être recréé en 10 minutes dans n'importe quelle région par n'importe qui ayant accès au dépôt Git.'"

    Add-Heading3 "Démo en Direct (5-7 min)"
    Add-Normal "1. Montrer les fichiers Bicep et expliquer 3-4 lignes clés (identity, omsagent, agentPoolProfiles)"
    Add-Normal "2. Montrer le cluster déjà déployé : kubectl get nodes, kubectl get pods -A"
    Add-Normal "3. Ouvrir Container Insights dans le portail → montrer les métriques en temps réel"
    Add-Normal "4. Déployer nginx en direct : kubectl apply -f test-nginx.yaml → montrer l'apparition dans Container Insights"

    Add-Heading3 "Bilan et Apprentissages (2-3 min)"
    Add-Normal "Ce que j'ai appris : Concepts AKS, syntaxe Bicep, monitoring Azure"
    Add-Normal "Points d'amélioration possibles : Multi-node, auto-scaling, alertes, CI/CD avec GitHub Actions"
    Add-Normal "Questions fréquentes à préparer :"
    Add-Normal "- Quelle est la différence entre AKS et une VM classique ?"
    Add-Normal "- Pourquoi Bicep plutôt que Terraform ?"
    Add-Normal "- Comment garantir la haute disponibilité ?"

    Add-Heading2 "Points Clés à Retenir"
    Add-Normal "AKS = Kubernetes managé → Azure gère le control plane, vous gérez les workloads"
    Add-Normal "Bicep = IaC déclaratif → L'infrastructure est du code, versionnable avec Git"
    Add-Normal "Container Insights = Visibilité complète → Logs + Métriques + Alertes en un seul endroit"
    Add-Normal "SystemAssigned Identity = Sécurité sans gestion de secrets → Bonne pratique Azure"

    # ================================================================
    # SAUVEGARDE
    # ================================================================
    $Doc.SaveAs([ref]$OutputPath)
    $Doc.Close()
    $Word.Quit()

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "  DOCUMENT WORD GENERE AVEC SUCCES !      " -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "  Fichier : $OutputPath" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Green

} catch {
    Write-Host "ERREUR lors de la generation du document Word :" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Assurez-vous que Microsoft Word est installe sur ce poste." -ForegroundColor Yellow

    # Nettoyage en cas d'erreur
    if ($Doc) { try { $Doc.Close($false) } catch {} }
    if ($Word) { try { $Word.Quit() } catch {} }
    exit 1
}
