# =============================================================
#  SCRIPT DE DEPLOIEMENT AKS - TP Azure / Kubernetes
#  Exécuter dans PowerShell depuis le dossier aks-tp/
#  Commande : cd C:\Users\ldard\aks-tp ; .\deploy.ps1
# =============================================================

$ResourceGroup = "rg-aks-tp"
$ClusterName   = "aks-tp-cluster"
$TemplateFile  = ".\main.bicep"
$ParamsFile    = ".\main.parameters.json"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  DEPLOIEMENT AKS - TP Azure / Kubernetes  " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# -----------------------------------------------------------
# ÉTAPE 4.1 — Connexion Azure
# -----------------------------------------------------------
Write-Host "`n[1/6] Connexion a Azure..." -ForegroundColor Yellow
az login
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERREUR : Echec de la connexion Azure." -ForegroundColor Red
    exit 1
}

# -----------------------------------------------------------
# ÉTAPE 4.2 — Vérification de l'abonnement actif
# -----------------------------------------------------------
Write-Host "`n[2/6] Verification de l abonnement actif..." -ForegroundColor Yellow
az account show --output table
Write-Host ""
Write-Host "Si ce n est pas le bon abonnement, executez :" -ForegroundColor Gray
Write-Host "  az account list --output table" -ForegroundColor Gray
Write-Host "  az account set --subscription <ID_ABONNEMENT>" -ForegroundColor Gray
Read-Host "Appuyez sur Entree pour continuer..."

# -----------------------------------------------------------
# ÉTAPE 4.3 — Vérification du resource group
# -----------------------------------------------------------
Write-Host "`n[3/6] Verification du resource group $ResourceGroup..." -ForegroundColor Yellow
$rgExists = az group exists --name $ResourceGroup
if ($rgExists -eq "false") {
    Write-Host "Resource group introuvable. Creation en cours..." -ForegroundColor Yellow
    az group create --name $ResourceGroup --location francecentral
} else {
    Write-Host "Resource group $ResourceGroup : OK" -ForegroundColor Green
}

# -----------------------------------------------------------
# ÉTAPE 4.4 — Déploiement Bicep
# -----------------------------------------------------------
Write-Host "`n[4/6] Deploiement du template Bicep (cela peut prendre 5 a 10 minutes)..." -ForegroundColor Yellow
az deployment group create `
    --resource-group $ResourceGroup `
    --template-file $TemplateFile `
    --parameters @$ParamsFile `
    --name "deploiement-aks-tp" `
    --output table

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERREUR : Le deploiement a echoue. Consultez les erreurs ci-dessus." -ForegroundColor Red
    Write-Host "Commande de diagnostic :" -ForegroundColor Gray
    Write-Host "  az deployment group show --resource-group $ResourceGroup --name deploiement-aks-tp" -ForegroundColor Gray
    exit 1
}
Write-Host "Deploiement termine avec succes !" -ForegroundColor Green

# -----------------------------------------------------------
# ÉTAPE 4.5 — Récupération des credentials kubectl
# -----------------------------------------------------------
Write-Host "`n[5/6] Configuration de kubectl..." -ForegroundColor Yellow
az aks get-credentials `
    --resource-group $ResourceGroup `
    --name $ClusterName `
    --overwrite-existing

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERREUR : Impossible de recuperer les credentials kubectl." -ForegroundColor Red
    exit 1
}
Write-Host "kubectl configure : OK" -ForegroundColor Green

# -----------------------------------------------------------
# ÉTAPE 4.6 — Vérifications du cluster
# -----------------------------------------------------------
Write-Host "`n[6/6] Verification du cluster Kubernetes..." -ForegroundColor Yellow

Write-Host "`n--- Noeuds du cluster ---" -ForegroundColor Cyan
kubectl get nodes -o wide

Write-Host "`n--- Pods de tous les namespaces ---" -ForegroundColor Cyan
kubectl get pods --all-namespaces

Write-Host "`n--- Services ---" -ForegroundColor Cyan
kubectl get services --all-namespaces

Write-Host "`n============================================" -ForegroundColor Green
Write-Host "  DEPLOIEMENT TERMINE AVEC SUCCES !        " -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Prochaines etapes :" -ForegroundColor Cyan
Write-Host "  1. Accedez au portail Azure : https://portal.azure.com" -ForegroundColor White
Write-Host "  2. Allez dans : Moniteur > Container Insights" -ForegroundColor White
Write-Host "  3. Deployer l app de test : kubectl apply -f test-nginx.yaml" -ForegroundColor White
