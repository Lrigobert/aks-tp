# Requêtes KQL — AKS Dashboard Log Analytics
# À exécuter dans : portail Azure → law-aks-tp → Logs

## 1. État de tous les pods AKS (dernière heure)
KubePodInventory
| where TimeGenerated > ago(1h)
| where ClusterName == "aks-tp-cluster"
| summarize arg_max(TimeGenerated, *) by Name
| project PodName=Name, Namespace, PodStatus, ContainerName, ClusterName
| order by Namespace asc

## 2. Utilisation CPU des conteneurs
Perf
| where TimeGenerated > ago(30m)
| where ObjectName == "K8SContainer"
| where CounterName == "cpuUsageNanoCores"
| summarize AvgCPU=avg(CounterValue)/1000000 by InstanceName, bin(TimeGenerated, 5m)
| render timechart

## 3. Utilisation mémoire des pods
KubePodInventory
| where TimeGenerated > ago(1h)
| where ClusterName == "aks-tp-cluster"
| where PodStatus == "Running"
| summarize count() by Namespace, bin(TimeGenerated, 5m)
| render timechart title="Pods Running par Namespace"

## 4. Logs des conteneurs nginx
ContainerLog
| where TimeGenerated > ago(1h)
| where Image contains "nginx"
| project TimeGenerated, ContainerID, LogEntry
| order by TimeGenerated desc
| take 50

## 5. Événements du cluster
KubeEvents
| where TimeGenerated > ago(2h)
| where ClusterName == "aks-tp-cluster"
| project TimeGenerated, Name, Namespace, Reason, Message, Type
| order by TimeGenerated desc

## 6. Santé des nœuds
KubeNodeInventory
| where TimeGenerated > ago(1h)
| summarize arg_max(TimeGenerated, *) by Computer
| project Node=Computer, Status, KubeletVersion, OSImage, ClusterName

## 7. Résumé global du cluster (pour la démo)
KubePodInventory
| where TimeGenerated > ago(1h)
| where ClusterName == "aks-tp-cluster"
| summarize arg_max(TimeGenerated, *) by Name
| summarize
    TotalPods     = count(),
    RunningPods   = countif(PodStatus == "Running"),
    PendingPods   = countif(PodStatus == "Pending"),
    FailedPods    = countif(PodStatus == "Failed")
| extend HealthRate = round(100.0 * RunningPods / TotalPods, 1)
