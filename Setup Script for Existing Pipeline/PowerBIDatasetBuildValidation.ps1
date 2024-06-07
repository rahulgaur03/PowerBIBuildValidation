# Program Name: Power BI Dataset Comparisonc   
# Date: 01/18/2024
# Author: Rahul Gaur
# Description: Created a PowerShell script for comparing Power BI dataset features, such as Measure Values, Table Attributes, Relationships, and Measure Definitions.

# Feature Descriptions:
# 1. Measure Value: Compares measure values in the Power BI dataset. [Feature under development]
# 2. Schema: Compares table attributes in the Power BI dataset.
# 3. Relationship: Compares relationships between tables in the Power BI dataset.
# 4. Definition: Compares measure definitions in the Power BI dataset.

# Modification Log
# ------------------------------------------------------------------------------------------------------------------
# Date         Author         Description of Changes
# ------------------------------------------------------------------------------------------------------------------
# 01/18/2024   Rahul Gaur        Initial script creation for Power BI dataset feature comparison for build pipeline.

if ("$(Build.Reason)" -eq "PullRequest") {
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
    Import-Module -Name MicrosoftPowerBIMgmt
    Connect-PowerBIServiceAccount -ServicePrincipal -Tenant "93e4b8d8-5034-3456-98a3-0ba2ca12ac84" -Credential (New-Object PSCredential "$(PowerBI-API-Client-ID)", (ConvertTo-SecureString -String "$(PowerBI-API-Client-Secret)" -AsPlainText -Force))

    $jsonObject = Get-Content -Raw -Path "$((Get-Location).path)\Tools\DatasetWorkspaceMapping.json" | ConvertFrom-Json
    $sourceBranchName, $targetBranchName = "$(System.PullRequest.SourceBranch)", "$(System.PullRequest.TargetBranch)" -replace 'refs/heads/', ''
    $changedFiles = git diff "origin/$targetBranchName...origin/$sourceBranchName" --name-only --diff-filter=M
    foreach ($file in $changedFiles) {
        if ($file -like "PowerBIDataset/*.pbix") {
            $DatasetName = [System.IO.Path]::GetFileNameWithoutExtension($file)
            $FilePath = "$((Get-Location).path)\$file"
            $ProdWorkspaceName = $jsonObject.$DatasetName.ProdWorkspaceName.Trim()
            $TestWorkspaceName = $jsonObject.$DatasetName.TestWorkspaceName.Trim()
            Write-Host "DatasetName: $DatasetName`nProdWorkspaceName: $ProdWorkspaceName`nTestWorkspaceName: $TestWorkspaceName"
            $TestWorkspaceID = (Get-PowerBIWorkspace -Name "$TestWorkspaceName").Id
            New-PowerBIReport -Path $FilePath -Name $DatasetName -WorkspaceId $TestWorkspaceID -ConflictAction "CreateOrOverwrite"
            Write-Host "Dataset Publish Successfully"
            & "$(Get-Location)\Tools\DataModelAnalyzer.ps1" -ClientID "$(PowerBI-API-Client-ID)" -ClientSecret "$(PowerBI-API-Client-Secret)" -SenderMailAccount "$(EDW-Email-ServiceAccount-UserName)" -SenderMailPassword "$(EDW-Email-ServiceAccount-Secret)" -DatasetName $DatasetName -ProdWorkspaceName $ProdWorkspaceName -TestWorkspaceName $TestWorkspaceName -DirectoryPath "$(Build.ArtifactStagingDirectory)"
        }
    }
}