$TenantID = "$(TenantID)"
$ClientID = "$(ClientID)"
$ClientSecret = "$(ClientSecret)"
$SenderMailAccount = "$(SenderMailAccount)"
$SenderMailPassword = "$(SenderMailPassword)"
$DatasetName
$TestWorkspaceName = "$(TestWorkspaceName)"
$ProdWorkspaceName = "$(ProdWorkspaceName)"
$WSTest = "powerbi://api.powerbi.com/v1.0/myorg/$TestWorkspaceName"
$WSProd = "powerbi://api.powerbi.com/v1.0/myorg/$ProdWorkspaceName"
$FeatureArray = "$(FeatureArray)"
$ShowChangesOnly = "$(ShowChangesOnly)"
$IsExportToExcel = "$(IsExportToExcel)"
$IsSendMail = "$(IsSendMail)"
$MailReceipients = "$(MailReceipients)"
$DirectoryPath = "$(Build.ArtifactStagingDirectory)"

Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
Import-Module -Name MicrosoftPowerBIMgmt
Import-Module ImportExcel
Import-Module SqlServer

$CurrentDate = Get-Date -Format "dddd MM/dd/yyyy"
$SAS = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($ClientID, $SAS)

function QueryExecutor {
    param($Server, $Query)
    if ($Server -eq "Processing") {
        $response = invoke-ascmd -Database $DatasetName -Query $Query -server $WSTest -ServicePrincipal -Tenant $TenantID -Credential $credential
    }
    elseif ($Server -eq "Publish") {
        $response = invoke-ascmd -Database $DatasetName -Query $Query -server $WSProd -ServicePrincipal -Tenant $TenantID -Credential $credential
    }
    $xml = [xml]$response
    $ns = New-Object Xml.XmlNamespaceManager $xml.NameTable
    $ns.AddNamespace("ns", "urn:schemas-microsoft-com:xml-analysis:rowset")
    return $xml.SelectNodes('//ns:row', $ns)
}

function MeasureValue {
    Write-Host "Comparing Measure values: Coming soon."
    $excelDataMeasureValue = @()
    $ReturnData = [PSCustomObject]@{
        'HTML'     = ""
        'ExcelData' = $excelDataMeasureValue
        'FName'     = "MeasureValue"
    }
    return $ReturnData
}

function TableAttribute {
    Write-Host "Comparing Table Attributes"	
    $Query = "SELECT [CATALOG_NAME] AS [TabularName], [DIMENSION_UNIQUE_NAME] AS [DimensionName], HIERARCHY_CAPTION AS [AttributeName], HIERARCHY_IS_VISIBLE AS [IsAttributeVisible] FROM `$system`.MDSchema_hierarchies WHERE CUBE_NAME  ='Model'"

    $AttributeTest = QueryExecutor -Server "Processing" -Query $Query
    $AttributeProd = QueryExecutor -Server "Publish" -Query $Query

    $AttributeVisbilityChanged, $AttributeAdded, $AttributeDeleted, $RowCount = 0, 0, 0, 1
    $RowCountTest, $RowCountProd = $AttributeTest.Count, $AttributeProd.Count

    $excelAttributes = @()

    $HTMLTable = @"
        <table style="width:100%; border:1px solid black; border-collapse: collapse;">
            <tr style="font-size:15px; font-weight: bold; padding:5px; color: white; border:1px solid black; background-color:#468ABD">
                <th style="border:1px solid black">#</th>
                <th style="border:1px solid black">Table Name</th>
                <th style="border:1px solid black">Attribute Name</th>
                <th style="border:1px solid black">Is Visible</th>
                <th style="border:1px solid black">Status</th>
            </tr>
"@

    for ($i = 0; $i -lt $RowCountTest; $i++) {
        $F1, $F2, $F3 = 0, 0, 0
        $R11, $R12, $R13 = $AttributeTest[$i].C1, $AttributeTest[$i].C2, $AttributeTest[$i].C3
        $R21, $R22, $R23 = '', '', ''
        for ($j = 0; $j -lt $RowCountProd; $j++) {
            $R21 = $AttributeProd[$j].C1
            $R22 = $AttributeProd[$j].C2
            $R23 = $AttributeProd[$j].C3

            if ($R11 -eq $R21) {
                $F1 = 1                

                if ($R12 -eq $R22) {
                    $F2 = 1                   

                    if ($R13 -eq $R23) {
                        $F3 = 1
                    }
                }
        
            }
        }

        if ($F3 -eq 1) {
            $Result, $BGColor = "No Change", ""
            if ($ShowChangesOnly -eq "1") { continue }
        }
        elseif ($F2 -eq 1 -and $F3 -eq 0) {
            $Result, $BGColor = "Attribute Visibility Changed", "#FFB52E"
            $AttributeVisbilityChanged++
        }
        elseif ($F1 -eq 1 -and $F2 -eq 0) {
            $Result, $BGColor = "Attribute Added", "#B3EDA7"
            $AttributeAdded++
        }
        elseif ($F1 -eq 0) {
            $Result, $BGColor = "Table/Attribute Added", "#B3EDA7"
            $AttributeAdded++
        }

        $HTMLTable += @"
        <tr style="border:1px solid black; padding-right:5px; padding-left:5px"> 
        <td style="border:1px solid black;text-align: left">$RowCount</td> 
        <td style="border:1px solid black;text-align: left">$R11</td> 
        <td style="border:1px solid black;text-align: left">$R12</td> 
        <td style="border:1px solid black;text-align: left">$R13</td> 
        <td style="border:1px solid black;text-align: left; background-color:$BGColor">$Result</td>
        </tr>
"@
        if ($IsExportToExcel -eq "1") {
            $excelAttributes += [PSCustomObject]@{
                'Table Name'     = $R11
                'Attribute Name' = $R12
                'Is Visible'     = $R13
                'Status'         = $Result
            }
        }
        
        $RowCount++
    }

    for ($i = 0; $i -lt $RowCountProd; $i++) {
        $F1, $F2, $F3 = 0, 0, 0
        $R11, $R12, $R13 = $AttributeProd[$i].C1, $AttributeProd[$i].C2, $AttributeProd[$i].C3
        $R21, $R22, $R23 = '', '', ''

        for ($j = 0; $j -lt $RowCountTest; $j++) {
            $R21, $R22, $R23 = $AttributeTest[$j].C1, $AttributeTest[$j].C2, $AttributeTest[$j].C3
            if ($R11 -eq $R21) {
                $F1 = 1
                if ($R12 -eq $R22) {
                    $F2 = 1
                    if ($R13 -eq $R23) {
                        $F3 = 1
                    }
                }
            }
        }
        if ($F1 -eq 1 -and $F2 -eq 0 -or $F1 -eq 0) {
            $Result = if ($F1 -eq 1 -and $F2 -eq 0) { "Attribute Deleted" } else { "Table/Attribute Deleted" }
            $AttributeDeleted++
            $HTMLTable += "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$R11</td> 
                <td style=""border:1px solid black;text-align: left"">$R12</td> 
                <td style=""border:1px solid black;text-align: left"">$R13</td> 
                <td style=""border:1px solid black;text-align: left; background-color:#ffa6a6"">$Result</td>
            </tr>"
            if ($IsExportToExcel -eq "1") {
                $excelAttributes += [PSCustomObject]@{
                    'Table Name'     = $R11
                    'Attribute Name' = $R12
                    'Is Visible'     = $R13
                    'Status'         = $Result
                }
            }
            $RowCount++
        }        
        
    }

    $HTMLTable += "</table>"
    $Summary = "<p><b>Attribute(s) Added:</b> $AttributeAdded; &emsp; <b>Attribute(s) Deleted:</b> $AttributeDeleted; &emsp; <b>Attribute Visibility Updated:</b> $AttributeVisbilityChanged</p><br><br>"
    $ReturnData = [PSCustomObject]@{
        'HTML'     = $Summary + $HTMLTable
        'ExcelData' = $excelAttributes
        'FName'     = "TableAttribute"
    }
    return $ReturnData
}

function Relationship {
    Write-Host "Comparing Relationships"
    $Query = "SELECT [MEASUREGROUP_NAME] AS [LeftHand], [DIMENSION_UNIQUE_NAME] AS [RightHand], [MEASUREGROUP_CARDINALITY] AS [LeftHandCard], [DIMENSION_CARDINALITY] AS [RightHandCard], [DIMENSION_GRANULARITY] AS [RightHandKey] FROM `$system`.MDSCHEMA_MEASUREGROUP_DIMENSIONS WHERE [CUBE_NAME] ='Model'"

    $RelationshipTest = QueryExecutor -Server "Processing" -Query $Query
    $RelationshipProd = QueryExecutor -Server "Publish" -Query $Query

    $RowCountTest, $RowCountProd = $RelationshipTest.Count, $RelationshipProd.Count
    $RowCount, $RAdded, $RDeleted, $CardinalityChanged, $RAttributeChanged = 1, 0, 0, 0, 0

    $excelDataRelationship = @()

    $HTMLTable = @"
        <table style="width:100%; border:1px solid black; border-collapse: collapse;">
            <tr style="font-size:15px; font-weight: bold; padding:5px; color: white; border:1px solid black; background-color:#468ABD">
                <th style="border:1px solid black">#</th>
                <th style="border:1px solid black">Fact Name</th>
                <th style="border:1px solid black">Dim Name</th>
                <th style="border:1px solid black">Fact Cardinality</th>
                <th style="border:1px solid black">Dim Cardinality</th>
                <th style="border:1px solid black">Relationship On Column</th>
                <th style="border:1px solid black">Status</th>
            </tr>
"@

    for ($i = 0; $i -lt $RowCountTest; $i++) {
        $F1, $F2, $F3, $F4, $F5 = 0, 0, 0, 0, 0
        $R10, $R20, $R11, $R21, $R12, $R22, $R13, $R23, $R14, $R24 = $RelationshipTest[$i].C0, '', $RelationshipTest[$i].C1, '', $RelationshipTest[$i].C2, '', $RelationshipTest[$i].C3, '', $RelationshipTest[$i].C4, ''

        for ($j = 0; $j -lt $RowCountProd; $j++) {
            $R20, $R21, $R22, $R23, $R24 = $RelationshipProd[$j].C0, $RelationshipProd[$j].C1, $RelationshipProd[$j].C2, $RelationshipProd[$j].C3, $RelationshipProd[$j].C4

            if ($R10 -eq $R20) {
                $F1 = 1                
                if ($R11 -eq $R21) {
                    $F2 = 1                   
                    if ($R12 -eq $R22) {
                        $F3 = 1
                        if ($R13 -eq $R23) {
                            $F4 = 1
                            if ($R14 -eq $R24) {
                                $F5 = 1
                            }
                        }
                    }
                }
            }
        }

        if ($F5 -eq 1) {
            $Result, $BGColor = "No Change", ""
            if ($ShowChangesOnly -eq "1") { continue }
        }
        elseif ($F4 -eq 1 -and $F5 -eq 0) {
            $Result, $BGColor = "Relationship Attribute Changed", "#FFB52E"
            $RAttributeChanged++
        }
        elseif (($F3 -eq 1 -and $F4 -eq 0) -or ($F2 -eq 1 -and $F3 -eq 0)) {
            $Result, $BGColor = "Cardinality Changed", "#FFB52E"
            $CardinalityChanged++
        }
        elseif ($F1 -eq 1 -and $F2 -eq 0) {
            $Result, $BGColor = "Relationship Added", "#B3EDA7"
            $RAdded++
        }
        else {
            continue
        }
        $HTMLTable += @"
        <tr style="border:1px solid black; padding-right:5px; padding-left:5px"> 
        <td style="border:1px solid black;text-align: left">$RowCount</td> 
        <td style="border:1px solid black;text-align: left">$R10</td> 
        <td style="border:1px solid black;text-align: left">$R11</td> 
        <td style="border:1px solid black;text-align: left">$R12</td>
        <td style="border:1px solid black;text-align: left">$R13</td>
        <td style="border:1px solid black;text-align: left">$R14</td> 
        <td style="border:1px solid black;text-align: left; background-color:$BGColor">$Result</td>
        </tr>
"@
        if ($IsExportToExcel -eq "1") {
            $excelDataRelationship += [PSCustomObject]@{
                'Fact Name'              = $R10
                'Dim Name'               = $R11
                'Fact Cardinality'       = $R12
                'Dim Cardinality'        = $R13
                'Relationship On Column' = $R14
                'Status'                 = $Result
            }
        }
        $RowCount++
    }

    for ($i = 0; $i -lt $RowCountProd; $i++) {
        $F1, $F2 = 0, 0
        $R10, $R11, $R12, $R13, $R14 = $RelationshipProd[$i].C0, $RelationshipProd[$i].C1, $RelationshipProd[$i].C2, $RelationshipProd[$i].C3, $RelationshipProd[$i].C4
        $R20, $R21, $R22, $R23, $R24 = '', '', '', '', ''
        
        for ($j = 0; $j -lt $RowCountTest; $j++) {
            $R20, $R21, $R22, $R23, $R24 = $RelationshipTest[$j].C0, $RelationshipTest[$j].C1, $RelationshipTest[$j].C2, $RelationshipTest[$j].C3, $RelationshipTest[$j].C4

            if ($R11 -eq $R21) {
                $F1 = 1
                if ($R12 -eq $R22) {
                    $F2 = 1
                }
            }
        }

        if (($F1 -eq 1 -and $F2 -eq 0) -or $F1 -eq 0) {
            $Result = "Relationship Deleted"
            $RDeleted++
            $HTMLTable += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
            <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
            <td style=""border:1px solid black;text-align: left"">$R10</td> 
            <td style=""border:1px solid black;text-align: left"">$R11</td> 
            <td style=""border:1px solid black;text-align: left"">$R12</td>
            <td style=""border:1px solid black;text-align: left"">$R13</td>
            <td style=""border:1px solid black;text-align: left"">$R14</td>
            <td style=""border:1px solid black;text-align: left; background-color:#ffa6a6"">$Result</td>
            </tr>"

            if ($IsExportToExcel -eq "1") {
                $excelDataRelationship += [PSCustomObject]@{
                    'Fact Name'              = $R10
                    'Dim Name'               = $R11
                    'Fact Cardinality'       = $R12
                    'Dim Cardinality'        = $R13
                    'Relationship On Column' = $R14
                    'Status'                 = $Result
                }
            }
            $RowCount++
        }
    }

    $HTMLTable += "</table>"
    $Summary = "<p><b>Relationship(s) Added:</b> $RAdded; &emsp; <b>Relationship(s) Deleted:</b> $RDeleted; &emsp; <b>Cardinality Changed:</b> $CardinalityChanged; &emsp; <b>Relationship Attribute Changed:</b> $RAttributeChanged</p><br><br>"
    $ReturnData = [PSCustomObject]@{
        'HTML'     = $Summary + $HTMLTable
        'ExcelData' = $excelDataRelationship
        'FName'     = "Relationship"
    }
    return $ReturnData
}

function MeasureDefinition {
    Write-Host "Comparing Measure Definition"
    $Query = "SELECT [MEASURE_NAME] AS [MeasureName], [DEFAULT_FORMAT_STRING] AS [MeasureFormat], [EXPRESSION] AS [MeasureDefinition], [MEASURE_IS_VISIBLE] AS [IsVisible] FROM `$SYSTEM`.MDSCHEMA_MEASURES WHERE CUBE_NAME  ='Model'"
    
    $DefinitionTest = QueryExecutor -Server "Processing" -Query $Query
    $DefinitionProd = QueryExecutor -Server "Publish" -Query $Query

    $RowCountTest, $RowCountProd = $DefinitionTest.Count, $DefinitionProd.Count
    $RowCount, $DefinitionUpdated, $FormatUpdated, $MeasureDeleted, $MeasureAdded, $VisibilityChanged = 1, 0, 0, 0, 0 , 0

    $excelDataMeasureDefinition = @()
    
    $HTMLTable = "
        <table style=""width:100%; border:1px solid black; border-collapse: collapse;"">
            <tr style=""font-size:15px; font-weight: bold; padding:5px; color: white; border:1px solid black; background-color:#468ABD"">
                <th style=""border:1px solid black"">#</th>
                <th style=""border:1px solid black"">Measure Name</th>
                <th style=""border:1px solid black"">Measure Format</th>
                <th style=""border:1px solid black"">Measure Definition</th>
                <th style=""border:1px solid black"">Old Definition/Format</th>
                <th style=""border:1px solid black"">Is Visbile</th>
                <th style=""border:1px solid black"">Status</th>
            </tr>"
    
    for ($i = 0; $i -lt $RowCountTest; $i++) {
        $F1, $F2, $F3, $F4 = 0, 0, 0, 0
        $R10, $R11, $R12, $R13 = $DefinitionTest[$i].C0, $DefinitionTest[$i].C1, $DefinitionTest[$i].C2, $DefinitionTest[$i].C3
        $R20, $R21, $R22, $R23 = '', '', '', ''
    
        for ($j = 0; $j -lt $RowCountProd; $j++) {
            $R20, $R21, $R22, $R23 = $DefinitionProd[$j].C0, $DefinitionProd[$j].C1, $DefinitionProd[$j].C2, $DefinitionProd[$j].C3

            if ($R10 -eq $R20) {
                $F1 = 1               
                if ($R11 -eq $R21) {
                    $F2 = 1                   
                    if ($R12 -eq $R22) {
                        $F3 = 1
                        if ($R13 -eq $R23) {
                            $F4 = 1
                        }
                    }
                }
            }
        }
    
        if ($F4 -eq 1) {
            $Result = "No Change"
            $OldResult = ""
            $BGColor = ""
            if ($ShowChangesOnly -eq "1") { continue }
        }
        elseif ($F3 -eq 1 -and $F4 -eq 0) {
            $Result = "Measure Visibility Changed"
            $VisibilityChanged++
            $OldResult = ""
            $BGColor = "#FFB52E"
        }
        elseif ($F2 -eq 1 -and $F3 -eq 0) {
            $Result = "Definition Updated"
            $DefinitionUpdated++
            $OldResult = $DefinitionProd[$i].C2
            $BGColor = "#FFB52E"
        }
        elseif ($F1 -eq 1 -and $F2 -eq 0) {
            $Result = "Format Updated"
            $FormatUpdated++
            $OldResult = $DefinitionProd[$i].C1
            $BGColor = "#FFB52E"
        }
        elseif ($F1 -eq 0) {
            $Result = "Measure Added"
            $OldResult = ""
            $BGColor = "#B3EDA7"
            $MeasureAdded++
        }
        else {
            continue
        }
        $HTMLTable += @"
        <tr style="border:1px solid black; padding-right:5px; padding-left:5px"> 
            <td style="border:1px solid black;text-align: left">$RowCount</td> 
            <td style="border:1px solid black;text-align: left">$R10</td> 
            <td style="border:1px solid black;text-align: left">$R11</td> 
            <td style="border:1px solid black;text-align: left">$R12</td>
            <td style="border:1px solid black;text-align: left">$OldResult</td>
            <td style="border:1px solid black;text-align: left">$R13</td>
            <td style="border:1px solid black;text-align: left; background-color:$BGColor">$Result</td>
        </tr>
"@

        if ($IsExportToExcel -eq "1") {
            $excelDataMeasureDefinition += [PSCustomObject]@{
                'Measure Name'          = $R10
                'Measure Format'        = $R11
                'Measure Definition'    = $R12
                'Old Definition/Format' = $OldResult
                'Is Visbile'            = $R13
                'Status'                = $Result
            }
        }
        $RowCount++
    }
    
    for ($i = 0; $i -lt $RowCountProd; $i++) {
        $F1 = 0
        $R10, $R11, $R12, $R13 = $DefinitionProd[$i].C0, $DefinitionProd[$i].C1, $DefinitionProd[$i].C2, $DefinitionProd[$i].C3
        $R20, $R21, $R22, $R23 = '', '', '', ''
        
        for ($j = 0; $j -lt $RowCountTest; $j++) {
            $R20, $R21, $R22, $R23 = $DefinitionTest[$j].C0, $DefinitionTest[$j].C1, $DefinitionTest[$j].C2, $DefinitionTest[$j].C3
            if ($R10 -eq $R20) { $F1 = 1 }
        }
    
        if ($F1 -eq 0) {
            $Result = "Measure Deleted"
            $MeasureDeleted++
            $HTMLTable += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$R10</td> 
                <td style=""border:1px solid black;text-align: left"">$R11</td> 
                <td style=""border:1px solid black;text-align: left"">$R12</td>
                <td style=""border:1px solid black;text-align: left""></td>
                <td style=""border:1px solid black;text-align: left"">$R13</td>
                <td style=""border:1px solid black;text-align: left; background-color:#ffa6a6"">$Result</td>
            </tr>"

            if ($IsExportToExcel -eq "1") {
                $excelDataMeasureDefinition += [PSCustomObject]@{
                    'Measure Name'          = $R10
                    'Measure Format'        = $R11
                    'Measure Definition'    = $R12
                    'Old Definition/Format' = ""
                    'Is Visbile'            = $R13
                    'Status'                = $Result
                }
            }
            $RowCount++
        }
    }
    
    $HTMLTable += "</table>"
    $Summary = "<p><b>Measure(s) Added:</b> $MeasureAdded; &emsp; <b>Measure(s) Deleted:</b> $MeasureDeleted; &emsp; <b>Definition Updated:</b> $DefinitionUpdated; &emsp; <b>Format Updated:</b> $FormatUpdated; &emsp; <b>Visibility Changed:</b> $VisibilityChanged</p><br><br>"
    $ReturnData = [PSCustomObject]@{
        'HTML'     = $Summary + $HTMLTable
        'ExcelData' = $excelDataMeasureDefinition
        'FName'     = "MeasureDefinition"
    }
    return $ReturnData
}

function SendMail {
    param($Subject, $Body)
    $Credential = New-Object System.Management.Automation.PSCredential($SenderMailAccount, (ConvertTo-SecureString -String $SenderMailPassword -AsPlainText -Force))
    $EmailParameters = @{
        From       = $SenderMailAccount
        To         = $MailReceipients -split ';'
        Subject    = $Subject
        Body       = $Body
        BodyAsHtml = $true
        SmtpServer = "smtp.office365.com"
        Port       = 587
        Credential = $Credential
        UseSsl     = $true
    }
    Send-MailMessage @EmailParameters
    Write-Host "Mail Sent"
}

function ExportToExcel {
    param($FunctionName, $ExcelResult)
    $ExcelPath = "$DirectoryPath/ModelComparison_$DatasetName.xlsx"
    for ($i = 0; $i -lt $FunctionName.Length; $i++) {
        $ExcelResult[$i] | Export-Excel -Path $ExcelPath -WorksheetName $FunctionName[$i] -AutoSize -BoldTopRow
    }
    Write-Host "Data Exported in Excel"
}

function ExecuteFeatures {
    $html,$ExcelResult, $FunctionName  = @(), @(), @()
    $FeatureArray.split(",") | ForEach-Object {
        $values = switch ($_){
            1 { MeasureValue }
            2 { TableAttribute }
            3 { Relationship }
            4 { MeasureDefinition }
        }

        $html += $values.HTML
        $ExcelResult += ,@($values.ExcelData)
        $FunctionName += $values.FName
    }



    $html = $html -join "<br>"
    if ($html -eq "") { return }
    
    $Header = "<h2>Verificaton Summary: </h2>" + ($FunctionName -join ", ") + "
        <p><b>Processing Workspace:</b> $WSTest</p>
        <p><b>Processing Dataset:</b> $DatasetName</p>
        <p><b>Publish Workspace:</b> $WSProd</p>
        <p><b>Publish Dataset:</b> $DatasetName</p>"
    $FBody = $Header + $html
    $Subject = "Power BI Dataset Comparison for " + $DatasetName + " as on " + $CurrentDate
    
    if ($IsSendMail -eq "1") {
        SendMail -Subject $Subject -Body $FBody 
    }
    if ($IsExportToExcel -eq "1") {
        ExportToExcel -FunctionName $FunctionName -ExcelResult $ExcelResult
    }
}

Connect-PowerBIServiceAccount -ServicePrincipal -Tenant $TenantID -Credential $credential
$TWID = (Get-PowerBIWorkspace -Name "$TestWorkspaceName").Id

$sourceBranchName, $targetBranchName = "$(System.PullRequest.SourceBranch)", "$(System.PullRequest.TargetBranch)" -replace 'refs/heads/', ''
$changedFiles = git diff "origin/$targetBranchName...origin/$sourceBranchName" --name-only --diff-filter=M

foreach ($file in $changedFiles) {
    if ($file -like "*.pbix") {
        $DatasetName = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $FilePath = "$((Get-Location).path)\$file"
        Write-Host "DatasetName: $DatasetName`nProdWorkspaceName: $ProdWorkspaceName`nTestWorkspaceName: $TestWorkspaceName"
        New-PowerBIReport -Path $FilePath -Name $DatasetName -WorkspaceId $TWID -ConflictAction "CreateOrOverwrite"
        Write-Host "Dataset Publish Successfully"
        ExecuteFeatures
    }
}