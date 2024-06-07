$TenantID = "$(TenantID)"
$ClientID = "$(ClientID)"
$ClientSecret = "$(ClientSecret)"
$SenderMailAccount = "$(SenderMailAccount)"
$SenderMailPassword = "$(SenderMailPassword)"
$DatasetName
$TestWorkspaceName = "$(TestWorkspaceName)"
$ProdWorkspaceName = "$(ProdWorkspaceName)"
$Workspace_Processing = "powerbi://api.powerbi.com/v1.0/myorg/$TestWorkspaceName"
$Workspace_Publish = "powerbi://api.powerbi.com/v1.0/myorg/$ProdWorkspaceName"
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
$SecuredApplicationSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($ClientID, $SecuredApplicationSecret)

function QueryExecutor {
    param($Server, $Query)
    if ($Server -eq "Processing") {
        $response = invoke-ascmd -Database $DatasetName -Query $Query -server $Workspace_Processing -ServicePrincipal -Tenant $TenantID -Credential $credential
    }
    elseif ($Server -eq "Publish") {
        $response = invoke-ascmd -Database $DatasetName -Query $Query -server $Workspace_Publish -ServicePrincipal -Tenant $TenantID -Credential $credential
    }
    $xml = [xml]$response
    $ns = New-Object Xml.XmlNamespaceManager $xml.NameTable
    $ns.AddNamespace("ns", "urn:schemas-microsoft-com:xml-analysis:rowset")
    return $xml.SelectNodes('//ns:row', $ns)
}

function MeasureValueComparison {
    Write-Host "Comparing Measure values: This feature will be coming soon."
    $excelDataMeasureValue = @()
    $ReturnData = [PSCustomObject]@{
        'HTML'     = ""
        'ExcelData' = $excelDataMeasureValue
        'FName'     = "MeasureValueComparison"
    }
    return $ReturnData
}

function TableAttributeComparison {
    Write-Host "Comparing Table Attributes"	
    $Query = "SELECT [CATALOG_NAME] AS [TabularName], [DIMENSION_UNIQUE_NAME] AS [DimensionName], HIERARCHY_CAPTION AS [AttributeName], HIERARCHY_IS_VISIBLE AS [IsAttributeVisible] FROM `$system`.MDSchema_hierarchies WHERE CUBE_NAME  ='Model'"

    $Dimension_Attribute_Processing = QueryExecutor -Server "Processing" -Query $Query
    $Dimension_Attribute_Publish = QueryExecutor -Server "Publish" -Query $Query

    $AttributeVisbilityChanged, $AttributeAdded, $AttributeDeleted, $RowCount = 0, 0, 0, 1
    $RowCount_Processing, $RowCount_Publish = $Dimension_Attribute_Processing.Count, $Dimension_Attribute_Publish.Count

    $excelDataTableAttributes = @()

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

    for ($i = 0; $i -lt $RowCount_Processing; $i++) {
        $flag1, $flag2, $flag3 = 0, 0, 0
        $Result11, $Result12, $Result13 = $Dimension_Attribute_Processing[$i].C1, $Dimension_Attribute_Processing[$i].C2, $Dimension_Attribute_Processing[$i].C3
        $Result21, $Result22, $Result23 = '', '', ''
        for ($j = 0; $j -lt $RowCount_Publish; $j++) {
            $Result21 = $Dimension_Attribute_Publish[$j].C1
            $Result22 = $Dimension_Attribute_Publish[$j].C2
            $Result23 = $Dimension_Attribute_Publish[$j].C3

            if ($Result11 -eq $Result21) {
                $flag1 = 1                

                if ($Result12 -eq $Result22) {
                    $flag2 = 1                   

                    if ($Result13 -eq $Result23) {
                        $flag3 = 1
                    }
                }
        
            }
        }

        if ($flag3 -eq 1) {
            $Result, $BackgroundColor = "No Change", ""
            if ($ShowChangesOnly -eq "1") { continue }
        }
        elseif ($flag2 -eq 1 -and $flag3 -eq 0) {
            $Result, $BackgroundColor = "Attribute Visibility Changed", "#FFB52E"
            $AttributeVisbilityChanged++
        }
        elseif ($flag1 -eq 1 -and $flag2 -eq 0) {
            $Result, $BackgroundColor = "Attribute Added", "#B3EDA7"
            $AttributeAdded++
        }
        elseif ($flag1 -eq 0) {
            $Result, $BackgroundColor = "Table/Attribute Added", "#B3EDA7"
            $AttributeAdded++
        }

        $HTMLTable += @"
        <tr style="border:1px solid black; padding-right:5px; padding-left:5px"> 
        <td style="border:1px solid black;text-align: left">$RowCount</td> 
        <td style="border:1px solid black;text-align: left">$Result11</td> 
        <td style="border:1px solid black;text-align: left">$Result12</td> 
        <td style="border:1px solid black;text-align: left">$Result13</td> 
        <td style="border:1px solid black;text-align: left; background-color:$BackgroundColor">$Result</td>
        </tr>
"@
        if ($IsExportToExcel -eq "1") {
            $excelDataTableAttributes += [PSCustomObject]@{
                'Table Name'     = $Result11
                'Attribute Name' = $Result12
                'Is Visible'     = $Result13
                'Status'         = $Result
            }
        }
        
        $RowCount++
    }

    for ($i = 0; $i -lt $RowCount_Publish; $i++) {
        $flag1, $flag2, $flag3 = 0, 0, 0
        $Result11, $Result12, $Result13 = $Dimension_Attribute_Publish[$i].C1, $Dimension_Attribute_Publish[$i].C2, $Dimension_Attribute_Publish[$i].C3
        $Result21, $Result22, $Result23 = '', '', ''

        for ($j = 0; $j -lt $RowCount_Processing; $j++) {
            $Result21, $Result22, $Result23 = $Dimension_Attribute_Processing[$j].C1, $Dimension_Attribute_Processing[$j].C2, $Dimension_Attribute_Processing[$j].C3
            if ($Result11 -eq $Result21) {
                $flag1 = 1
                if ($Result12 -eq $Result22) {
                    $flag2 = 1
                    if ($Result13 -eq $Result23) {
                        $flag3 = 1
                    }
                }
            }
        }
        if ($flag1 -eq 1 -and $flag2 -eq 0 -or $flag1 -eq 0) {
            $Result = if ($flag1 -eq 1 -and $flag2 -eq 0) { "Attribute Deleted" } else { "Table/Attribute Deleted" }
            $AttributeDeleted++
            $HTMLTable += "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td> 
                <td style=""border:1px solid black;text-align: left"">$Result13</td> 
                <td style=""border:1px solid black;text-align: left; background-color:#ffa6a6"">$Result</td>
            </tr>"
            if ($IsExportToExcel -eq "1") {
                $excelDataTableAttributes += [PSCustomObject]@{
                    'Table Name'     = $Result11
                    'Attribute Name' = $Result12
                    'Is Visible'     = $Result13
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
        'ExcelData' = $excelDataTableAttributes
        'FName'     = "TableAttributeComparison"
    }
    return $ReturnData
}

function RelationshipComparison {
    Write-Host "Comparing Relationships"
    $Query = "SELECT [MEASUREGROUP_NAME] AS [LeftHand], [DIMENSION_UNIQUE_NAME] AS [RightHand], [MEASUREGROUP_CARDINALITY] AS [LeftHandCard], [DIMENSION_CARDINALITY] AS [RightHandCard], [DIMENSION_GRANULARITY] AS [RightHandKey] FROM `$system`.MDSCHEMA_MEASUREGROUP_DIMENSIONS WHERE [CUBE_NAME] ='Model'"

    $Relationship_Processing = QueryExecutor -Server "Processing" -Query $Query
    $Relationship_Publish = QueryExecutor -Server "Publish" -Query $Query

    $RowCount_Processing, $RowCount_Publish = $Relationship_Processing.Count, $Relationship_Publish.Count
    $RowCount, $RelationshipAdded, $RelationshipDeleted, $CardinalityChanged, $RelationshipAttributeChanged = 1, 0, 0, 0, 0

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

    for ($i = 0; $i -lt $RowCount_Processing; $i++) {
        $flag1, $flag2, $flag3, $flag4, $flag5 = 0, 0, 0, 0, 0
        $Result10, $Result20, $Result11, $Result21, $Result12, $Result22, $Result13, $Result23, $Result14, $Result24 = $Relationship_Processing[$i].C0, '', $Relationship_Processing[$i].C1, '', $Relationship_Processing[$i].C2, '', $Relationship_Processing[$i].C3, '', $Relationship_Processing[$i].C4, ''

        for ($j = 0; $j -lt $RowCount_Publish; $j++) {
            $Result20, $Result21, $Result22, $Result23, $Result24 = $Relationship_Publish[$j].C0, $Relationship_Publish[$j].C1, $Relationship_Publish[$j].C2, $Relationship_Publish[$j].C3, $Relationship_Publish[$j].C4

            if ($Result10 -eq $Result20) {
                $flag1 = 1                
                if ($Result11 -eq $Result21) {
                    $flag2 = 1                   
                    if ($Result12 -eq $Result22) {
                        $flag3 = 1
                        if ($Result13 -eq $Result23) {
                            $flag4 = 1
                            if ($Result14 -eq $Result24) {
                                $flag5 = 1
                            }
                        }
                    }
                }
            }
        }

        if ($flag5 -eq 1) {
            $Result, $BackgroundColor = "No Change", ""
            if ($ShowChangesOnly -eq "1") { continue }
        }
        elseif ($flag4 -eq 1 -and $flag5 -eq 0) {
            $Result, $BackgroundColor = "Relationship Attribute Changed", "#FFB52E"
            $RelationshipAttributeChanged++
        }
        elseif (($flag3 -eq 1 -and $flag4 -eq 0) -or ($flag2 -eq 1 -and $flag3 -eq 0)) {
            $Result, $BackgroundColor = "Cardinality Changed", "#FFB52E"
            $CardinalityChanged++
        }
        elseif ($flag1 -eq 1 -and $flag2 -eq 0) {
            $Result, $BackgroundColor = "Relationship Added", "#B3EDA7"
            $RelationshipAdded++
        }
        else {
            continue
        }
        $HTMLTable += @"
        <tr style="border:1px solid black; padding-right:5px; padding-left:5px"> 
        <td style="border:1px solid black;text-align: left">$RowCount</td> 
        <td style="border:1px solid black;text-align: left">$Result10</td> 
        <td style="border:1px solid black;text-align: left">$Result11</td> 
        <td style="border:1px solid black;text-align: left">$Result12</td>
        <td style="border:1px solid black;text-align: left">$Result13</td>
        <td style="border:1px solid black;text-align: left">$Result14</td> 
        <td style="border:1px solid black;text-align: left; background-color:$BackgroundColor">$Result</td>
        </tr>
"@
        if ($IsExportToExcel -eq "1") {
            $excelDataRelationship += [PSCustomObject]@{
                'Fact Name'              = $Result10
                'Dim Name'               = $Result11
                'Fact Cardinality'       = $Result12
                'Dim Cardinality'        = $Result13
                'Relationship On Column' = $Result14
                'Status'                 = $Result
            }
        }
        $RowCount++
    }

    for ($i = 0; $i -lt $RowCount_Publish; $i++) {
        $flag1, $flag2 = 0, 0
        $Result10, $Result11, $Result12, $Result13, $Result14 = $Relationship_Publish[$i].C0, $Relationship_Publish[$i].C1, $Relationship_Publish[$i].C2, $Relationship_Publish[$i].C3, $Relationship_Publish[$i].C4
        $Result20, $Result21, $Result22, $Result23, $Result24 = '', '', '', '', ''
        
        for ($j = 0; $j -lt $RowCount_Processing; $j++) {
            $Result20, $Result21, $Result22, $Result23, $Result24 = $Relationship_Processing[$j].C0, $Relationship_Processing[$j].C1, $Relationship_Processing[$j].C2, $Relationship_Processing[$j].C3, $Relationship_Processing[$j].C4

            if ($Result11 -eq $Result21) {
                $flag1 = 1
                if ($Result12 -eq $Result22) {
                    $flag2 = 1
                }
            }
        }

        if (($flag1 -eq 1 -and $flag2 -eq 0) -or $flag1 -eq 0) {
            $Result = "Relationship Deleted"
            $RelationshipDeleted++
            $HTMLTable += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
            <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
            <td style=""border:1px solid black;text-align: left"">$Result10</td> 
            <td style=""border:1px solid black;text-align: left"">$Result11</td> 
            <td style=""border:1px solid black;text-align: left"">$Result12</td>
            <td style=""border:1px solid black;text-align: left"">$Result13</td>
            <td style=""border:1px solid black;text-align: left"">$Result14</td>
            <td style=""border:1px solid black;text-align: left; background-color:#ffa6a6"">$Result</td>
            </tr>"

            if ($IsExportToExcel -eq "1") {
                $excelDataRelationship += [PSCustomObject]@{
                    'Fact Name'              = $Result10
                    'Dim Name'               = $Result11
                    'Fact Cardinality'       = $Result12
                    'Dim Cardinality'        = $Result13
                    'Relationship On Column' = $Result14
                    'Status'                 = $Result
                }
            }
            $RowCount++
        }
    }

    $HTMLTable += "</table>"
    $Summary = "<p><b>Relationship(s) Added:</b> $RelationshipAdded; &emsp; <b>Relationship(s) Deleted:</b> $RelationshipDeleted; &emsp; <b>Cardinality Changed:</b> $CardinalityChanged; &emsp; <b>Relationship Attribute Changed:</b> $RelationshipAttributeChanged</p><br><br>"
    $ReturnData = [PSCustomObject]@{
        'HTML'     = $Summary + $HTMLTable
        'ExcelData' = $excelDataRelationship
        'FName'     = "RelationshipComparison"
    }
    return $ReturnData
}

function MeasureDefinitionComparison {
    Write-Host "Comparing Measure Definition"
    $Query = "SELECT [MEASURE_NAME] AS [MeasureName], [DEFAULT_FORMAT_STRING] AS [MeasureFormat], [EXPRESSION] AS [MeasureDefinition], [MEASURE_IS_VISIBLE] AS [IsVisible] FROM `$SYSTEM`.MDSCHEMA_MEASURES WHERE CUBE_NAME  ='Model'"
    
    $MeasureDefinition_Processing = QueryExecutor -Server "Processing" -Query $Query
    $MeasureDefinition_Publish = QueryExecutor -Server "Publish" -Query $Query

    $RowCount_Processing, $RowCount_Publish = $MeasureDefinition_Processing.Count, $MeasureDefinition_Publish.Count
    $RowCount, $MeasureDefinitionUpdated, $MeasureFormatUpdated, $MeasureDeleted, $MeasureAdded, $MeasureVisibilityChanged = 1, 0, 0, 0, 0 , 0

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
    
    for ($i = 0; $i -lt $RowCount_Processing; $i++) {
        $flag1, $flag2, $flag3, $flag4 = 0, 0, 0, 0
        $Result10, $Result11, $Result12, $Result13 = $MeasureDefinition_Processing[$i].C0, $MeasureDefinition_Processing[$i].C1, $MeasureDefinition_Processing[$i].C2, $MeasureDefinition_Processing[$i].C3
        $Result20, $Result21, $Result22, $Result23 = '', '', '', ''
    
        for ($j = 0; $j -lt $RowCount_Publish; $j++) {
            $Result20, $Result21, $Result22, $Result23 = $MeasureDefinition_Publish[$j].C0, $MeasureDefinition_Publish[$j].C1, $MeasureDefinition_Publish[$j].C2, $MeasureDefinition_Publish[$j].C3

            if ($Result10 -eq $Result20) {
                $flag1 = 1               
                if ($Result11 -eq $Result21) {
                    $flag2 = 1                   
                    if ($Result12 -eq $Result22) {
                        $flag3 = 1
                        if ($Result13 -eq $Result23) {
                            $flag4 = 1
                        }
                    }
                }
            }
        }
    
        if ($flag4 -eq 1) {
            $Result = "No Change"
            $OldResult = ""
            $BGColor = ""
            if ($ShowChangesOnly -eq "1") { continue }
        }
        elseif ($flag3 -eq 1 -and $flag4 -eq 0) {
            $Result = "Measure Visibility Changed"
            $MeasureVisibilityChanged++
            $OldResult = ""
            $BGColor = "#FFB52E"
        }
        elseif ($flag2 -eq 1 -and $flag3 -eq 0) {
            $Result = "Definition Updated"
            $MeasureDefinitionUpdated++
            $OldResult = $MeasureDefinition_Publish[$i].C2
            $BGColor = "#FFB52E"
        }
        elseif ($flag1 -eq 1 -and $flag2 -eq 0) {
            $Result = "Format Updated"
            $MeasureFormatUpdated++
            $OldResult = $MeasureDefinition_Publish[$i].C1
            $BGColor = "#FFB52E"
        }
        elseif ($flag1 -eq 0) {
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
            <td style="border:1px solid black;text-align: left">$Result10</td> 
            <td style="border:1px solid black;text-align: left">$Result11</td> 
            <td style="border:1px solid black;text-align: left">$Result12</td>
            <td style="border:1px solid black;text-align: left">$OldResult</td>
            <td style="border:1px solid black;text-align: left">$Result13</td>
            <td style="border:1px solid black;text-align: left; background-color:$BGColor">$Result</td>
        </tr>
"@

        if ($IsExportToExcel -eq "1") {
            $excelDataMeasureDefinition += [PSCustomObject]@{
                'Measure Name'          = $Result10
                'Measure Format'        = $Result11
                'Measure Definition'    = $Result12
                'Old Definition/Format' = $OldResult
                'Is Visbile'            = $Result13
                'Status'                = $Result
            }
        }
        $RowCount++
    }
    
    for ($i = 0; $i -lt $RowCount_Publish; $i++) {
        $flag1 = 0
        $Result10, $Result11, $Result12, $Result13 = $MeasureDefinition_Publish[$i].C0, $MeasureDefinition_Publish[$i].C1, $MeasureDefinition_Publish[$i].C2, $MeasureDefinition_Publish[$i].C3
        $Result20, $Result21, $Result22, $Result23 = '', '', '', ''
        
        for ($j = 0; $j -lt $RowCount_Processing; $j++) {
            $Result20, $Result21, $Result22, $Result23 = $MeasureDefinition_Processing[$j].C0, $MeasureDefinition_Processing[$j].C1, $MeasureDefinition_Processing[$j].C2, $MeasureDefinition_Processing[$j].C3
            if ($Result10 -eq $Result20) { $flag1 = 1 }
        }
    
        if ($flag1 -eq 0) {
            $Result = "Measure Deleted"
            $MeasureDeleted++
            $HTMLTable += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left""></td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left; background-color:#ffa6a6"">$Result</td>
            </tr>"

            if ($IsExportToExcel -eq "1") {
                $excelDataMeasureDefinition += [PSCustomObject]@{
                    'Measure Name'          = $Result10
                    'Measure Format'        = $Result11
                    'Measure Definition'    = $Result12
                    'Old Definition/Format' = ""
                    'Is Visbile'            = $Result13
                    'Status'                = $Result
                }
            }
            $RowCount++
        }
    }
    
    $HTMLTable += "</table>"
    $Summary = "<p><b>Measure(s) Added:</b> $MeasureAdded; &emsp; <b>Measure(s) Deleted:</b> $MeasureDeleted; &emsp; <b>Definition Updated:</b> $MeasureDefinitionUpdated; &emsp; <b>Format Updated:</b> $MeasureFormatUpdated; &emsp; <b>Visibility Changed:</b> $MeasureVisibilityChanged</p><br><br>"
    $ReturnData = [PSCustomObject]@{
        'HTML'     = $Summary + $HTMLTable
        'ExcelData' = $excelDataMeasureDefinition
        'FName'     = "MeasureDefinitionComparison"
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
            1 { MeasureValueComparison }
            2 { TableAttributeComparison }
            3 { RelationshipComparison }
            4 { MeasureDefinitionComparison }
        }
    
        # Assign values to the respective arrays
        $html += $values.HTML
        $ExcelResult += ,@($values.ExcelData)
        $FunctionName += $values.FName
    }



    $html = $html -join "<br>"
    if ($html -eq "") { return }
    
    $Header = "<h2>Verificaton Summary: </h2>" + ($FunctionName -join ", ") + "
        <p><b>Processing Workspace:</b> $Workspace_Processing</p>
        <p><b>Processing Dataset:</b> $DatasetName</p>
        <p><b>Publish Workspace:</b> $Workspace_Publish</p>
        <p><b>Publish Dataset:</b> $DatasetName</p>"
    $FinalOutputBody = $Header + $html
    $Subject = "Power BI Dataset Comparison for " + $DatasetName + " as on " + $CurrentDate
    
    if ($IsSendMail -eq "1") {
        SendMail -Subject $Subject -Body $FinalOutputBody 
    }
    if ($IsExportToExcel -eq "1") {
        ExportToExcel -FunctionName $FunctionName -ExcelResult $ExcelResult
    }
}

Connect-PowerBIServiceAccount -ServicePrincipal -Tenant $TenantID -Credential $credential
$TestWorkspaceID = (Get-PowerBIWorkspace -Name "$TestWorkspaceName").Id

$sourceBranchName, $targetBranchName = "$(System.PullRequest.SourceBranch)", "$(System.PullRequest.TargetBranch)" -replace 'refs/heads/', ''
$changedFiles = git diff "origin/$targetBranchName...origin/$sourceBranchName" --name-only --diff-filter=M
foreach ($file in $changedFiles) {
    if ($file -like "*.pbix") {
        $DatasetName = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $FilePath = "$((Get-Location).path)\$file"
        Write-Host "DatasetName: $DatasetName`nProdWorkspaceName: $ProdWorkspaceName`nTestWorkspaceName: $TestWorkspaceName"
        New-PowerBIReport -Path $FilePath -Name $DatasetName -WorkspaceId $TestWorkspaceID -ConflictAction "CreateOrOverwrite"
        Write-Host "Dataset Publish Successfully"
        ExecuteFeatures
    }
}
