#Feature Number 1-Measure Value; 2-Schema; 3-Relationship; 4-Definition
# param(
#     $ProdWorkspaceName,
#     $DatasetName,
#     $SenderMailAccount,
#     $SenderMailPassword,
#     $ClientID,
#     $ClientSecret
# )

$ProdWorkspaceName = "Sprouts EDW"
$DatasetName = "Shrink Adjusted Margin"
$SenderMailAccount = "svc-edwemailtest@sproutsfm.onmicrosoft.com"
$SenderMailPassword = 'fK)1Lfy69I$toSUlTDrS'
$ClientID = "ddb776a3-ea78-4d50-8e51-2aa140618647"
$ClientSecret ="dwhibl7x7FC/IRE1eM5aeMEhSveYdfaKclCQCIaasGQ="

Import-Module SqlServer
$tenantId = "39e4b8d8-5034-4760-84a3-0ba2ca73ac84"
$Workspace_Processing = "powerbi://api.powerbi.com/v1.0/myorg/Sprouts EDW - Test"
$Workspace_Publish = "powerbi://api.powerbi.com/v1.0/myorg/$ProdWorkspaceName"
$ModelName_Processing = $DatasetName
$ModelName_Publish = $DatasetName
$FeatureArray = "2"
$MailReceipients = "rahulg@maqsoftware.com"
$CurrentDate = Get-Date -Format "dddd MM/dd/yyyy"

$SecuredApplicationSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($ClientID, $SecuredApplicationSecret)

function XMLToObject {
    param($response)
    $xml = [xml]$response
    $ns = New-Object Xml.XmlNamespaceManager $xml.NameTable
    $ns.AddNamespace("ns", "urn:schemas-microsoft-com:xml-analysis:rowset")
    return $xml.SelectNodes('//ns:row', $ns)
}

function TableAttributeComparison {
    $FeatureName = "Table/Attribute Comparison"
    $Query = "SELECT [CATALOG_NAME] AS [TabularName], [DIMENSION_UNIQUE_NAME] AS [DimensionName], HIERARCHY_CAPTION AS [AttributeName], HIERARCHY_IS_VISIBLE AS [IsAttributeVisible] FROM `$system`.MDSchema_hierarchies WHERE CUBE_NAME  ='Model'"

    $Dimension_Attribute_Processing = XMLToObject -response (invoke-ascmd -Database $ModelName_Processing -Query $Query -server $Workspace_Processing -ServicePrincipal -Tenant $tenantId -Credential $credential)
    $Dimension_Attribute_Publish = XMLToObject -response (invoke-ascmd -Database $ModelName_Publish -Query $Query -server $Workspace_Publish -ServicePrincipal -Tenant $tenantId -Credential $credential)

    $AttributeVisbilityChanged = 0
    $AttributeAdded = 0
    $AttributeDeleted = 0
    $RowCount = 1
    $RowCount_Processing, $RowCount_Publish = $Dimension_Attribute_Processing.Count, $Dimension_Attribute_Publish.Count

    $table_html = @"
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
        $Result11, $Result12, $Result13 = $Dimension_Attribute_Processing[$i].C1, $Dimension_Attribute_Processing[$i].C2, $Dimension_Attribute_Processing[$i].C3
        $flag1, $flag2, $flag3 = 0, 0, 0

        for ($j = 0; $j -lt $RowCount_Publish; $j++) {
            $Result21, $Result22, $Result23 = $Dimension_Attribute_Publish[$j].C1, $Dimension_Attribute_Publish[$j].C2, $Dimension_Attribute_Publish[$j].C3
            if ($Result11 -eq $Result21 -and $Result12 -eq $Result22 -and $Result13 -eq $Result23) { $flag3 = 1 }
            elseif ($Result11 -eq $Result21 -and $Result12 -eq $Result22) { $flag2 = 1 }
            elseif ($Result11 -eq $Result21) { $flag1 = 1 }
        }

        if ($flag3 -eq 1) { $Result, $bgcolor = "No Change", "" }
        elseif ($flag2 -eq 1) { $Result, $bgcolor = "Attribute Visbility Changed", "#FFB52E"; $AttributeVisbilityChanged++ }
        elseif ($flag1 -eq 1) { $Result, $bgcolor = "Attribute Added", "#B3EDA7"; $AttributeAdded++ }
        else { $Result, $bgcolor = "Table/Attribute Added", "#B3EDA7"; $AttributeAdded++ }

        $table_html += "<tr style='border:1px solid black; padding-right:5px; padding-left:5px'>
            <td style='border:1px solid black;text-align: left'>$RowCount</td>
            <td style='border:1px solid black;text-align: left'>$Result11</td>
            <td style='border:1px solid black;text-align: left'>$Result12</td>
            <td style='border:1px solid black;text-align: left'>$Result13</td>
            <td style='border:1px solid black;text-align: left; background-color:$bgcolor'>$Result</td>
        </tr>"
        $RowCount++
    }

    for ($i = 0; $i -lt $RowCount_Publish; $i++) {
        $Result11, $Result12, $Result13 = $Dimension_Attribute_Publish[$i].C1, $Dimension_Attribute_Publish[$i].C2, $Dimension_Attribute_Publish[$i].C3
        $flag1, $flag2, $flag3 = 0, 0, 0

        for ($j = 0; $j -lt $RowCount_Processing; $j++) {
            $Result21, $Result22, $Result23 = $Dimension_Attribute_Processing[$j].C1, $Dimension_Attribute_Processing[$j].C2, $Dimension_Attribute_Processing[$j].C3
            if ($Result11 -eq $Result21 -and $Result12 -eq $Result22 -and $Result13 -eq $Result23) { $flag3 = 1 }
            elseif ($Result11 -eq $Result21 -and $Result12 -eq $Result22) { $flag2 = 1 }
            elseif ($Result11 -eq $Result21) { $flag1 = 1 }
        }

        if ($flag1 -eq 1 -and $flag2 -eq 0) { $Result, $bgcolor = "Attribute Deleted", "#ffa6a6"; $AttributeDeleted++ }
        else { $Result, $bgcolor = "Table/Attribute Deleted", "#ffa6a6"; $AttributeDeleted++ }

        $table_html += "<tr style='border:1px solid black; padding-right:5px; padding-left:5px'>
            <td style='border:1px solid black;text-align: left'>$RowCount</td>
            <td style='border:1px solid black;text-align: left'>$Result11</td>
            <td style='border:1px solid black;text-align: left'>$Result12</td>
            <td style='border:1px solid black;text-align: left'>$Result13</td>
            <td style='border:1px solid black;text-align: left; background-color:$bgcolor'>$Result</td>
        </tr>"
        $RowCount++
    }

    $table_html += "</table>"
    $Summary = "<p><b>Attribute(s) Added:</b> $AttributeAdded; &emsp; <b>Attribute(s) Deleted:</b> $AttributeDeleted; &emsp; <b>Attribute Visibility Updated:</b> $AttributeVisbilityChanged</p><br><br>"
    return $Summary + $table_html
}


function RelationshipComparison {
    $FeatureName = "Relationship Comparison"
    $Query = "SELECT [MEASUREGROUP_NAME] AS [LeftHand], [DIMENSION_UNIQUE_NAME] AS [RightHand], [MEASUREGROUP_CARDINALITY] AS [LeftHandCard], [DIMENSION_CARDINALITY] AS [RightHandCard], [DIMENSION_GRANULARITY] AS [RightHandKey] FROM `$system`.MDSCHEMA_MEASUREGROUP_DIMENSIONS WHERE [CUBE_NAME] ='Model'"

    $Relationship_Processing = XMLToObject -response (invoke-ascmd -Database $ModelName_Processing -Query $Query -server $Workspace_Processing -ServicePrincipal -Tenant $tenantId -Credential $credential)
    $Relationship_Publish = XMLToObject -response (invoke-ascmd -Database $ModelName_Publish -Query $Query -server $Workspace_Publish -ServicePrincipal -Tenant $tenantId -Credential $credential)

    $RowCount_Processing, $RowCount_Publish = $Relationship_Processing.Count, $Relationship_Publish.Count
    $RowCount, $RelationshipAdded, $RelationshipDeleted, $CardinalityChanged, $RelationshipAttributeChanged = 1, 0, 0, 0, 0

    $table_html = @"
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
        $Result10, $Result11, $Result12, $Result13, $Result14 = $Relationship_Processing[$i].C0, $Relationship_Processing[$i].C1, $Relationship_Processing[$i].C2, $Relationship_Processing[$i].C3, $Relationship_Processing[$i].C4
        $Result20, $Result21, $Result22, $Result23, $Result24 = '', '', '', '', ''

        for ($j = 0; $j -lt $RowCount_Publish; $j++) {
            $Result20, $Result21, $Result22, $Result23, $Result24 = $Relationship_Publish[$j].C0, $Relationship_Publish[$j].C1, $Relationship_Publish[$j].C2, $Relationship_Publish[$j].C3, $Relationship_Publish[$j].C4

            if ($Result10 -eq $Result20 -and $Result11 -eq $Result21 -and $Result12 -eq $Result22 -and $Result13 -eq $Result23 -and $Result14 -eq $Result24) { $flag5 = 1 }
            elseif ($Result10 -eq $Result20 -and $Result11 -eq $Result21 -and $Result12 -eq $Result22 -and $Result13 -eq $Result23) { $flag4 = 1 }
            elseif ($Result10 -eq $Result20 -and $Result11 -eq $Result21 -and $Result12 -eq $Result22) { $flag3 = 1 }
            elseif ($Result10 -eq $Result20 -and $Result11 -eq $Result21) { $flag2 = 1 }
            elseif ($Result10 -eq $Result20) { $flag1 = 1 }
        }

        if ($flag5 -eq 1) { $Result, $bgcolor = "No Change", "" }
        elseif ($flag4 -eq 1) { $Result, $bgcolor = "Relationship Attribute Changed", "#FFB52E"; $RelationshipAttributeChanged++ }
        elseif ($flag3 -eq 1 -or ($flag2 -eq 1 -and $flag3 -eq 0)) { $Result, $bgcolor = "Cardinality Changed", "#FFB52E"; $CardinalityChanged++ }
        elseif ($flag1 -eq 1 -and $flag2 -eq 0) { $Result, $bgcolor = "Relationship Added", "#B3EDA7"; $RelationshipAdded++ }

        $table_html += "<tr style='border:1px solid black; padding-right:5px; padding-left:5px'>
            <td style='border:1px solid black;text-align: left'>$RowCount</td>
            <td style='border:1px solid black;text-align: left'>$Result10</td>
            <td style='border:1px solid black;text-align: left'>$Result11</td>
            <td style='border:1px solid black;text-align: left'>$Result12</td>
            <td style='border:1px solid black;text-align: left'>$Result13</td>
            <td style='border:1px solid black;text-align: left'>$Result14</td>
            <td style='border:1px solid black;text-align: left; background-color:$bgcolor'>$Result</td>
        </tr>"
        $RowCount++
    }

    for ($i = 0; $i -lt $RowCount_Publish; $i++) {
        $flag1 = 0
        $Result10, $Result11, $Result12, $Result13, $Result14 = $Relationship_Publish[$i].C0, $Relationship_Publish[$i].C1, $Relationship_Publish[$i].C2, $Relationship_Publish[$i].C3, $Relationship_Publish[$i].C4
        $Result20, $Result21, $Result22, $Result23, $Result24 = '', '', '', '', ''

        for ($j = 0; $j -lt $RowCount_Processing; $j++) {
            $Result20, $Result21, $Result22, $Result23, $Result24 = $Relationship_Processing[$j].C0, $Relationship_Processing[$j].C1, $Relationship_Processing[$j].C2, $Relationship_Processing[$j].C3, $Relationship_Processing[$j].C4

            if ($Result11 -eq $Result21 -and $Result12 -eq $Result22) { $flag1 = 1 }
        }

        if ($flag1 -eq 1 -or $flag1 -eq 0) {
            $Result = "Relationship Deleted"
            $RelationshipDeleted++
            $table_html += "<tr style='border:1px solid black; padding-right:5px; padding-left:5px'>
                <td style='border:1px solid black;text-align: left'>$RowCount</td>
                <td style='border:1px solid black;text-align: left'>$Result10</td>
                <td style='border:1px solid black;text-align: left'>$Result11</td>
                <td style='border:1px solid black;text-align: left'>$Result12</td>
                <td style='border:1px solid black;text-align: left'>$Result13</td>
                <td style='border:1px solid black;text-align: left'>$Result14</td>
                <td style='border:1px solid black;text-align: left; background-color:#ffa6a6'>$Result</td>
            </tr>"
            $RowCount++
        }
    }

    $table_html += "</table>"
    $Summary = "<p><b>Relationship(s) Added:</b> $RelationshipAdded; &emsp; <b>Relationship(s) Deleted:</b> $RelationshipDeleted; &emsp; <b>Cardinality Changed:</b> $CardinalityChanged; &emsp; <b>Relationship Attribute Changed:</b> $RelationshipAttributeChanged</p><br><br>"
    return $Summary + $table_html
}


function MeasureDefinitionComparison {
    $FeatureName = "Measure Definition Comparison"

    $Query = "SELECT [MEASURE_NAME] AS [MeasureName], [DEFAULT_FORMAT_STRING] AS [MeasureFormat], [EXPRESSION] AS [MeasureDefinition], [MEASURE_IS_VISIBLE] AS [IsVisible] FROM `$SYSTEM`.MDSCHEMA_MEASURES WHERE CUBE_NAME  ='Model'"
    
    $MeasureDefinition_Processing = XMLToObject -response (invoke-ascmd -Database $ModelName_Processing -Query $Query -server $Workspace_Processing -ServicePrincipal -Tenant $tenantId -Credential $credential)
    $MeasureDefinition_Publish = XMLToObject -response (invoke-ascmd -Database $ModelName_Publish -Query $Query -server $Workspace_Publish -ServicePrincipal -Tenant $tenantId -Credential $credential)

    $RowCount_Processing, $RowCount_Publish = $MeasureDefinition_Processing.Count, $MeasureDefinition_Publish.Count
    $RowCount, $MeasureDefinitionUpdated, $MeasureFormatUpdated, $MeasureDeleted, $MeasureAdded, $MeasureVisibilityChanged = 1, 0, 0, 0, 0 ,0
    
    $table_html = "
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
            $Result20 = $MeasureDefinition_Publish[$j].C0
            $Result21 = $MeasureDefinition_Publish[$j].C1
            $Result22 = $MeasureDefinition_Publish[$j].C2
            $Result23 = $MeasureDefinition_Publish[$j].C3
    
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
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left""></td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left"">$Result</td>
            </tr>"
        }
        elseif ($flag3 -eq 1 -and $flag4 -eq 0) {
            $Result = "Measure Visibility Changed"
            $MeasureVisibilityChanged++
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left""></td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left; background-color:#FFB52E"">$Result</td>
            </tr>"
        }
        elseif ($flag1 -eq 1 -and $flag3 -eq 0) {
            $Result = "Definition Updated"
            $MeasureDefinitionUpdated++
            $OldResult = $MeasureDefinition_Publish[$i].C2
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left"">$OldResult</td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left; background-color:#FFB52E"">$Result</td>
            </tr>"
        }
        elseif ($flag1 -eq 1 -and $flag2 -eq 0) {
            $Result = "Format Updated"
            $MeasureFormatUpdated++
            $OldResult = $MeasureDefinition_Publish[$i].C1
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left"">$OldResult</td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left; background-color:#FFB52E"">$Result</td>
            </tr>"
        }
        elseif ($flag1 -eq 0) {
            $Result = "Measure Added"
            $MeasureAdded++
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left""></td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left; background-color:#B3EDA7"">$Result</td>
            </tr>"
        }
        $RowCount++
    }
    
    for ($i = 0; $i -lt $RowCount_Publish; $i++) {
        $flag1 = 0
        $Result10, $Result11, $Result12, $Result13 = $MeasureDefinition_Publish[$i].C0, $MeasureDefinition_Publish[$i].C1, $MeasureDefinition_Publish[$i].C2, $MeasureDefinition_Publish[$i].C3
        $Result20, $Result21, $Result22, $Result23 = '', '', '', ''
        
        for ($j = 0; $j -lt $RowCount_Processing; $j++) {
            $Result20 = $MeasureDefinition_Processing[$j].C0
            $Result21 = $MeasureDefinition_Processing[$j].C1
            $Result22 = $MeasureDefinition_Processing[$j].C2
            $Result23 = $MeasureDefinition_Processing[$j].C3
    
            if ($Result11 -eq $Result21) {
                $flag1 = 1 
            }
        }
    
        if ($flag1 -eq 0) {
            $Result = "Measure Deleted"
            $MeasureDeleted++
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left""></td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left; background-color:#ffa6a6"">$Result</td>
            </tr>"
            $RowCount++
        }
    }
    
    $table_html += "</table>"
    $Summary = "<p><b>Measure(s) Added:</b> $MeasureAdded; &emsp; <b>Measure(s) Deleted:</b> $MeasureDeleted; &emsp; <b>Definition Updated:</b> $MeasureDefinitionUpdated; &emsp; <b>Format Updated:</b> $MeasureFormatUpdated; &emsp; <b>Visibility Changed:</b> $MeasureVisibilityChanged</p><br><br>"
    return  $Summary + $table_html
}

if ([string]::IsNullOrEmpty($FeatureArray)) {

    Write-Host "Comparing Table Attributes"		
    $html_t = TableAttributeComparison    

    Write-Host "Comparing Relationships"
    $html_r = RelationshipComparison
	
    Write-Host "Comparing Measure Definition"
    $html_m = MeasureDefinitionComparison

    $html = $html_t + "<br>" + $html_r + "<br>" + $html_m
}
else {
    $Array = $FeatureArray.split(",")
    $html = ""
    foreach ($i in $Array) {
        if ($i -eq 1) {
            Write-Host "Comparing Measure values"
            $html_i = "MeasureValueComparison"
        }
        elseif ($i -eq 2) {
            Write-Host "Comparing Table Attributes"
            $html_i = TableAttributeComparison    
        }
        elseif ($i -eq 3) {
            Write-Host "Comparing Relationships"
            $html_i = RelationshipComparison
        }
        elseif ($i -eq 4) {
            Write-Host "Comparing Measure Definition"
            $html_i = MeasureDefinitionComparison
        }
        $html = $html + $html_i + "<br>"
    }
}


$Body = "<h2>Verificaton Summary: " + $FeatureName + "</h2>
<p><b>Processing Workspace:</b> $Workspace_Processing</p>
<p><b>Processing Dataset:</b> $ModelName_Processing</p>
<p><b>Publish Workspace:</b> $Workspace_Publish</p>
<p><b>Publish Dataset:</b> $ModelName_Publish</p>
"

$temp = $Body + $html
# $temp | Out-File -FilePath "File.txt"

$PartialFiltered = $temp.replace("'", "\'")
$Result = $PartialFiltered.replace('"', '`"')

# Assuming $Results is a DataTable

$Subject = "Power BI Dataset Comparison for " + $ModelName_Publish +" as on "+ $CurrentDate

function Send-ToEmail([string]$email, [string]$attachmentpath, [string]$body, [string]$subject){
    $message = new-object Net.Mail.MailMessage;
    $message.From = $SenderMailAccount;
    $message.To.Add($email);
    $message.Subject = $subject;
    $message.Body = $temp;
    $message.IsBodyHtml = $true;
    $smtp = new-object Net.Mail.SmtpClient("smtp.office365.com", "587");
    $smtp.EnableSSL = $true	;
    $smtp.Credentials = New-Object System.Net.NetworkCredential($SenderMailAccount, $SenderMailPassword);
    $smtp.send($message);
    write-host "Mail Sent" ; 
}
Send-ToEmail -email $MailReceipients -body $Result -subject $subject