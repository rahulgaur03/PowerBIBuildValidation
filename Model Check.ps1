##Arguments to be passed
# param(
# $Workspace_Processing,
# $Workspace_Publish,
# $ModelName_Processing,
# $ModelName_Publish,
# $FeatureArray,
# $MailReceipients,
# $SenderMailAccount,
# $SenderMailPassword,
# $PowerBILogin,
# $PowerBIPassword
# )

#Feature Number 1-Measure Value; 2-Schema; 3-Relationship; 4-Definition
Import-Module SqlServer
$tenantId = "39e4b8d8-5034-4760-84a3-0ba2ca73ac84"
$Workspace_Processing = "powerbi://api.powerbi.com/v1.0/myorg/Sprouts%20EDW%20-%20Test"
$Workspace_Publish = "powerbi://api.powerbi.com/v1.0/myorg/Sprouts%20EDW"
$ModelName_Processing = "Refresh Tracker"
$ModelName_Publish = "Refresh Tracker"
$FeatureArray = "1"
$MailReceipients = ""
$SenderMailAccount = ""
$SenderMailPassword = ""
$PowerBIPassword = ""
$PowerBILogin =""

$CurrentDate = Get-Date -Format "dddd MM/dd/yyyy"
$PowerBIEndpoint_Processing = "$Workspace_Processing;Initial catalog=$ModelName_Processing"
$PowerBIEndpoint_Publish = "$Workspace_Publish;Initial catalog=$ModelName_Publish"

$SecuredApplicationSecret = ConvertTo-SecureString -String "dwhibl7x7FC/IRE1eM5aeMEhSveYdfaKclCQCIaasGQ=" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential("ddb776a3-ea78-4d50-8e51-2aa140618647", $SecuredApplicationSecret)

function XMLToObject {
    param (
        $response
    )
    $xml = [xml]$response
    $ns = New-Object Xml.XmlNamespaceManager $xml.NameTable
    $ns.AddNamespace("ns", "urn:schemas-microsoft-com:xml-analysis:rowset")
    return $xml.SelectNodes('//ns:row', $ns)
}

function MeasureValueComparison {
    $FeatureName = "Measure Value Comparison"

    ## Query to get all the Measure Names from the Model
    $Query = "SELECT [MEASURE_CAPTION] AS [MEASURE] FROM `$SYSTEM`.MDSCHEMA_MEASURES WHERE CUBE_NAME = 'Model'"
    $response = invoke-ascmd -Database $ModelName_Processing -Query $Query -server $Workspace_Processing -ServicePrincipal -Tenant $tenantId -Credential $credential
    $Results  = XMLToObject -response $response

    ## Looping for each measure, firing it on the Model and getting the Value
    $MeasureOutput_Processing =  New-Object System.Collections.ArrayList
    $MeasureOutput_Publish = New-Object System.Collections.ArrayList

    foreach($Measure in $Results){
        Write-Host $Measure.C0
        $Query = "SELECT ["+ $Measure.C0 +"] ON 0 FROM [Model]"
        $response = invoke-ascmd -Database $ModelName_Processing -Query $Query -server $Workspace_Processing -ServicePrincipal -Tenant $tenantId -Credential $credential
        $measureValue = [xml]$response | Select-Xml -Namespace @{ ns = 'urn:schemas-microsoft-com:xml-analysis:mddataset' } -XPath '//ns:CellData/ns:Cell/ns:Value' | ForEach-Object { $_.Node.InnerText }
        $MeasureOutput_Processing.Add($measureValue)

        $Query = "SELECT ["+ $Measure.C0 +"] ON 0 FROM [Model]"
        $response = invoke-ascmd -Database $ModelName_Publish -Query $Query -server $Workspace_Publish -ServicePrincipal -Tenant $tenantId -Credential $credential
        $measureValue = [xml]$response | Select-Xml -Namespace @{ ns = 'urn:schemas-microsoft-com:xml-analysis:mddataset' } -XPath '//ns:CellData/ns:Cell/ns:Value' | ForEach-Object { $_.Node.InnerText }
        $MeasureOutput_Publish.Add($measureValue) 
    }
    
    ## HTML Code for Mail
    $table_html = " <table style=""width:100%; border:1px solid black; border-collapse: collapse;""><tr style=""font-size:15px; font-weight: bold; padding:5px; color: white; border:1px solid black; background-color:#468ABD""><th style=""border:1px solid black;text-align: right"">#</th><th style=""border:1px solid black"">Measure Name</th><th style=""border:1px solid black;
    text-align: right;"">Processing Model</th><th style=""border:1px solid black;text-align: right;"">Published Model</th><th style=""border:1px solid black;text-align: right;"">Variance %</th><th style=""border:1px solid black"">Status</th></tr>"
    
    ## Comparing the Values from both the Models
    $rowCount = 0
    for($i = 0; $i -lt $MeasureOutput_Processing.Count; $i++){
        $Status = 'Pass'
      
        if($MeasureOutput_Processing[$i].GetType().Name -eq 'Double' -or $MeasureOutput_Processing[$i].GetType().Name -eq 'Int64' -or $MeasureOutput_Processing[$i].GetType().Name -eq 'Int16'){
            $processing = [math]::Round($MeasureOutput_Processing[$i], 1)
            $publish = [math]::Round($MeasureOutput_Publish[$i], 1)
        }

        if ( $publish -eq 0 ) {
            $varianceprc = 0
        }
        else {
            $varianceprc = [math]::Round((($Processing - $Publish)/$Publish), 1)
        }
        
        
        if ( $processing -isnot [DBNull]) 
        {
           if ( $processing - $publish  -gt 0)
              {
                  $Status = 'Fail'
              }
          elseif ( $processing - $publish  -lt 0)
              {
                 $Status = 'Fail'
              }
        }
        else
        {
           $Status = 'Pass'
           $processing = 0
           $publish = 0
        }
        if ( $status -eq 'Fail' ) {
            $rowCount++
            $table_html += "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> <td style=""border:1px solid black;text-align: right"">$rowCount</td> <td style=""border:1px solid black"">$($Results.Rows[$i]["Measure"])</td> <td style=""border:1px solid black;text-align: right"">$processing</td> <td style=""border:1px solid black;text-align: right"">$publish</td> <td style=""border:1px solid black;text-align: right"">$varianceprc</td> <td style=""border:1px solid black; background-color:#ffa6a6"">$Status</td> </tr>"
        }
           
    }
    $totalCount = $MeasureOutput_Processing.Count;
    $passCount = $MeasureOutput_Processing.Count - $rowCount;
    $failCount = $rowCount;
    for($i = 0; $i -lt $MeasureOutput_Processing.Count; $i++){
        $Status = 'Pass'
        
        if($MeasureOutput_Processing[$i].GetType().Name -eq 'Double' -or $MeasureOutput_Processing[$i].GetType().Name -eq 'Int64' -or $MeasureOutput_Processing[$i].GetType().Name -eq 'Int16'){
            $processing = [math]::Round($MeasureOutput_Processing[$i], 1)
            $publish = [math]::Round($MeasureOutput_Publish[$i], 1)
        }

        if ( $publish -eq 0 ) {
            $varianceprc = 0
        }
        else {
            $varianceprc = [math]::Round((($Processing - $Publish)/$Publish), 1)
        }
        
        
        if ( $processing -isnot [DBNull]) 
        {
           if ( $processing - $publish  -gt 0)
              {
                  $Status = 'Fail'
              }
          elseif ( $processing - $publish  -lt 0)
              {
                 $Status = 'Fail'
              }
        }
        else
        {
           $Status = 'Pass'
           $processing = 0
           $publish = 0
        }
        if ( $status -eq 'Pass' ) {
          $rowCount++
          $table_html += "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> <td style=""border:1px solid black;text-align: right"">$rowCount</td> <td style=""border:1px solid black"">$($Results.Rows[$i]["Measure"])</td> <td style=""border:1px solid black;text-align: right"">$processing</td> <td style=""border:1px solid black;text-align: right"">$publish</td> <td style=""border:1px solid black;text-align: right"">$varianceprc</td> <td style=""border:1px solid black; background-color:#B3EDA7"">$Status</td> </tr>"
        }
           
    }
    
    $table_html += "</table>"

    $Summary = "
        <p><b>Total Measures:</b> $totalCount; &emsp; <b>Passed:</b> $passcount; &emsp; <b>Failed:</b> $failCount</p>
        <br><br>
        "
	return  $Summary + $table_html
}


function Table_AttributeComparison {
    $FeatureName = "Table/Attribute Comparison"

    $Query = "SELECT [CATALOG_NAME] AS [TabularName],
        [DIMENSION_UNIQUE_NAME] AS [DimensionName],
        HIERARCHY_CAPTION AS [AttributeName],
        HIERARCHY_IS_VISIBLE AS [IsAttributeVisible]
    FROM `$system`.MDSchema_hierarchies
    WHERE CUBE_NAME  ='Model'"
    
    $response = invoke-ascmd -Database $ModelName_Processing -Query $Query -server $Workspace_Processing -ServicePrincipal -Tenant $tenantId -Credential $credential
    $Dimension_Attribute_Processing = XMLToObject -response $response

    $response = invoke-ascmd -Database $ModelName_Publish -Query $Query -server $Workspace_Publish -ServicePrincipal -Tenant $tenantId -Credential $credential
    $Dimension_Attribute_Publish = XMLToObject -response $response
    
    $RowCount_Processing = $Dimension_Attribute_Processing.Count
    $RowCount_Publish = $Dimension_Attribute_Publish.Count
    $RowCount = 1
    $AttributeVisbilityChanged = 0
    $AttributeAdded = 0
    $AttributeDeleted = 0
    
    $table_html = "
        <table style=""width:100%; border:1px solid black; border-collapse: collapse;"">
            <tr style=""font-size:15px; font-weight: bold; padding:5px; color: white; border:1px solid black; background-color:#468ABD"">
                <th style=""border:1px solid black"">#</th>
                <th style=""border:1px solid black"">Table Name</th>
                <th style=""border:1px solid black"">Attribute Name</th>
                <th style=""border:1px solid black"">Is Visible</th>
                <th style=""border:1px solid black"">Status</th>
            </tr>"
    
    for ($i = 0; $i -lt $RowCount_Processing; $i++) {
        $flag1 = 0
        $flag2 = 0
        $flag3 = 0
        $Result11 = $Dimension_Attribute_Processing[$i].C1
        $Result21 = ''
        $Result12 = $Dimension_Attribute_Processing[$i].C2
        $Result22 = ''
        $Result13 = $Dimension_Attribute_Processing[$i].C3
        $Result23 = ''
    
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
            $Result = "No Change"
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td> 
                <td style=""border:1px solid black;text-align: left"">$Result13</td> 
                <td style=""border:1px solid black;text-align: left"">$Result</td>
            </tr>"
        }
        elseif ($flag2 -eq 1 -and $flag3 -eq 0) {
            $Result = "Attribute Visbility Changed"
            $AttributeVisbilityChanged++
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td> 
                <td style=""border:1px solid black;text-align: left"">$Result13</td> 
                <td style=""border:1px solid black;text-align: left; background-color:#FFB52E"">$Result</td>
            </tr>"
        }
        elseif ($flag1 -eq 1 -and $flag2 -eq 0) {
            $Result = "Attribute Added"
            $AttributeAdded++
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td> 
                <td style=""border:1px solid black;text-align: left"">$Result13</td> 
                <td style=""border:1px solid black;text-align: left; background-color:#B3EDA7"">$Result</td>
            </tr>"
        }
        elseif ($flag1 -eq 0) {
            $Result = "Table/Attribute Added"
            $AttributeAdded++
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td> 
                <td style=""border:1px solid black;text-align: left"">$Result13</td> 
                <td style=""border:1px solid black;text-align: left; background-color:#B3EDA7"">$Result</td>
            </tr>"
        }
        $RowCount++
    }
    
    for ($i = 0; $i -lt $RowCount_Publish; $i++) {
        $flag1 = 0
        $flag2 = 0
        $flag3 = 0
        $a = "C"+"$i"
        $Result11 = $Dimension_Attribute_Publish[$i].C1
        $Result21 = ''
        $Result12 = $Dimension_Attribute_Publish[$i].C2
        $Result22 = ''
        $Result13 = $Dimension_Attribute_Publish[$i].C3
        $Result23 = ''
    
        for ($j = 0; $j -lt $RowCount_Processing; $j++) {
            $Result21 = $Dimension_Attribute_Processing[$j].C1
            $Result22 = $Dimension_Attribute_Processing[$j].C2
            $Result23 = $Dimension_Attribute_Processing[$j].C3
    
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
    
        if ($flag1 -eq 1 -and $flag2 -eq 0) {
            $Result = "Attribute Deleted"
            $AttributeDeleted++
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td> 
                <td style=""border:1px solid black;text-align: left"">$Result13</td> 
                <td style=""border:1px solid black;text-align: left; background-color:#ffa6a6"">$Result</td>
            </tr>"
        }
        elseif ($flag1 -eq 0) {
            $Result = "Table/Attribute Deleted"
            $AttributeDeleted++
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td> 
                <td style=""border:1px solid black;text-align: left"">$Result13</td> 
                <td style=""border:1px solid black;text-align: left; background-color:#ffa6a6"">$Result</td>
            </tr>"
        }
        $RowCount++
    }
    
    $table_html += "</table>"
    $Summary = "
        <p><b>Attribute(s) Added:</b> $AttributeAdded; &emsp; <b>Attribute(s) Deleted:</b> $AttributeDeleted; &emsp; <b>Attribute Visibility Updated:</b> $AttributeVisbilityChanged</p>
        <br><br>
        "
    return  $Summary + $table_html
}

function RelationshipComparison {
    $FeatureName = "Relationship Comparison"

    $Query = "SELECT [MEASUREGROUP_NAME] AS [LeftHand],
        [DIMENSION_UNIQUE_NAME] AS [RightHand],
        [MEASUREGROUP_CARDINALITY] AS [LeftHandCard],      
        [DIMENSION_CARDINALITY] AS [RightHandCard],
        [DIMENSION_GRANULARITY] AS [RightHandKey]
    FROM `$system`.MDSCHEMA_MEASUREGROUP_DIMENSIONS
    WHERE [CUBE_NAME] ='Model'"

    $response = invoke-ascmd -Database $ModelName_Processing -Query $Query -server $Workspace_Processing -ServicePrincipal -Tenant $tenantId -Credential $credential
    $Relationship_Processing = XMLToObject -response $response

    $response = invoke-ascmd -Database $ModelName_Publish -Query $Query -server $Workspace_Publish -ServicePrincipal -Tenant $tenantId -Credential $credential
    $Relationship_Publish = XMLToObject -response $response
    
    $RowCount_Processing = $Relationship_Processing.Count
    $RowCount_Publish = $Relationship_Publish.Count
    $RowCount = 1
    $RelationshipAdded = 0
    $RelationshipDeleted = 0
    $CardinalityChanged = 0
    $RelationshipAttributeChanged = 0
    
    $table_html = "
        <table style=""width:100%; border:1px solid black; border-collapse: collapse;"">
            <tr style=""font-size:15px; font-weight: bold; padding:5px; color: white; border:1px solid black; background-color:#468ABD"">
                <th style=""border:1px solid black"">#</th>
                <th style=""border:1px solid black"">Fact Name</th>
                <th style=""border:1px solid black"">Dim Name</th>
                <th style=""border:1px solid black"">Fact Cardinality</th>
                <th style=""border:1px solid black"">Dim Cardinality</th>
                <th style=""border:1px solid black"">Relationship On Column</th>
                <th style=""border:1px solid black"">Status</th>
            </tr>"
    
    for ($i = 0; $i -lt $RowCount_Processing; $i++) {
        $flag1 = 0
        $flag2 = 0
        $flag3 = 0
        $flag4 = 0
        $flag5 = 0
        $Result10 = $Relationship_Processing[$i].C0
        $Result20 = ''
        $Result11 = $Relationship_Processing[$i].C1
        $Result21 = ''
        $Result12 = $Relationship_Processing[$i].C2
        $Result22 = ''
        $Result13 = $Relationship_Processing[$i].C3
        $Result23 = ''
        $Result14 = $Relationship_Processing[$i].C4
        $Result24 = ''
    
        for ($j = 0; $j -lt $RowCount_Publish; $j++) {
            $Result20 = $Relationship_Publish[$j].C0
            $Result21 = $Relationship_Publish[$j].C1
            $Result22 = $Relationship_Publish[$j].C2
            $Result23 = $Relationship_Publish[$j].C3
            $Result24 = $Relationship_Publish[$j].C4
    
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
            $Result = "No Change"
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left"">$Result14</td> 
                <td style=""border:1px solid black;text-align: left"">$Result</td>
            </tr>"
        }
        elseif ($flag4 -eq 1 -and $flag5 -eq 0) {
            $Result = "Relationship Attribute Changed"
            $RelationshipAttributeChanged++
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left"">$Result14</td> 
                <td style=""border:1px solid black;text-align: left; background-color:#FFB52E"">$Result</td>
            </tr>"
        }
        elseif (($flag3 -eq 1 -and $flag4 -eq 0) -or ($flag2 -eq 1 -and $flag3 -eq 0)) {
            $Result = "Cardinality Changed"
            $CardinalityChanged++
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left"">$Result14</td> 
                <td style=""border:1px solid black;text-align: left; background-color:#FFB52E"">$Result</td>
            </tr>"
        }
        elseif ($flag1 -eq 1 -and $flag2 -eq 0) {
            $Result = "Relationship Added"
            $RelationshipAdded++
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left"">$Result14</td> 
                <td style=""border:1px solid black;text-align: left; background-color:#B3EDA7"">$Result</td>
            </tr>"
        }
        $RowCount++
    }
    
    for ($i = 0; $i -lt $RowCount_Publish; $i++) {
        $flag1 = 0
        $flag2 = 0
        $Result10 = $Relationship_Publish[$i].C0
        $Result20 = ''
        $Result11 = $Relationship_Publish[$i].C1
        $Result21 = ''
        $Result12 = $Relationship_Publish[$i].C2
        $Result22 = ''
        $Result13 = $Relationship_Publish[$i].C3
        $Result23 = ''
        $Result14 = $Relationship_Publish[$i].C4
        $Result24 = ''
    
        for ($j = 0; $j -lt $RowCount_Processing; $j++) {
            $Result20 = $Relationship_Processing[$j].C0
            $Result21 = $Relationship_Processing[$j].C0
            $Result22 = $Relationship_Processing[$j].C0
            $Result23 = $Relationship_Processing[$j].C0
            $Result24 = $Relationship_Processing[$j].C0
    
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
            $table_html += 
            "<tr style=""border:1px solid black; padding-right:5px; padding-left:5px""> 
                <td style=""border:1px solid black;text-align: left"">$RowCount</td> 
                <td style=""border:1px solid black;text-align: left"">$Result10</td> 
                <td style=""border:1px solid black;text-align: left"">$Result11</td> 
                <td style=""border:1px solid black;text-align: left"">$Result12</td>
                <td style=""border:1px solid black;text-align: left"">$Result13</td>
                <td style=""border:1px solid black;text-align: left"">$Result14</td>
                <td style=""border:1px solid black;text-align: left; background-color:#ffa6a6"">$Result</td>
            </tr>"
        
            $RowCount++
        }
    }
    
    $table_html += "</table>"
    $Summary = "
        <p><b>Relationship(s) Added:</b> $RelationshipAdded; &emsp; <b>Relationship(s) Deleted:</b> $RelationshipDeleted; &emsp; <b>Cardinality Changed:</b> $CardinalityChanged; &emsp; <b>Relationship Attribute Changed:</b> $RelationshipAttributeChanged</p>
        <br><br>
        "
    return  $Summary + $table_html
}

function MeasureDefinitionComparison {
    $FeatureName = "Measure Definition Comparison"

    $Query = "SELECT [MEASURE_NAME] AS [MeasureName],
        [DEFAULT_FORMAT_STRING] AS [MeasureFormat],
        [EXPRESSION] AS [MeasureDefinition],
        [MEASURE_IS_VISIBLE] AS [IsVisible]
    FROM `$SYSTEM`.MDSCHEMA_MEASURES
    WHERE CUBE_NAME  ='Model'"
    
    $response = invoke-ascmd -Database $ModelName_Processing -Query $Query -server $Workspace_Processing -ServicePrincipal -Tenant $tenantId -Credential $credential
    $MeasureDefinition_Processing = XMLToObject -response $response

    $response = invoke-ascmd -Database $ModelName_Publish -Query $Query -server $Workspace_Publish -ServicePrincipal -Tenant $tenantId -Credential $credential
    $MeasureDefinition_Publish = XMLToObject -response $response

    
    $RowCount_Processing = $MeasureDefinition_Processing.Count
    $RowCount_Publish = $MeasureDefinition_Publish.Count
    $RowCount = 1
    $MeasureDefinitionUpdated = 0
    $MeasureFormatUpdated = 0
    $MeasureDeleted = 0
    $MeasureAdded = 0
    $MeasureVisibilityChanged = 0
    
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
        $flag1 = 0
        $flag2 = 0
        $flag3 = 0
        $flag4 = 0
        $Result10 = $MeasureDefinition_Processing[$i].C0
        $Result20 = ''
        $Result11 = $MeasureDefinition_Processing[$i].C1
        $Result21 = ''
        $Result12 = $MeasureDefinition_Processing[$i].C2
        $Result22 = ''
        $Result13 = $MeasureDefinition_Processing[$i].C3
        $Result23 = ''
    
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
        $Result10 = $MeasureDefinition_Publish[$i].C0
        $Result20 = ''
        $Result11 = $MeasureDefinition_Publish[$i].C1
        $Result21 = ''
        $Result12 = $MeasureDefinition_Publish[$i].C2
        $Result22 = ''
        $Result13 = $MeasureDefinition_Publish[$i].C3
        $Result23 = ''
    
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
    $Summary = "
        <p><b>Measure(s) Added:</b> $MeasureAdded; &emsp; <b>Measure(s) Deleted:</b> $MeasureDeleted; &emsp; <b>Definition Updated:</b> $MeasureDefinitionUpdated; &emsp; <b>Format Updated:</b> $MeasureFormatUpdated; &emsp; <b>Visibility Changed:</b> $MeasureVisibilityChanged</p>
        <br><br>
        "
    return  $Summary + $table_html
}

if ([string]::IsNullOrEmpty($FeatureArray)) {
    Write-Host "Comparing Measure values"
    $html = MeasureValueComparison
	
    Write-Host "Comparing Table Attributes"		
    $html = Table_AttributeComparison    

    Write-Host "Comparing Relationships"
    $html = RelationshipComparison
	
    Write-Host "Comparing Measure Definition"
    $html = MeasureDefinitionComparison
}
else {
    $Array = $FeatureArray.split(",")
    $html = ""
    foreach ($i in $Array) {
        if ($i -eq 1) {
            Write-Host "Comparing Measure values"
            $html_i = MeasureValueComparison
        }
        elseif ($i -eq 2) {
            Write-Host "Comparing Table Attributes"
            $html_i = Table_AttributeComparison    
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


Write-Host 'Connnections Closed'

$Body = "<h2>Verificaton Summary: " + $FeatureName + "</h2>
<p><b>Processing Workspace:</b> $Workspace_Processing</p>
<p><b>Processing Dataset:</b> $ModelName_Processing</p>
<p><b>Publish Workspace:</b> $Workspace_Publish</p>
<p><b>Publish Dataset:</b> $ModelName_Publish</p>
"

$temp = $Body + $html
$temp | Out-File -FilePath "File.txt"

$PartialFiltered = $temp.replace("'", "\'")
$Result = $PartialFiltered.replace('"', '`"')

# Assuming $Results is a DataTable

# $Username = $SenderMailAccount;
# $Subject = "Power BI Dataset Comparison for " + $ModelName_Publish +" as on "+ $CurrentDate

# function Send-ToEmail([string]$email, [string]$attachmentpath, [string]$body, [string]$subject){

#     $message = new-object Net.Mail.MailMessage;
#     $message.From = $SenderMailAccount;
#     $message.To.Add($email);
#     $message.Subject = $subject;
#     $message.Body = $temp;
#     $message.IsBodyHtml = $true;
#     $password = $SenderMailPassword
#     $smtp = new-object Net.Mail.SmtpClient("smtp.office365.com", "587");
#     $smtp.EnableSSL = $true	;
#     $smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $Password);
#     $smtp.send($message);
#     write-host "Mail Sent" ; 
# }

# Send-ToEmail -email $MailReceipients -body $Result -subject $subject
