# Program Name: Power BI Dataset Comparison
# Date: 01/14/2024
# Author: V-RGAUR
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
# 01/14/2024   V-RGAUR        Initial script creation for Power BI dataset feature comparison and automated mail generation.

param(
    $TenantID = "39e4b8d8-5034-4760-84a3-0ba2ca73ac84",
    $ClientID = "ddb776a3-ea78-4d50-8e51-2aa140618647",
    $ClientSecret = "dwhibl7x7FC/IRE1eM5aeMEhSveYdfaKclCQCIaasGQ=",
    $SenderMailAccount = "svc-edwemailtest@sproutsfm.onmicrosoft.com",
    $SenderMailPassword = 'fK)1Lfy69I$toSUlTDrS',
    $DatasetName = "Refresh Tracker",
    $TestWorkspaceName = "Sprouts EDW - Test",
    $ProdWorkspaceName = "Sprouts EDW",
    $Workspace_Processing = "powerbi://api.powerbi.com/v1.0/myorg/$TestWorkspaceName",
    $Workspace_Publish = "powerbi://api.powerbi.com/v1.0/myorg/$ProdWorkspaceName",
    $FeatureArray = "2",
    $ShowChangesOnly = "0",
    $MailReceipients = "v-rgaur@sprouts.com"
)

Import-Module SqlServer
Import-Module -Name ImportExcel

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
    return ""
}

function TableAttributeComparison {
    Write-Host "Comparing Table Attributes"	
    $Query = "SELECT [CATALOG_NAME] AS [TabularName], [DIMENSION_UNIQUE_NAME] AS [DimensionName], HIERARCHY_CAPTION AS [AttributeName], HIERARCHY_IS_VISIBLE AS [IsAttributeVisible] FROM `$system`.MDSchema_hierarchies WHERE CUBE_NAME  ='Model'"

    $Dimension_Attribute_Processing = QueryExecutor -Server "Processing" -Query $Query
    $Dimension_Attribute_Publish = QueryExecutor -Server "Publish" -Query $Query

    $AttributeVisbilityChanged, $AttributeAdded, $AttributeDeleted, $RowCount = 0, 0, 0, 1
    $RowCount_Processing, $RowCount_Publish = $Dimension_Attribute_Processing.Count, $Dimension_Attribute_Publish.Count

    $ExcelPath = "ModelComparisonData.xlsx"

    $excelData = @()

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

        # Add data to the $excelData array
        $excelData += [PSCustomObject]@{
            'Table Name' = $Result11
            'Attribute Name' = $Result12
            'Is Visible' = $Result13
            'Status' = $Result
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
            $Result = if ($flag1 -eq 1 -and $flag2 -eq 0) { "Attribute" } else { "Table/Attribute" }
            $AttributeDeleted++
            
            # Add data to the $excelData array
            $excelData += [PSCustomObject]@{
                'Table Name' = $Result11
                'Attribute Name' = $Result12
                'Is Visible' = $Result13
                'Status' = "$Result Deleted"
            }
        }
        
        $RowCount++
    }

    # Export to Excel
    $excelData | Export-Excel -Path $ExcelPath -WorksheetName 'TableAttributes1' -AutoSize -BoldTopRow
    $excelData | Export-Excel -Path $ExcelPath -WorksheetName 'TableAttributes2' -AutoSize -BoldTopRow
    $excelData | Export-Excel -Path $ExcelPath -WorksheetName 'TableAttributes3' -AutoSize -BoldTopRow
}

# Call the function
TableAttributeComparison