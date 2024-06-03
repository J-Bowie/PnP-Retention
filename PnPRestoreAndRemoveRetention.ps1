####################################
# MODIFY THESE VARIABLES
####################################
$AdminUrl = ""      # Admin URL for SharePoint Online GEO Region / TL Admin Site
$label = ""         # Label Name to Remove from Files
$csvFilePath = ""   # Temporary CSV file to save progress
####################################

function Split-Collection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object[]] $InputObject,
        [Parameter(Position = 0)]
        [ValidateRange(1, [int]::MaxValue)]
        [int] $ChunkSize = 5
    )
    begin {
        $list = [System.Collections.Generic.List[object]]::new()
    }

    process {
        foreach($item in $InputObject) {
            $list.Add($item)
            if($list.Count -eq $ChunkSize) {
                $PSCmdlet.WriteObject($list.ToArray())
                $list.Clear()
            }
        }
    }
    end {
        if($list.Count) {
            $PSCmdlet.WriteObject($list.ToArray())
        }
    }
}

Connect-PnPOnline -Url $AdminUrl -Interactive
$Sites = Get-PnPTenantSite
$AllSites = @()
$countSites = 0;

if (-Not (Test-Path -Path $csvFilePath)) {
    $header = "Site URL"
    $header | Out-File -FilePath $csvFilePath -Encoding UTF8
}

#Keep track of which SPO sites have been completed
$completedSites = Import-Csv -Path $csvFilePath | Select-Object -ExpandProperty "Site URL"

$TotalSites = $Sites.Count - $completedSites.Count

ForEach($Site in $Sites)
{
    if ($Site.Url -notin $completedSites) {
        $countSites++;
        Write-Progress -activity "Processing $($Site.Url)" -status "$countSites out of $TotalSites completed"

        Try
        {      
            Connect-PnPOnline -Url $Site.Url -Interactive
            $ctx = Get-PnPContext
            $DeletedItemsBySystemAccount = Get-PnPRecycleBinItem | Where-Object {$_.DeletedByName -eq "System Account"}
            Write-Host "Obtained $($DeletedItemsBySystemAccount.Count) files deleted by System Account on $($Site.Url)"
            if ($DeletedItemsBySystemAccount.Count -gt 0) {
                $ids = [string]::Empty
                $i = 0
                foreach ($item in $DeletedItemsBySystemAccount) {
                    # Recycle restoration logic
                    $ids += '"' + $item.Id + '",'
                    $i++
                    if ($i -gt 200) {
                        $JSON_Restore = "{""ids"":[$($ids.TrimEnd(","))]}"
                        Invoke-PnPSPRestMethod -Method Post -Url "$($ctx.Url)/_api/site/RecycleBin/RestoreByIds" -Content $JSON_Restore
                        Write-Host "     Created JSON Payload to Restore files from Recycle Bin: $($_.Count) files" -ForegroundColor Yellow
                        $ids = [string]::Empty
                        $i = 0
                        Start-Sleep -Seconds 1
                    }
                }
                # Remainder handler
                if ($i -gt 0) {
                    $JSON_Restore = "{""ids"":[$($ids.TrimEnd(","))]}"
                    Invoke-PnPSPRestMethod -Method Post -Url "$($ctx.Url)/_api/site/RecycleBin/RestoreByIds" -Content $JSON_Restore
                    Write-Host "     Created JSON Payload to Restore files from Recycle Bin: $($_.Count) files" -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                }
            }

            $AllSites = (Get-PnPSubWeb -Recurse -IncludeRootWeb).Url

            foreach ($ss in $AllSites) {
                Write-Host "Getting information for Site: $($Site.Url)"
                Connect-PnPOnline -Url $ss -Interactive

                try {
                    $ctx = Get-PnPContext
                    $DocLibs = Get-PnPList | Where-Object {$_.BaseType -eq "DocumentLibrary"}

                    foreach($lib in $DocLibs)
                    {
                        try
                        {
                            $items = Get-PnPListItem -List $lib.Id -PageSize 5000 | Where-Object { $_.FieldValues._ComplianceTag -eq $label} | Select-Object
                            if($items.Count -gt 0)
                            {
                                Write-Host "   $($items.Count) items with $label label found in Document Library: " -ForegroundColor Magenta            

                                $items.Id | Split-Collection 200 | ForEach-Object {
                                    $JSON ="{
                                        ""blockDelete"": false,
                                        ""blockEdit"": false,
                                        ""complianceTagValue"": """",
                                        ""itemIds"": [
                                            $($_ -join "","")
                                        ],
                                        ""listUrl"": ""$($lib.RootFolder.ServerRelativeUrl)""
                                    }"
                                    Write-Host "     Created JSON Payload to Remove Label: $label from $($_.Count) files" -ForegroundColor Yellow

                                    Invoke-PnPSPRestMethod -Method Post -Url "$($ctx.Url)/_api/SP.CompliancePolicy.SPPolicyStoreProxy.ApplyLabelOnBulkItems()" -ContentType "application/json;odata=verbose" -Content $JSON | Select-Object StatusCode
                                }
                            }
                            else
                            {
                                Write-Host "   No items with $label label found in $lib." -ForegroundColor Green
                            }
                        }
                        catch
                        {
                            Write-Host  "Error: $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                } catch [System.Net.WebException] {
                    if ($_.Exception.Response.StatusCode -eq 429) {
                        Start-Sleep -Seconds 1
                    }
                }
            }
            Add-Content -Path $csvFilePath -Value $Site.Url
        }
        catch{
            Write-Host "Error occured $($Site.Url) : $_.Exception.Message"   -Foreground Red;
        }

    }
}
