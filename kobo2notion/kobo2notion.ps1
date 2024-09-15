$ErrorActionPreference = "Stop"

function Install-ModuleIfMissing {
    param(
        [string]$moduleName
    )

    $module = Get-Module -ListAvailable -Name $moduleName

    if ($null -eq $module) {
        Write-Host "Module $moduleName is not installed."
        Write-Host "Auto installing $moduleName..."
        Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
    
        $module = Get-Module -ListAvailable -Name $moduleName
        if ($null -eq $module) {
            Write-Host "Module $moduleName install failed。" -ForegroundColor Red
            exit 1
        }
        else {
            Write-Host "Module $moduleName install success。"
        }
    }
}

function Read-Config() {
    param(
        [string]$configPath
    )
    $properties = [PSCustomObject]@{}
    #$properties = @{}

    if (Test-Path -Path $configPath) {
        Write-Host "Load config file: $configPath"
    
        $config = Get-Content -Path $configPath
    
        foreach ($line in $config) {
            $words = $line.Split('=', 2)
            #$properties.add($words[0].Trim(), $words[1].Trim())
            $properties | Add-Member -MemberType NoteProperty -Name $words[0].Trim() -Value $words[1].Trim()
        }

        if (-not $properties.notionToken -or -not $properties.notionDatabaseId -or -not $properties.koboDatabasePath) {
            Write-Host "Error: 'notionToken' or 'notionDatabaseId' or 'koboDatabasePath' is missing or empty in the config file." -ForegroundColor Red
            exit 1
        }
    }
    else {
        Write-Host "Not found config file: $configPath" -ForegroundColor Red
        exit 1
    }

    return $properties
}

function Get-CurrentDateTime {
    $TimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Taipei Standard Time")
    $DateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), $TimeZone.Id)

    return $DateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz")
}

$configPath = Join-Path -Path $PSScriptRoot -ChildPath "config.txt"

Install-ModuleIfMissing -moduleName "PSSQLite"
$properties = Read-Config -configPath $configPath

# Create Notion API headers
$headers = @{
    "Authorization"  = "Bearer $($properties.notionToken)"
    "Content-Type"   = "application/json"
    "Notion-Version" = "2022-06-28"
}


# Query book list
$getBookListQuery = @"
    SELECT DISTINCT content.ContentId, content.Title, content.Subtitle, content.Attribution AS Author, content.Publisher, content.ISBN 
    FROM Bookmark INNER JOIN content 
    ON Bookmark.VolumeID = content.ContentID 
    ORDER BY content.Title
"@
$bookList = Invoke-SqliteQuery -Query $getBookListQuery -DataSource $properties.koboDatabasePath

foreach ($book in $bookList) {
    try {
        $title = $book.Title

        Write-Host "Query notion database id: $($properties.notionDatabaseId)"
        # Query Notion book
        $notionQueryUrl = "https://api.notion.com/v1/databases/$($properties.notionDatabaseId)/query"
        $filter = @{
            filter = @{
                and = @(
                    @{ property = "Title"; title = @{ contains = $title } }
                    @{ property = "Synced"; checkbox = @{ equals = $false } }
                )
            }
        }
        
        $filterJsonBody = $filter | ConvertTo-Json -Depth 10
        Write-Debug $notionQueryUrl
        Write-Debug $filterJsonBody

        $response = Invoke-RestMethod -Uri $notionQueryUrl -Method Post -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($filterJsonBody))
        Write-Debug ($response | ConvertTo-Json -Depth 10)
        
        # Determine whether the corresponding book is found
        $valid = $false
        if ($response.results.Count -eq 1) {
            $valid = $true
        }
        elseif ($response.results.Count -gt 1) {
            Write-Host "$title matched multiple items."
        }
        else {
            Write-Host "$title was skipped."
        }

        if ($valid) {
            $pageId = $response.results[0].id
            $blocks = @()

            # Find all blocks in this page
            $getBlocksUrl = "https://api.notion.com/v1/blocks/$pageId/children"
            $response = Invoke-RestMethod -Uri $getBlocksUrl -Method Get -Headers $headers
            $existingBlocks = $response.results

            # Filter bulleted_list_item and heading_2 block
            $highlightBlocks = $existingBlocks | Where-Object { $_.type -eq 'bulleted_list_item' -or $_.type -eq 'heading_2' }

            # Delete all highlight blocks
            foreach ($block in $highlightBlocks) {
                $deleteBlockUrl = "https://api.notion.com/v1/blocks/$($block.id)"
                $response = Invoke-RestMethod -Uri $deleteBlockUrl -Method Delete -Headers $headers
                Write-Debug $response
                Write-Host "Delete block success, blockId: $($block.id), pageId: $pageId"
            }

            # Query all highlights of this book
            $getHighlightsQuery = @"
                SELECT Bookmark.Text 
                FROM Bookmark INNER JOIN content 
                ON Bookmark.VolumeID = content.ContentID 
                WHERE content.ContentID = '$($book.ContentID)' 
                AND Bookmark.Text IS NOT NULL AND Bookmark.Text != ""
                ORDER BY content.DateCreated DESC
"@

            $highlightsList = Invoke-SqliteQuery -Query $getHighlightsQuery -DataSource $properties.koboDatabasePath

            # Create heading2: Highlights
            $blocks += @{
                object    = "block"
                type      = "heading_2"
                heading_2 = @{
                    rich_text = @(
                        @{
                            type = "text"
                            text = @{
                                content = "Highlights"
                            }
                        }
                    )
                }
            }
            
            # Create list item for each highligt
            foreach ($highlight in $highlightsList) {
                if (-not [string]::IsNullOrWhiteSpace($highlight.Text)) {
                    $blocks += @{
                        type               = "bulleted_list_item"
                        bulleted_list_item = @{
                            rich_text = @(
                                @{
                                    type = "text"
                                    text = @{
                                        content = $highlight.Text
                                    }
                                }
                            )
                        }
                    }
                }
            }

            # add highlight to notion page
            $appendUrl = "https://api.notion.com/v1/blocks/$pageId/children"
            $blocksJsonBody = @{
                children = $blocks
            } | ConvertTo-Json -Depth 10
            
            Write-Debug $appendUrl
            Write-Debug $blocksJsonBody
            
            $response = Invoke-RestMethod -Uri $appendUrl -Method Patch -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($blocksJsonBody))
            Write-Debug $response

            Write-Host "Add highlights to notion page success, pageId: $pageId"

            # update notion page highlight checkbox status
            $updateUrl = "https://api.notion.com/v1/pages/$pageId"
            $updateJsonBody = @{
                properties = @{
                    Synced            = @{
                        checkbox = $true
                    }
                    Publisher         = @{
                        "rich_text" = @(
                            @{
                                type       = "text"
                                text       = @{
                                    content = $book.Publisher
                                }
                                plain_text = $book.Publisher
                            }
                        )
                    }
                    Author            = @{
                        "rich_text" = @(
                            @{
                                type       = "text"
                                text       = @{
                                    content = $book.Author
                                }
                                plain_text = $book.Author
                            }
                        )
                    }
                    ISBN              = @{
                        "rich_text" = @(
                            @{
                                type       = "text"
                                text       = @{
                                    content = $book.ISBN
                                }
                                plain_text = $book.ISBN
                            }
                        )
                    }
                    "Highlight Count" = @{
                        number = $highlightsList.Count
                    }
                    "Sync Date"       = @{
                        date = @{
                            start = Get-CurrentDateTime
                        }
                    }
                }
            } | ConvertTo-Json -Depth 10

            Write-Debug $updateUrl
            Write-Debug $updateJsonBody

            $response = Invoke-RestMethod -Uri $updateUrl -Method Patch -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($updateJsonBody))
            Write-Debug $response
            
            Write-Host "Update notion page highlight checkbox status success, pageId: $pageId"

            Write-Host "Uploaded highlights for $title."
        }
    }
    catch {
        Write-Host "Error with $($book.Title): $_"
    }
}
