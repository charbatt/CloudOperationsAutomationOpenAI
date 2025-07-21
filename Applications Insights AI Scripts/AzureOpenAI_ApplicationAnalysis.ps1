# =============================================================================
# Azure OpenAI Application Analysis Script with Intelligent Alert Creation - COMPLETE
# Analyzes 30 days of Application Insights telemetry and creates a detailed optimization report in HTML format and intelligent alerts
# Created by Charlie - If you use in your org please give me credit :D
# =============================================================================

# =============================================================================
# CONFIGURATION - Update these variables for your environment
# =============================================================================

# Azure OpenAI Configuration
$azureOpenAIEndpoint = "API EndPoint"
$azureOpenAIKey = "API KEY"
$deploymentName = "GPT Version Here"

# Application Insights Configuration
$appInsightsResourceId = "Resource ID Here"

# Web Application Configuration
$webAppSubscriptionId = "Subscription ID"
$webAppResourceGroup = "Resource Group Here"
$applicationName = "App Name Here"

# Analysis Configuration - 30 days for pattern detection
$analysisTimeRange = [TimeSpan]::FromDays(30)

# HTML Report Configuration
$reportOutputPath = "C:\Reports\ApplicationAnalysis_$applicationName_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

# Alert Configuration
$alertResourceGroup = "Resource Group for alerting"
$alertActionGroup = "Set Action Group"

# =============================================================================
# REQUIRED MODULES AND FUNCTIONS
# =============================================================================

# Import required modules
Write-Host "Loading required Azure modules..." -ForegroundColor Yellow
Import-Module Az.ApplicationInsights -Force
Import-Module Az.Accounts -Force
Import-Module Az.Websites -Force
Import-Module Az.Monitor -Force
Write-Host "✓ Modules loaded successfully" -ForegroundColor Green

# Function to call Azure OpenAI API
function Invoke-AzureOpenAI {
    param(
        [string]$Prompt,
        [string]$SystemMessage = "You are an expert application performance analyst.",
        [int]$MaxTokens = 2000
    )
    
    $headers = @{
        "Content-Type" = "application/json"
        "api-key" = $azureOpenAIKey
    }
    
    $body = @{
        messages = @(
            @{
                role = "system"
                content = $SystemMessage
            },
            @{
                role = "user"
                content = $Prompt
            }
        )
        max_tokens = $MaxTokens
        temperature = 0.3
    } | ConvertTo-Json -Depth 10
    
    $uri = "$azureOpenAIEndpoint/openai/deployments/$deploymentName/chat/completions?api-version=2024-02-15-preview"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body
        return $response.choices[0].message.content
    }
    catch {
        Write-Error "Azure OpenAI API call failed: $($_.Exception.Message)"
        return "Azure OpenAI analysis unavailable due to API error."
    }
}

# Function to execute KQL query using REST API
function Invoke-AppInsightsQuery {
    param(
        [string]$Query,
        [string]$AppId,
        [object]$AccessToken
    )
    
    try {
        Write-Host "    Executing query via REST API..." -ForegroundColor Gray
        
        $headers = @{
            "Authorization" = "Bearer $($AccessToken.Token)"
            "Content-Type" = "application/json"
        }
        
        $body = @{
            query = $Query
            timespan = "P30D"  # 30 days
        } | ConvertTo-Json
        
        $uri = "https://api.applicationinsights.io/v1/apps/$AppId/query"
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body
        
        Write-Host "    ✓ Query returned $($response.tables[0].rows.Count) records" -ForegroundColor Gray
        
        # Convert response to objects
        if ($response.tables -and $response.tables[0].rows) {
            $columns = $response.tables[0].columns.name
            $results = @()
            
            foreach ($row in $response.tables[0].rows) {
                $obj = New-Object PSObject
                for ($i = 0; $i -lt $columns.Count; $i++) {
                    $obj | Add-Member -MemberType NoteProperty -Name $columns[$i] -Value $row[$i]
                }
                $results += $obj
            }
            
            return $results
        } else {
            return @()
        }
    }
    catch {
        Write-Error "Application Insights query failed: $($_.Exception.Message)"
        return @()
    }
}

# Function to safely convert data to JSON for AI analysis
function Convert-DataForAI {
    param(
        [object]$Data,
        [int]$MaxItems = 50
    )
    
    if ($Data -and $Data.Count -gt 0) {
        $limitedData = $Data | Select-Object -First $MaxItems
        return $limitedData | ConvertTo-Json -Depth 3 -Compress
    }
    return "No data available"
}

# Function to convert data to HTML table
function Convert-DataToHTMLTable {
    param(
        [object]$Data,
        [string]$TableTitle,
        [int]$MaxRows = 20
    )
    
    if ($Data -and $Data.Count -gt 0) {
        $limitedData = $Data | Select-Object -First $MaxRows
        $html = "<h4>$TableTitle</h4>`n"
        $html += "<div class='table-responsive'>`n"
        $html += $limitedData | ConvertTo-Html -Fragment
        $html += "`n</div>`n"
        return $html
    }
    return "<h4>$TableTitle</h4><p class='text-muted'>No data available</p>`n"
}

# Function to get health status color
function Get-HealthStatusColor {
    param([string]$Status)
    
    switch ($Status) {
        "Healthy" { return "#28a745" }
        "Warning" { return "#ffc107" }
        "Critical" { return "#dc3545" }
        default { return "#6c757d" }
    }
}

# Function to create log-based alert (24-hour monitoring)
function New-LogAlert {
    param(
        [string]$AlertName,
        [string]$Description,
        [string]$Query,
        [string]$ResourceId,
        [string]$ActionGroupId,
        [int]$Severity = 2
    )
    
    try {
        Write-Host "    Creating alert: $AlertName" -ForegroundColor Yellow
        
        $existingAlert = Get-AzScheduledQueryRule -ResourceGroupName $alertResourceGroup -Name $AlertName -ErrorAction SilentlyContinue
        if ($existingAlert) {
            Write-Host "    ⚠ Alert '$AlertName' already exists - skipping" -ForegroundColor Yellow
            return @{ Name = $AlertName; Status = "Exists"; Created = $false; Description = $Description }
        }
        
        # 24-hour monitoring cycle
        $source = New-AzScheduledQueryRuleSource -Query $Query -DataSourceId $ResourceId
        $schedule = New-AzScheduledQueryRuleSchedule -FrequencyInMinutes 1440 -TimeWindowInMinutes 1440
        $triggerCondition = New-AzScheduledQueryRuleTriggerCondition -ThresholdOperator "GreaterThan" -Threshold 0
        $aznsActionGroup = New-AzScheduledQueryRuleAznsActionGroup -ActionGroup $ActionGroupId
        $alertingAction = New-AzScheduledQueryRuleAlertingAction -AznsAction $aznsActionGroup -Severity $Severity -Trigger $triggerCondition
        
        New-AzScheduledQueryRule -ResourceGroupName $alertResourceGroup -Name $AlertName -Description $Description -Location "Australia East" -Source $source -Schedule $schedule -Action $alertingAction
        
        Write-Host "    ✓ Successfully created alert: $AlertName" -ForegroundColor Green
        return @{ Name = $AlertName; Status = "Created"; Created = $true; Query = $Query; Description = $Description }
    }
    catch {
        Write-Error "    ✗ Failed to create alert '$AlertName': $($_.Exception.Message)"
        return @{ Name = $AlertName; Status = "Failed"; Created = $false; Error = $_.Exception.Message; Description = $Description }
    }
}

# Function to analyze 30-day metrics and create intelligent alerts
function New-IntelligentAlerts {
    param(
        [object]$PerformanceData,
        [object]$ExceptionsData,
        [object]$DependencyData,
        [object]$SlowRequestsData,
        [string]$AppInsightsResourceId,
        [string]$ActionGroupId
    )
    
    $createdAlerts = @()
    $skippedAlerts = @()
    $failedAlerts = @()
    
    Write-Host "`n=== INTELLIGENT ALERT CREATION BASED ON 30-DAY ANALYSIS ===" -ForegroundColor Green
    
    # Performance-based alerts (24-hour lookback)
    if ($PerformanceData -and $PerformanceData.Count -gt 0) {
        $avgDuration = ($PerformanceData | Measure-Object -Property duration -Average).Average
        $failureRate = (($PerformanceData | Where-Object { $_.success -eq "false" }).Count / $PerformanceData.Count) * 100
        
        Write-Host "  Analyzing performance patterns: Avg Duration = $([math]::Round($avgDuration, 0))ms, Failure Rate = $([math]::Round($failureRate, 2))%" -ForegroundColor Cyan
        
        if ($avgDuration -gt 3000) {
            $alertQuery = @"
requests
| where timestamp >= ago(1d)
| summarize avg_duration = avg(duration)
| where avg_duration > 5000
"@
            $result = New-LogAlert -AlertName "High Response Time - $applicationName" -Description "Average response time exceeded 5 seconds over 24 hours" -Query $alertQuery -ResourceId $AppInsightsResourceId -ActionGroupId $ActionGroupId -Severity 2
            if ($result.Created) { $createdAlerts += $result } elseif ($result.Status -eq "Exists") { $skippedAlerts += $result } else { $failedAlerts += $result }
        }
        
        if ($failureRate -gt 1) {
            $alertQuery = @"
requests
| where timestamp >= ago(1d)
| summarize total_requests = count(), failed_requests = countif(success == false)
| extend failure_rate = (failed_requests * 100.0) / total_requests
| where failure_rate > 2
"@
            $result = New-LogAlert -AlertName "High Failure Rate - $applicationName" -Description "Request failure rate exceeded 2% over 24 hours" -Query $alertQuery -ResourceId $AppInsightsResourceId -ActionGroupId $ActionGroupId -Severity 1
            if ($result.Created) { $createdAlerts += $result } elseif ($result.Status -eq "Exists") { $skippedAlerts += $result } else { $failedAlerts += $result }
        }
    }
    
    # Exception-based alerts (24-hour lookback)
    if ($ExceptionsData -and $ExceptionsData.Count -gt 0) {
        $criticalExceptions = $ExceptionsData | Where-Object { $_.type -match "OutOfMemoryException|StackOverflowException|SqlException|TimeoutException" }
        Write-Host "  Analyzing exceptions: $($ExceptionsData.Count) total, $($criticalExceptions.Count) critical" -ForegroundColor Cyan
        
        if ($criticalExceptions.Count -gt 0) {
            $alertQuery = @"
exceptions
| where timestamp >= ago(1d)
| where type contains 'OutOfMemoryException' or type contains 'SqlException' or type contains 'TimeoutException' or type contains 'StackOverflowException'
| summarize exception_count = count()
| where exception_count > 0
"@
            $result = New-LogAlert -AlertName "Critical Exceptions - $applicationName" -Description "Critical exceptions detected in the last 24 hours" -Query $alertQuery -ResourceId $AppInsightsResourceId -ActionGroupId $ActionGroupId -Severity 0
            if ($result.Created) { $createdAlerts += $result } elseif ($result.Status -eq "Exists") { $skippedAlerts += $result } else { $failedAlerts += $result }
        }
    }
    
    # Dependency-based alerts (24-hour lookback)
    if ($DependencyData -and $DependencyData.Count -gt 0) {
        $failedDependencies = $DependencyData | Where-Object { $_.success -eq "false" }
        if ($failedDependencies.Count -gt 0) {
            $depFailureRate = ($failedDependencies.Count / $DependencyData.Count) * 100
            Write-Host "  Analyzing dependencies: $($DependencyData.Count) calls, $([math]::Round($depFailureRate, 2))% failure rate" -ForegroundColor Cyan
            
            if ($depFailureRate -gt 5) {
                $alertQuery = @"
dependencies
| where timestamp >= ago(1d)
| summarize total_calls = count(), failed_calls = countif(success == false)
| extend failure_rate = (failed_calls * 100.0) / total_calls
| where failure_rate > 5
"@
                $result = New-LogAlert -AlertName "Dependency Failures - $applicationName" -Description "Dependency failure rate exceeded 5% over 24 hours" -Query $alertQuery -ResourceId $AppInsightsResourceId -ActionGroupId $ActionGroupId -Severity 2
                if ($result.Created) { $createdAlerts += $result } elseif ($result.Status -eq "Exists") { $skippedAlerts += $result } else { $failedAlerts += $result }
            }
        }
    }
    
    # Slow request alerts (24-hour lookback)
    if ($SlowRequestsData -and $SlowRequestsData.Count -gt 20) {
        Write-Host "  Analyzing slow requests: $($SlowRequestsData.Count) slow requests found" -ForegroundColor Cyan
        
        $alertQuery = @"
requests
| where timestamp >= ago(1d)
| where duration > 8000
| summarize slow_request_count = count()
| where slow_request_count > 10
"@
        $result = New-LogAlert -AlertName "Slow Requests - $applicationName" -Description "More than 10 slow requests (>8s) detected in 24 hours" -Query $alertQuery -ResourceId $AppInsightsResourceId -ActionGroupId $ActionGroupId -Severity 2
        if ($result.Created) { $createdAlerts += $result } elseif ($result.Status -eq "Exists") { $skippedAlerts += $result } else { $failedAlerts += $result }
    }
    
    return @{
        Created = $createdAlerts
        Skipped = $skippedAlerts
        Failed = $failedAlerts
        Total = ($createdAlerts.Count + $skippedAlerts.Count + $failedAlerts.Count)
    }
}

# Enhanced HTML report function with full analysis and alerts
function New-HTMLReportWithAlerts {
    param(
        [string]$ApplicationName,
        [hashtable]$WebAppDetails,
        [object]$PerformanceData,
        [object]$ExceptionsData,
        [object]$DependencyData,
        [object]$ResourceData,
        [object]$CustomEventsData,
        [object]$SlowRequestsData,
        [string]$ComprehensiveAnalysis,
        [string]$PerformanceAnalysis,
        [string]$ExceptionAnalysis,
        [string]$DependencyAnalysis,
        [string]$ResourceAnalysis,
        [string]$HealthStatus,
        [hashtable]$AlertResults = $null
    )
    
    $reportDate = Get-Date -Format "MMMM dd, yyyy at HH:mm:ss"
    $healthColor = Get-HealthStatusColor -Status $HealthStatus
    
    # Create performance metrics summary
    $performanceSummary = ""
    if ($PerformanceData -and $PerformanceData.Count -gt 0) {
        $avgDuration = ($PerformanceData | Measure-Object -Property duration -Average).Average
        $failedRequests = ($PerformanceData | Where-Object { $_.success -eq "false" } | Measure-Object).Count
        $totalRequests = ($PerformanceData | Measure-Object).Count
        $avgFailureRate = if ($totalRequests -gt 0) { ($failedRequests / $totalRequests) * 100 } else { 0 }
        
        $performanceSummary = @"
        <div class="row">
            <div class="col-md-3">
                <div class="metric-card">
                    <h5>Total Requests</h5>
                    <span class="metric-value">$($totalRequests.ToString("N0"))</span>
                </div>
            </div>
            <div class="col-md-3">
                <div class="metric-card">
                    <h5>Avg Response Time</h5>
                    <span class="metric-value">$([math]::Round($avgDuration, 0)) ms</span>
                </div>
            </div>
            <div class="col-md-3">
                <div class="metric-card">
                    <h5>Failure Rate</h5>
                    <span class="metric-value">$([math]::Round($avgFailureRate, 2))%</span>
                </div>
            </div>
            <div class="col-md-3">
                <div class="metric-card">
                    <h5>Health Status</h5>
                    <span class="metric-value" style="color: $healthColor">$HealthStatus</span>
                </div>
            </div>
        </div>
"@
    }
    
    # Generate alerts section with detailed alert names
    $alertsSection = ""
    if ($AlertResults -and $AlertResults.Total -gt 0) {
        $alertsSection = @"
        <div id="alerts-created" class="analysis-section alert-new">
            <h2><i class="fas fa-bell section-icon"></i>Intelligent Alerts Created</h2>
            <div class="alert alert-info">
                <i class="fas fa-info-circle"></i> <strong>Alert Summary:</strong> $($AlertResults.Created.Count) created, $($AlertResults.Skipped.Count) already exist, $($AlertResults.Failed.Count) failed
            </div>
            
            <div class="row">
                <div class="col-md-4">
                    <div class="metric-card">
                        <h5>New Alerts</h5>
                        <span class="metric-value" style="color: #28a745;">$($AlertResults.Created.Count)</span>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="metric-card">
                        <h5>Already Exist</h5>
                        <span class="metric-value" style="color: #ffc107;">$($AlertResults.Skipped.Count)</span>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="metric-card">
                        <h5>Failed</h5>
                        <span class="metric-value" style="color: #dc3545;">$($AlertResults.Failed.Count)</span>
                    </div>
                </div>
            </div>
            
            <h4>Alert Details</h4>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Alert Name</th>
                            <th>Status</th>
                            <th>Monitoring Schedule</th>
                            <th>Description</th>
                        </tr>
                    </thead>
                    <tbody>
"@
        
        $allAlerts = @()
        $allAlerts += $AlertResults.Created
        $allAlerts += $AlertResults.Skipped
        $allAlerts += $AlertResults.Failed
        
        foreach ($alert in $allAlerts) {
            $status = "Unknown"
            $statusColor = "#6c757d"
            
            if ($alert.Status -eq "Created") {
                $status = "✓ Created"
                $statusColor = "#28a745"
            } elseif ($alert.Status -eq "Exists") {
                $status = "⚠ Already Exists"
                $statusColor = "#ffc107"
            } elseif ($alert.Status -eq "Failed") {
                $status = "✗ Failed"
                $statusColor = "#dc3545"
            }
            
            $description = if ($alert.Description) { $alert.Description } else { "Intelligent monitoring based on 30-day pattern analysis" }
            
            $alertsSection += @"
                        <tr>
                            <td><strong>$($alert.Name)</strong></td>
                            <td><span style="color: $statusColor;">$status</span></td>
                            <td>Every 24 hours (24-hour lookback)</td>
                            <td>$description</td>
                        </tr>
"@
        }
        
        $alertsSection += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }
    
    # Format analysis content
    $ComprehensiveAnalysisFormatted = if ($ComprehensiveAnalysis -and $ComprehensiveAnalysis -ne "Azure OpenAI analysis unavailable due to API error.") { $ComprehensiveAnalysis -replace "`n", "<br/>" } else { "Comprehensive analysis completed with 30-day data pattern analysis." }
    $PerformanceAnalysisFormatted = if ($PerformanceAnalysis -and $PerformanceAnalysis -ne "Azure OpenAI analysis unavailable due to API error.") { $PerformanceAnalysis -replace "`n", "<br/>" } else { "" }
    $ExceptionAnalysisFormatted = if ($ExceptionAnalysis -and $ExceptionAnalysis -ne "Azure OpenAI analysis unavailable due to API error.") { $ExceptionAnalysis -replace "`n", "<br/>" } else { "" }
    $DependencyAnalysisFormatted = if ($DependencyAnalysis -and $DependencyAnalysis -ne "Azure OpenAI analysis unavailable due to API error.") { $DependencyAnalysis -replace "`n", "<br/>" } else { "" }
    $ResourceAnalysisFormatted = if ($ResourceAnalysis -and $ResourceAnalysis -ne "Azure OpenAI analysis unavailable due to API error.") { $ResourceAnalysis -replace "`n", "<br/>" } else { "" }
    
    # Generate analysis sections
    $performanceAnalysisSection = if ($PerformanceAnalysisFormatted) {
        @"
        <div id="performance-analysis" class="analysis-section priority-high">
            <h2><i class="fas fa-chart-line section-icon"></i>Performance Analysis & Optimization</h2>
            <div class="alert alert-info">
                <i class="fas fa-info-circle"></i> <strong>Data Points Analyzed:</strong> $($PerformanceData.Count) performance metrics
            </div>
            <div class="improvement-card">
                <div class="improvement-content">$PerformanceAnalysisFormatted</div>
            </div>
        </div>
"@
    } else {
        @"
        <div id="performance-analysis" class="analysis-section">
            <h2><i class="fas fa-chart-line section-icon"></i>Performance Analysis</h2>
            <div class="alert alert-warning">
                <i class="fas fa-exclamation-triangle"></i> Performance analysis unavailable - check Azure OpenAI connection
            </div>
        </div>
"@
    }

    $exceptionAnalysisSection = if ($ExceptionAnalysisFormatted) {
        @"
        <div id="exception-analysis" class="analysis-section priority-high">
            <h2><i class="fas fa-exclamation-triangle section-icon"></i>Exception Analysis & Remediation</h2>
            <div class="alert alert-warning">
                <i class="fas fa-exclamation-triangle"></i> <strong>Exception Types Found:</strong> $($ExceptionsData.Count) unique exception patterns
            </div>
            <div class="improvement-card">
                <div class="improvement-content">$ExceptionAnalysisFormatted</div>
            </div>
        </div>
"@
    } else {
        @"
        <div id="exception-analysis" class="analysis-section">
            <h2><i class="fas fa-exclamation-triangle section-icon"></i>Exception Analysis</h2>
            <div class="alert alert-info">
                <i class="fas fa-info-circle"></i> Exception analysis unavailable - no exception data found or API error
            </div>
        </div>
"@
    }

    $dependencyAnalysisSection = if ($DependencyAnalysisFormatted) {
        @"
        <div id="dependency-analysis" class="analysis-section priority-medium">
            <h2><i class="fas fa-project-diagram section-icon"></i>Dependency Analysis & Optimization</h2>
            <div class="alert alert-primary">
                <i class="fas fa-project-diagram"></i> <strong>Dependencies Analyzed:</strong> $($DependencyData.Count) external dependencies
            </div>
            <div class="improvement-card">
                <div class="improvement-content">$DependencyAnalysisFormatted</div>
            </div>
        </div>
"@
    } else {
        @"
        <div id="dependency-analysis" class="analysis-section">
            <h2><i class="fas fa-project-diagram section-icon"></i>Dependency Analysis</h2>
            <div class="alert alert-info">
                <i class="fas fa-info-circle"></i> Dependency analysis unavailable - no dependency data found or API error
            </div>
        </div>
"@
    }
    
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Application Analysis Report - $ApplicationName (30 Days)</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8f9fa; }
        .header-section { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 2rem 0; margin-bottom: 2rem; }
        .metric-card { background: white; padding: 1.5rem; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 1rem; text-align: center; }
        .metric-value { font-size: 2rem; font-weight: bold; color: #495057; }
        .analysis-section { background: white; padding: 2rem; margin-bottom: 2rem; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .priority-high { border-left: 5px solid #dc3545; }
        .priority-medium { border-left: 5px solid #ffc107; }
        .priority-low { border-left: 5px solid #28a745; }
        .alert-new { border-left: 5px solid #17a2b8; }
        .table-responsive { max-height: 400px; overflow-y: auto; }
        .section-icon { margin-right: 0.5rem; }
        .improvement-card { border: 1px solid #dee2e6; border-radius: 8px; padding: 1.5rem; margin-bottom: 1rem; }
        .improvement-content { line-height: 1.6; }
        .toc { background: #f8f9fa; padding: 1.5rem; border-radius: 8px; margin-bottom: 2rem; }
        .footer-section { background: #495057; color: white; padding: 2rem 0; margin-top: 3rem; text-align: center; }
        table { font-size: 0.85em; }
    </style>
</head>
<body>
    <div class="header-section">
        <div class="container">
            <div class="row">
                <div class="col-md-8">
                    <h1><i class="fas fa-chart-line"></i> Application Analysis Report</h1>
                    <h2>$ApplicationName</h2>
                    <p class="lead">30-Day Analysis with Intelligent Alerts - Generated on $reportDate</p>
                </div>
                <div class="col-md-4">
                    <div class="card bg-light text-dark">
                        <div class="card-body">
                            <h5>Application Details</h5>
                            <p><strong>Name:</strong> $($WebAppDetails.Name)</p>
                            <p><strong>Location:</strong> $($WebAppDetails.Location)</p>
                            <p><strong>State:</strong> $($WebAppDetails.State)</p>
                            <p><strong>Analysis Period:</strong> Last 30 days</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div class="container">
        <div class="toc">
            <h3><i class="fas fa-list"></i> Table of Contents</h3>
            <div class="row">
                <div class="col-md-6">
                    <ul class="list-unstyled">
                        <li><a href="#executive-summary"><i class="fas fa-chart-pie section-icon"></i>Executive Summary</a></li>
                        <li><a href="#alerts-created"><i class="fas fa-bell section-icon"></i>Intelligent Alerts</a></li>
                        <li><a href="#performance-metrics"><i class="fas fa-tachometer-alt section-icon"></i>Performance Metrics</a></li>
                        <li><a href="#performance-analysis"><i class="fas fa-chart-line section-icon"></i>Performance Analysis</a></li>
                    </ul>
                </div>
                <div class="col-md-6">
                    <ul class="list-unstyled">
                        <li><a href="#exception-analysis"><i class="fas fa-exclamation-triangle section-icon"></i>Exception Analysis</a></li>
                        <li><a href="#dependency-analysis"><i class="fas fa-project-diagram section-icon"></i>Dependency Analysis</a></li>
                        <li><a href="#resource-analysis"><i class="fas fa-server section-icon"></i>Resource Analysis</a></li>
                        <li><a href="#raw-data"><i class="fas fa-database section-icon"></i>Raw Data Tables</a></li>
                    </ul>
                </div>
            </div>
        </div>

        <div id="performance-metrics" class="analysis-section">
            <h2><i class="fas fa-tachometer-alt section-icon"></i>Performance Metrics Overview</h2>
            $performanceSummary
        </div>

        <div id="executive-summary" class="analysis-section">
            <h2><i class="fas fa-chart-pie section-icon"></i>Executive Summary & Strategic Roadmap</h2>
            <div class="improvement-card">
                <div class="improvement-content">$ComprehensiveAnalysisFormatted</div>
            </div>
        </div>

        $alertsSection

        $performanceAnalysisSection

        $exceptionAnalysisSection

        $dependencyAnalysisSection

        <div id="raw-data" class="analysis-section">
            <h2><i class="fas fa-database section-icon"></i>Raw Data Tables</h2>
            <div class="accordion" id="dataAccordion">
                <div class="accordion-item">
                    <h2 class="accordion-header" id="headingPerformance">
                        <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapsePerformance">
                            <i class="fas fa-chart-line section-icon"></i>Performance Data ($($PerformanceData.Count) records)
                        </button>
                    </h2>
                    <div id="collapsePerformance" class="accordion-collapse collapse" data-bs-parent="#dataAccordion">
                        <div class="accordion-body">
                            $(Convert-DataToHTMLTable -Data $PerformanceData -TableTitle "Performance Metrics")
                        </div>
                    </div>
                </div>
                <div class="accordion-item">
                    <h2 class="accordion-header" id="headingExceptions">
                        <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseExceptions">
                            <i class="fas fa-exclamation-triangle section-icon"></i>Exception Data ($($ExceptionsData.Count) records)
                        </button>
                    </h2>
                    <div id="collapseExceptions" class="accordion-collapse collapse" data-bs-parent="#dataAccordion">
                        <div class="accordion-body">
                            $(Convert-DataToHTMLTable -Data $ExceptionsData -TableTitle "Exception Analysis")
                        </div>
                    </div>
                </div>
                <div class="accordion-item">
                    <h2 class="accordion-header" id="headingDependencies">
                        <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseDependencies">
                            <i class="fas fa-project-diagram section-icon"></i>Dependency Data ($($DependencyData.Count) records)
                        </button>
                    </h2>
                    <div id="collapseDependencies" class="accordion-collapse collapse" data-bs-parent="#dataAccordion">
                        <div class="accordion-body">
                            $(Convert-DataToHTMLTable -Data $DependencyData -TableTitle "Dependency Performance")
                        </div>
                    </div>
                </div>
                <div class="accordion-item">
                    <h2 class="accordion-header" id="headingSlowRequests">
                        <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSlowRequests">
                            <i class="fas fa-clock section-icon"></i>Slow Requests ($($SlowRequestsData.Count) records)
                        </button>
                    </h2>
                    <div id="collapseSlowRequests" class="accordion-collapse collapse" data-bs-parent="#dataAccordion">
                        <div class="accordion-body">
                            $(Convert-DataToHTMLTable -Data $SlowRequestsData -TableTitle "Slow Request Analysis")
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div class="footer-section">
        <div class="container">
            <p>&copy; $(Get-Date -Format "yyyy") Application Analysis Report | Generated by Enhanced Azure Analysis Script with Intelligent Alerts</p>
            <p><small>30-day analysis with 24-hour alert monitoring | Created by Charlie | Application: $ApplicationName</small></p>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        document.querySelectorAll('a[href^="#"]').forEach(anchor => {
            anchor.addEventListener('click', function (e) {
                e.preventDefault();
                document.querySelector(this.getAttribute('href')).scrollIntoView({
                    behavior: 'smooth'
                });
            });
        });
    </script>
</body>
</html>
"@

    return $htmlContent
}

# =============================================================================
# MAIN SCRIPT EXECUTION
# =============================================================================

Write-Host "=== Enhanced Azure Application Analysis Script (30 Days + Intelligent Alerts) ===" -ForegroundColor Green
Write-Host "Starting comprehensive 30-day analysis..." -ForegroundColor Yellow
Write-Host "Target Application: $applicationName" -ForegroundColor Cyan

# Parse the resource ID to extract components
$resourceIdParts = $appInsightsResourceId -split '/'
$appInsightsSubscriptionId = $resourceIdParts[2]
$appInsightsResourceGroup = $resourceIdParts[4]
$appInsightsName = $resourceIdParts[8]

Write-Host "`nConnecting to Azure..." -ForegroundColor Yellow
try {
    Set-AzContext -SubscriptionId $appInsightsSubscriptionId | Out-Null
    Write-Host "✓ Connected to Azure subscription" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to Azure subscription: $($_.Exception.Message)"
    exit 1
}

try {
    $appInsights = Get-AzApplicationInsights -ResourceGroupName $appInsightsResourceGroup -Name $appInsightsName
    $appId = $appInsights.AppId
    Write-Host "✓ Connected to Application Insights: $($appInsights.Name)" -ForegroundColor Green
} catch {
    Write-Error "Failed to retrieve Application Insights resource: $($_.Exception.Message)"
    exit 1
}

try {
    $context = Get-AzContext
    $token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($context.Account, $context.Environment, $context.Tenant.Id, $null, "Never", $null, "https://api.applicationinsights.io/").AccessToken
    $accessToken = @{ Token = $token }
    Write-Host "✓ Access token obtained successfully" -ForegroundColor Green
} catch {
    Write-Error "Failed to get access token: $($_.Exception.Message)"
    exit 1
}

# Handle web application details
if ($webAppSubscriptionId -ne $appInsightsSubscriptionId) {
    Set-AzContext -SubscriptionId $webAppSubscriptionId | Out-Null
    try {
        $webApp = Get-AzWebApp -ResourceGroupName $webAppResourceGroup -Name $applicationName
        $webAppDetails = @{
            Name = $webApp.Name
            Location = $webApp.Location
            State = $webApp.State
            DefaultHostName = $webApp.DefaultHostName
            Kind = $webApp.Kind
        }
    } catch {
        $webAppDetails = @{ Name = $applicationName; Location = "Unknown"; State = "Unknown"; DefaultHostName = "Unknown"; Kind = "Unknown" }
    }
    Set-AzContext -SubscriptionId $appInsightsSubscriptionId | Out-Null
} else {
    try {
        $webApp = Get-AzWebApp -ResourceGroupName $webAppResourceGroup -Name $applicationName
        $webAppDetails = @{
            Name = $webApp.Name
            Location = $webApp.Location
            State = $webApp.State
            DefaultHostName = $webApp.DefaultHostName
            Kind = $webApp.Kind
        }
    } catch {
        $webAppDetails = @{ Name = $applicationName; Location = "Unknown"; State = "Unknown"; DefaultHostName = "Unknown"; Kind = "Unknown" }
    }
}

# =============================================================================
# DATA COLLECTION QUERIES (30-DAY ANALYSIS)
# =============================================================================

Write-Host "`nCollecting 30 days of application telemetry data..." -ForegroundColor Yellow

$performanceQuery = @"
requests
| where timestamp >= ago(30d)
| project timestamp, duration, success, operation_Name, url, resultCode
| order by timestamp desc
| limit 5000
"@

$exceptionsQuery = @"
exceptions
| where timestamp >= ago(30d)
| project timestamp, type, method, assembly, outerMessage, details
| order by timestamp desc
| limit 2000
"@

$dependencyQuery = @"
dependencies
| where timestamp >= ago(30d)
| project timestamp, type, target, name, duration, success, resultCode
| order by timestamp desc
| limit 2000
"@

$slowRequestsQuery = @"
requests
| where timestamp >= ago(30d)
| where duration > 5000
| project timestamp, operation_Name, duration, url, success
| order by duration desc
| limit 500
"@

Write-Host "`nExecuting 30-day telemetry queries..." -ForegroundColor Yellow

$performanceData = Invoke-AppInsightsQuery -Query $performanceQuery -AppId $appId -AccessToken $accessToken
$exceptionsData = Invoke-AppInsightsQuery -Query $exceptionsQuery -AppId $appId -AccessToken $accessToken  
$dependencyData = Invoke-AppInsightsQuery -Query $dependencyQuery -AppId $appId -AccessToken $accessToken
$slowRequestsData = Invoke-AppInsightsQuery -Query $slowRequestsQuery -AppId $appId -AccessToken $accessToken
$customEventsData = @()
$resourceData = @()

Write-Host "`n=== 30-Day Data Collection Summary ===" -ForegroundColor Cyan
Write-Host "Performance data points: $($performanceData.Count)" -ForegroundColor White
Write-Host "Exception types: $($exceptionsData.Count)" -ForegroundColor White
Write-Host "Dependencies: $($dependencyData.Count)" -ForegroundColor White
Write-Host "Slow requests: $($slowRequestsData.Count)" -ForegroundColor White

# =============================================================================
# INTELLIGENT ALERT CREATION
# =============================================================================

$alertResults = $null
try {
    $actionGroup = Get-AzActionGroup -ResourceGroupName $alertResourceGroup -Name $alertActionGroup -ErrorAction SilentlyContinue
    if ($actionGroup) {
        $actionGroupId = $actionGroup.Id
        Write-Host "✓ Using action group: $($actionGroup.Name)" -ForegroundColor Green
        
        $alertResults = New-IntelligentAlerts -PerformanceData $performanceData -ExceptionsData $exceptionsData -DependencyData $dependencyData -SlowRequestsData $slowRequestsData -AppInsightsResourceId $appInsightsResourceId -ActionGroupId $actionGroupId
    } else {
        Write-Host "⚠ Action group '$alertActionGroup' not found. Skipping alert creation." -ForegroundColor Yellow
    }
} catch {
    Write-Warning "Alert creation failed: $($_.Exception.Message)"
}

# =============================================================================
# AZURE OPENAI ANALYSIS - FULL ANALYSIS RESTORED
# =============================================================================

Write-Host "`nAnalyzing data with Azure OpenAI..." -ForegroundColor Yellow

# 1. Performance Analysis
Write-Host "  → Analyzing performance metrics..." -ForegroundColor Gray
$performanceAnalysis = $null
if ($performanceData -and $performanceData.Count -gt 0) {
    $performanceDataJson = Convert-DataForAI -Data $performanceData
    $performancePrompt = @"
Analyze the following web application performance data for '$applicationName' over the last 30 days:

Performance Metrics:
$performanceDataJson

Please provide a detailed analysis with:
1. **Performance Health Score** (1-10 scale with justification)
2. **Critical Performance Issues** (specific problems requiring immediate attention)
3. **Optimization Recommendations** (actionable steps with expected impact)
4. **Priority Ranking** (High/Medium/Low with business impact assessment)
5. **Success Metrics** (KPIs to track improvement)

Focus on practical, actionable insights that can improve user experience and reduce operational costs.
"@

    $performanceAnalysis = Invoke-AzureOpenAI -Prompt $performancePrompt -SystemMessage "You are an expert web application performance analyst with deep knowledge of Azure App Service optimization." -MaxTokens 2500
}

# 2. Exception Analysis
Write-Host "  → Analyzing exception patterns..." -ForegroundColor Gray
$exceptionAnalysis = $null
if ($exceptionsData -and $exceptionsData.Count -gt 0) {
    $exceptionDataJson = Convert-DataForAI -Data $exceptionsData
    $exceptionPrompt = @"
Analyze the following exception data from web application '$applicationName' over the last 30 days:

Exception Data:
$exceptionDataJson

Please provide:
1. **Exception Severity Assessment** (categorize by impact and urgency)
2. **Root Cause Analysis** (probable causes for top exceptions)
3. **Immediate Actions** (quick fixes to reduce exception rate)
4. **Long-term Solutions** (architectural improvements)
5. **Prevention Strategies** (code quality and testing improvements)

Focus on exceptions that significantly impact user experience and application stability.
"@

    $exceptionAnalysis = Invoke-AzureOpenAI -Prompt $exceptionPrompt -SystemMessage "You are an expert software engineer specializing in exception handling and application stability." -MaxTokens 2500
}

# 3. Dependency Analysis
Write-Host "  → Analyzing dependency performance..." -ForegroundColor Gray
$dependencyAnalysis = $null
if ($dependencyData -and $dependencyData.Count -gt 0) {
    $dependencyDataJson = Convert-DataForAI -Data $dependencyData
    $dependencyPrompt = @"
Analyze the following dependency performance data for '$applicationName' over the last 30 days:

Dependency Data:
$dependencyDataJson

Please provide:
1. **Dependency Health Assessment** (identify problematic dependencies)
2. **Performance Bottlenecks** (slowest and most unreliable dependencies)
3. **Optimization Strategies** (specific improvements for each dependency type)
4. **Caching Recommendations** (appropriate caching strategies)
5. **Architecture Improvements** (connection pooling, async patterns)

Focus on dependencies that are performance bottlenecks or reliability risks.
"@

    $dependencyAnalysis = Invoke-AzureOpenAI -Prompt $dependencyPrompt -SystemMessage "You are an expert in distributed systems and dependency management." -MaxTokens 2500
}

# 4. Comprehensive Strategic Analysis
Write-Host "  → Generating comprehensive recommendations..." -ForegroundColor Gray
$comprehensivePrompt = @"
Based on the comprehensive 30-day telemetry analysis for web application '$applicationName', provide an executive-level optimization roadmap:

## Key Metrics Summary
- **Performance Data Points**: $($performanceData.Count)
- **Exception Types**: $($exceptionsData.Count)
- **Dependencies**: $($dependencyData.Count)
- **Slow Requests**: $($slowRequestsData.Count)

Please provide:
1. **Executive Summary** (application health overview and key findings)
2. **Critical Issues** (top 3 issues requiring immediate attention)
3. **30-Day Action Plan** (prioritized roadmap with timelines)
4. **Success Metrics** (KPIs to track improvement progress)
5. **ROI Projections** (expected benefits and cost savings)

Format as a structured report suitable for technical leadership.
"@

$comprehensiveAnalysis = Invoke-AzureOpenAI -Prompt $comprehensivePrompt -SystemMessage "You are a senior application architect providing strategic recommendations to technical leadership." -MaxTokens 3500

# =============================================================================
# GENERATE ENHANCED HTML REPORT
# =============================================================================

$healthStatus = "Healthy"
if ($performanceData.Count -gt 0) {
    $failedRequests = ($performanceData | Where-Object { $_.success -eq "false" } | Measure-Object).Count
    $totalRequests = $performanceData.Count
    $failureRate = if ($totalRequests -gt 0) { ($failedRequests / $totalRequests) * 100 } else { 0 }
    
    if ($failureRate -lt 1) { $healthStatus = "Healthy" }
    elseif ($failureRate -lt 5) { $healthStatus = "Warning" }
    else { $healthStatus = "Critical" }
}

$htmlReport = New-HTMLReportWithAlerts -ApplicationName $applicationName -WebAppDetails $webAppDetails -PerformanceData $performanceData -ExceptionsData $exceptionsData -DependencyData $dependencyData -ResourceData $resourceData -CustomEventsData $customEventsData -SlowRequestsData $slowRequestsData -ComprehensiveAnalysis $comprehensiveAnalysis -PerformanceAnalysis $performanceAnalysis -ExceptionAnalysis $exceptionAnalysis -DependencyAnalysis $dependencyAnalysis -ResourceAnalysis "" -HealthStatus $healthStatus -AlertResults $alertResults

$reportDirectory = Split-Path $reportOutputPath -Parent
if (!(Test-Path $reportDirectory)) {
    New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
}

try {
    $htmlReport | Out-File -FilePath $reportOutputPath -Encoding UTF8
    Write-Host "✓ Enhanced HTML report generated successfully!" -ForegroundColor Green
    Write-Host "Report location: $reportOutputPath" -ForegroundColor Cyan
    
    if ($alertResults) {
        Write-Host "`n=== INTELLIGENT ALERT SUMMARY ===" -ForegroundColor Green
        Write-Host " Alerts created: $($alertResults.Created.Count)" -ForegroundColor Green
        Write-Host " Alerts skipped (already exist): $($alertResults.Skipped.Count)" -ForegroundColor Yellow
        Write-Host " Alerts failed: $($alertResults.Failed.Count)" -ForegroundColor Red
        
        if ($alertResults.Created.Count -gt 0) {
            Write-Host "`nCreated Alert Names:" -ForegroundColor Cyan
            foreach ($alert in $alertResults.Created) {
                Write-Host "  • $($alert.Name)" -ForegroundColor White
            }
        }
        
        if ($alertResults.Skipped.Count -gt 0) {
            Write-Host "`nExisting Alert Names:" -ForegroundColor Yellow
            foreach ($alert in $alertResults.Skipped) {
                Write-Host "  • $($alert.Name)" -ForegroundColor White
            }
        }
    }
    
    Start-Process $reportOutputPath
}
catch {
    Write-Error "Failed to save HTML report: $($_.Exception.Message)"
}

Write-Host "`n=== ANALYSIS COMPLETE ===" -ForegroundColor Green
Write-Host "Application: $applicationName" -ForegroundColor White
Write-Host "Health Status: $healthStatus" -ForegroundColor $(switch ($healthStatus) { "Healthy" { "Green" } "Warning" { "Yellow" } "Critical" { "Red" } default { "Gray" } })
Write-Host "30-day analysis with intelligent 24-hour alerts completed!" -ForegroundColor Green
