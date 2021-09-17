function Invoke-TestHarness
{
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [System.String]
        $TestResultsFile,

        [Parameter()]
        [System.String]
        $DscTestsPath,

        [Parameter()]
        [Switch]
        $IgnoreCodeCoverage
    )

    $MaximumFunctionCount = 9999
    Write-Verbose -Message 'Starting all Microsoft365DSC tests'

    $repoDir = Join-Path -Path $PSScriptRoot -ChildPath '..\' -Resolve

    <#
    $testCoverageFiles = @()
    if ($IgnoreCodeCoverage.IsPresent -eq $false)
    {
        Get-ChildItem -Path "$repoDir\modules\Microsoft365DSC\DSCResources\**\*.psm1" -Recurse | ForEach-Object {
            if ($_.FullName -notlike '*\DSCResource.Tests\*')
            {
                $testCoverageFiles += $_.FullName
            }
        }
    }
    #>

    Import-Module -Name "$repoDir\modules\Microsoft365DSC\Microsoft365DSC.psd1"
    #$testsToRun = @()

    # $versionsPath = Join-Path -Path $repoDir -ChildPath "\Tests\Unit\Stubs\"

    # Run Unit Tests
    # Import the first stub found so that there is a base module loaded before the tests start
    $firstStub = Join-Path -Path $repoDir `
        -ChildPath "\Tests\Unit\Stubs\Microsoft365.psm1"
    Import-Module $firstStub -WarningAction SilentlyContinue

    $stubPath = Join-Path -Path $repoDir `
            -ChildPath "\Tests\Unit\Stubs\Microsoft365.psm1"
    <#$testsToRun += @(@{
            'Path'       = (Join-Path -Path $repoDir -ChildPath "\Tests\Unit")
            'Parameters' = @{
                'CmdletModule' = $stubPath
            }
        })#>

    # DSC Common Tests
    $getChildItemParameters = @{
        Path    = (Join-Path -Path $repoDir -ChildPath "\Tests\Unit")
        Recurse = $true
        Filter  = '*.Tests.ps1'
    }

    # Get all tests '*.Tests.ps1'.
    $commonTestFiles = Get-ChildItem @getChildItemParameters

    # Remove DscResource.Tests unit tests.
    $commonTestFiles = $commonTestFiles | Where-Object -FilterScript {
        $_.FullName -notmatch 'DSCResource.Tests\\Tests'
    }

    #$testsToRun += @( $commonTestFiles.FullName )

    $filesToExecute = @($commonTestFiles.FullName)
    <#foreach ($testToRun in $testsToRun)
    {
        $filesToExecute += $testToRun
    }#>

    # Build Pester configuration
    $PesterConfig = New-PesterConfiguration
    $PesterConfig.Run.PassThru = $True
    $PesterConfig.Run.Path = $filesToExecute

    if ([String]::IsNullOrEmpty($TestResultsFile) -eq $false) {
        # Enable NUnit output
        $PesterConfig.TestResult.Enabled=$True
        $PesterConfig.TestResult.OutputFormat='JUnitXml'
        $PesterConfig.TestResult.OutputPath=$TestResultsFile
    }

    if ($IgnoreCodeCoverage.IsPresent -eq $false) {
        # Files/Folders to use for Code Coverage
        $CodeCoveragePaths=(Get-childitem "$repoDir\modules\Microsoft365DSC" -Exclude 'Examples','Dependencies','Modules'|Select-Object -ExpandProperty FullName)

        # Enable CodeCoverage output
        $PesterConfig.CodeCoverage.Enabled=$True
        $PesterConfig.CodeCoverage.Path=$CodeCoveragePaths
        # Code Coverage without breakpoints, should be faster (a lot)
        # https://github.com/pester/Pester/releases/tag/5.3.0
        $PesterConfig.CodeCoverage.UseBreakpoints=$False
        $PesterConfig.CodeCoverage.OutputPath='CodeCov.xml'
    }

    $Results=Invoke-Pester -Configuration $PesterConfig

    return $results
}
