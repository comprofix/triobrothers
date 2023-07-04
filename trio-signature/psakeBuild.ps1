properties {
    $script = "$PSScriptRoot\Signatures"
}

task default -depends Analyze, Test

task Analyze {
    $saResults = Invoke-ScriptAnalyzer -Path $script -Severity @('Error') -Recurse -Verbose:$false
    if ($saResults) {
        $saResults | Format-Table
        Write-Error -Message 'One or more Script Analyzer errors/warnings where found. Build cannot continue!'
    }
}

task Test {
    $testResults = Invoke-Pester -Path $PSScriptRoot -PassThru
    if ($testResults.FailedCount -gt 0) {
        $testResults | Format-List
        Write-Error -Message 'One or more Pester tests failed. Build cannot continue!'
    }
}

task Deploy -depends Analyze, Test {
    Remove-Item -Recurse -Force "\\filesrv.trio.local\e$\Signatures\"
    Invoke-PSDeploy -Path "$PSScriptRoot\signature.PSDeploy.ps1" -Force -Verbose:$VerbosePreference
}
