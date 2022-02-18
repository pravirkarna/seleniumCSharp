$applicationName = $args[0]
$subModule = $args[1]
$testCaseToBeExecuted = $args[2]
$executionCatog = $args[3]
$envname = $args[4]
$parallezam = $args[5]
write-host  Arguments: $applicationName $subModule $testCaseToBeExecuted $executionCatog $envname $parallezam 
cd C:\Scripts\release
cmd /c .\Maxim.Automation.Ui.TestSuiteRunner.exe $applicationName $subModule $testCaseToBeExecuted $executionCatog $envname $parallezam 
write-host $LastExitCode
exit $LastExitCode