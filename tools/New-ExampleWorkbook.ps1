[CmdletBinding()]
param([string]$OutputPath)
$a=@((Join-Path $PSScriptRoot 'create_example_workbook.py'))
if($OutputPath){$a+=@('--output',$OutputPath)}
& python @a
if($LASTEXITCODE){throw 'Example workbook creation failed.'}
