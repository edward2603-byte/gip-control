[CmdletBinding()]
param([string]$OutputPath)
$a=@((Join-Path $PSScriptRoot 'create_stage0.py'))
if($OutputPath){$a+=@('--output',$OutputPath)}
& python @a
if($LASTEXITCODE){throw 'Stage 0 prototype creation failed.'}
