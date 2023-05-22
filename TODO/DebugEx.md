# DebugEx
## PowerShell script
```PowerShell
function Read-VBADebugEx {
	get-content -path 'C:\Users\User\desktop\debugex.log' -Tail 1 -Wait | % { 
		$RemoveQuotations = $_.TrimStart('"').TrimEnd('"')
		$Color = $RemoveQuotations.Split(',')[0] 
		$Timestamp = $RemoveQuotations.Split(',',2)[1].Substring(0,11)
		$Message = $RemoveQuotations.Split(',',2)[1].Substring(11)
		Write-Host $Timestamp -Foreground White -NoNewLine
		Write-Host $Message -Foreground $Color
	}   
}
Set-Alias LogVBA Read-VBADebugEx
```

## Methods
```vb
Public Sub Many(ByVal ArrayToLog As Variant, _
                Optional ByVal Topic As Variant, _
                Optional ByVal LogLevel As LogLevel = -1)
Public Sub Variable(ByVal VariableToLog As Variant, _
                Optional ByVal Topic As Variant, _
                Optional ByVal LogLevel As LogLevel = -1)
Public Sub Message(ByVal Message As String, _
                Optional ByVal Topic As Variant, _
                Optional ByVal LogLevel As LogLevel = -1)
Public Sub LogStop(ByVal Message As String, _
                Optional ByVal Topic As Variant, _
                Optional ByVal LogLevel As LogLevel = -1)
Public Sub LogHR()
Public Sub LogClear()
Public Sub SetDefaultLevel(ByVal LogLevel As LogLevel)
Public Sub SetFilterLevel(ByVal LogLevel As LogLevel)
Public Sub StartLogging()
Public Sub StopLogging()
```

## TODO
- [ ] Separate prefix (bullet and index) from variant in printing arrays
- [ ] Move Array rank checking code to Array helpers
- [ ] Change log output so that PowerShell can split suffix (` (VarType = n, TypeName = x)`) and print it in `DarkGrey`.
- [ ] Do we use it as Predeclared class or as instantiated?
- [ ] Do we add a special case for printing Range() objects? (to include Address, and maybe Areas.Count?)
  - If so, what else will we end up creating special cases for?