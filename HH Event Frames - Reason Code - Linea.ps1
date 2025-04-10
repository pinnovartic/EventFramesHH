#region Execution Time
$TimeQuerySt = ConvertFrom-AFRelativeTime -RelativeTime "1-mar-25"
#$TimeQuerySt = $TimeQuerySt.ToLocalTime() 
$TimeQueryEt = $TimeQuerySt.AddDays(1)
#$TimeQueryEt = $TimeQueryEt.ToLocalTime()
$HH = ($TimeQueryEt - $TimeQuerySt).TotalHours
#endregion
$Linea = "Linea RE010"

#region Config read
$Str_AFServer = "CNECLPI03"
$Str_AFUser = "CNECLPI03\Administrador"
$Str_AFPassword = "Sgtm2020"
$Str_AFDatabase = "Ecometales"
#endregion

#region PI AF Connection
$secure_pass = ConvertTo-SecureString -String $Str_AFPassword -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential ($Str_AFUser, $secure_pass)
    try{
        $AFServer = Get-AFServer $Str_AFServer
        $AF_Connection = Connect-AFServer -WindowsCredential $credentials -AFServer $AFServer
        $AFDB = Get-AFDatabase -Name $Str_AFDatabase -AFServer $AFServer
        $TempElement = Get-AFElement -AFDatabase $AFDB -Name "Estado Equipos"
        $LineaElement = Get-AFElement -AFElement $TempElement -Name $Linea
        
    }catch { 
        $e = $_.Exception
        $msg = $e.Message
        while ($e.InnerException) {
            $e = $e.InnerException
            $msg += "`n" + $e.Message
            }
        $msg
    }
#endregion

#region PI AF Event Frames

$CurrentDate = Get-Date

While ($TimeQueryEt.ToLocalTime() -lt $CurrentDate){
    
    $HHDowntime_Disp = 0
    $HHDowntime_Util = 0    

    if ($TimeQueryEt.ToLocalTime().Hour -eq 23){
        $TimeQueryEt = $TimeQueryEt.AddHours(1)
    }
    if ($TimeQuerySt.ToLocalTime().Hour -eq 1 -and $TimeQueryEt.ToLocalTime().Hour -eq 1){
        $TimeQueryEt = $TimeQueryEt.AddHours(-1)
    }

    $MyEventFramesOverlapped = Find-AFEventFrame -StartTime ($TimeQuerySt) -EndTime ($TimeQueryEt) -AFSearchMode Overlapped -MaxCount 1000 -ReferencedElementNameFilter $Linea -AFDatabase $AFDB
    
    foreach($Event in $MyEventFramesOverlapped){
        $EventTemplate = $Event.Template        
        $EventST = $Event.StartTime
        $EventET = $Event.EndTime
        
        #Evento cruza limites de ventana de revisión
        If ($EventST.LocalTime -lt $TimeQuerySt.ToLocalTime()){
            $EventST = $TimeQuerySt
        }
        If ($EventET.LocalTime -gt $TimeQueryEt.ToLocalTime()){
            $EventET = $TimeQueryEt
        }
        $EventDuration = ($EventET - $EventST).TotalHours

        If ($EventTemplate.Name -eq "Downtime Disponibilidad"){
            $HHDowntime_Disp = $HHDowntime_Disp + $EventDuration

        }
        If ($EventTemplate.Name -eq "Downtime Utilización"){
            $HHDowntime_Util = $HHDowntime_Util + $EventDuration
        }             
    }

    $HHDia = ($TimeQueryEt - $TimeQuerySt).TotalHours
    #Write-Host $TimeQuerySt.ToLocalTime()
    #Write-Host $TimeQueryEt.ToLocalTime()
    #Write-Host $HHDia
    $HHDisp = $HHDia - $HHDowntime_Disp
    $HHUtil = $HHDia - $HHDowntime_Util

    $ElementAttrTiempoDisponibilidad = Get-AFAttribute -AFElement $LineaElement -Name "HH Disponibilidad"            
    Set-Variable -Name AFValue_HHDisp -Value (New-Object 'OSIsoft.AF.Asset.AFValue')
    $AFValue_HHDisp.Timestamp = $TimeQuerySt
    $AFValue_HHDisp.Value = $HHDisp        
    $ElementAttrTiempoDisponibilidad.SetValue($AFValue_HHDisp)

    $ElementAttrTiempoUtilizacion = Get-AFAttribute -AFElement $LineaElement -Name "HH Utilizacion"            
    Set-Variable -Name AFValue_HHUtil -Value (New-Object 'OSIsoft.AF.Asset.AFValue')
    $AFValue_HHUtil.Timestamp = $TimeQuerySt
    $AFValue_HHUtil.Value = $HHUtil        
    $ElementAttrTiempoUtilizacion.SetValue($AFValue_HHUtil)

    $TimeQuerySt = $TimeQueryEt
    $TimeQueryEt = $TimeQuerySt.AddDays(1)
}