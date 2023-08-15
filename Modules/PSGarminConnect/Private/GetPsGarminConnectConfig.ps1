function GetPsGarminConnectConfig {
    PARAM(
        [switch]$Name
    )

    if($Name){
        return "PsGarminConnect_add5f32d-05e9-4c68-b062-2c583064b43d"
    }
    else{
        return (Get-Variable -Name (GetPsGarminConnectConfig -Name) -ValueOnly -Scope Global)
    }
}
