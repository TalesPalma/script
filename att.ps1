# Mapeia os nomes dos clientes
$clientes = @{
    1 = "AMIGO DE PATAS";          2 = "ANCLIVEPA";                3 = "ANIMAL_CARE";
    4 = "ANIMAL_CLINIC_PR";        5 = "CAO QUE MIA";              6 = "CCVET";
    7 = "CHADDAD";                 8 = "CLINICA_FORTE_RS";         9 = "CLINICA_NAVE";
    10 = "CLIVET_QUINTAO";        11 = "CV_BETANIA";              12 = "CV_ROMULO EDGARD";
    13 = "CV_SANTA_CLARA";        14 = "CV_SANTA_QUITERIA";       15 = "CV_VALEDOSOL";
    16 = "CV_VET_CENTER";         17 = "CVREFERENCIA_VET";        18 = "DOGCARE_SP";
    19 = "DOGCENTERPR";           20 = "DOGDOGCATCAT";            21 = "DOM_VETERINARIA";
    22 = "ESPACOPET_AGRO";        23 = "ESPACOPET_BETIM";         24 = "GARRA_HV_VET";
    25 = "GARRA_RS";              26 = "GUTIERREZ";               27 = "HAUKEMIA";
    28 = "HOMOLOGACAO";           29 = "HOSPITAL_VET_SC";         30 = "HOSPITALBOLSON";
    31 = "HV_BATEL";              32 = "HV_NOVA_LIMA";            33 = "HV_POPULAR_GV";
    34 = "HV-SAO BERNARDO";       35 = "IEMEV_RJ";                36 = "KRIAR_VET";
    37 = "LAGOASANTA";            38 = "LU_NEUBARTH";             39 = "MAESTRO";
    40 = "MATERDOG_MG";           41 = "MEDICO_BICHOS_PR";        42 = "MEDVET";
    43 = "MUNDO_ANIMAL";          44 = "OUROVET";                 45 = "PET_LOVE";
    46 = "PETC_PAMPULHA";         47 = "PETCARE-VET004";          48 = "PETNAUTAS";
    49 = "POA_PET_CARE";          50 = "PONTO_CAO_BANGU";         51 = "PRONTODOG";
    52 = "PRONTOVET";             53 = "PUC_PR_TOLEDO";           54 = "QUATROPATAS";
    55 = "RENALPET";              56 = "SALVARE";                 57 = "SOS_ANIMAL";
    58 = "SOS_HV_SC";             59 = "SPADOSPETS";              60 = "UFMG-MG";
    61 = "UFPR";                  62 = "UFR_RJ";                  63 = "UNIFOR_CE";
    64 = "UNIPET";                65 = "VET_CLINIC";              66 = "VET_POINT";
    67 = "VETCENTER_SC";          68 = "VETCLIN";                 69 = "VETMASTER";
    70 = "VIAPIANA";              71 = "VISIOVET";                77 = "VITA";
    78 = "XAOLIN_HV"
}

Write-Host "`nSelecione o cliente que deseja atualizar:`n"
$clientes.GetEnumerator() | Sort-Object Key | ForEach-Object {
    Write-Host ("{0,2} - {1}" -f $_.Key, $_.Value)
}

$NumberSelect = Read-Host "`nDigite o número do cliente"
if (-not $clientes.ContainsKey([int]$NumberSelect)) {
    Write-Host "Cliente inválido. Encerrando."
    exit
}
$clienteNome = $clientes[[int]$NumberSelect]

Write-Host "`n1 - Atualizar o DoctorVet."
Write-Host "2 - Atualizar o Finantec."
Write-Host "3 - Atualizar ambos."

$opcao = Read-Host "`nEscolha a opção"
$basePath = "D:\SiematecSolucoes\$clienteNome"

function Atualizar-Exe ($TARGET) {
    if (-not (Test-Path $TARGET)) {
        Write-Host "Arquivo EXE não encontrado: $TARGET"
        return
    }

    $FOLDER = Split-Path $TARGET -Parent
    $BASENAME = [System.IO.Path]::GetFileNameWithoutExtension($TARGET)
    $EXT = [System.IO.Path]::GetExtension($TARGET)

    $OLD_FILE = Join-Path $FOLDER "$BASENAME.old$EXT"
    if (Test-Path $OLD_FILE) {
        Write-Host "Removendo arquivo antigo: $OLD_FILE"
        Remove-Item -Force -ErrorAction SilentlyContinue $OLD_FILE
    }

    $SHORTCUT = Join-Path $FOLDER "$BASENAME - Shortcut.lnk"
    $WshShell = New-Object -ComObject WScript.Shell
    $ShortcutObj = $WshShell.CreateShortcut($SHORTCUT)
    $ShortcutObj.TargetPath = $TARGET
    $ShortcutObj.Arguments = "/u"
    $ShortcutObj.WorkingDirectory = $FOLDER
    $ShortcutObj.Save()

    Rename-Item -Path $TARGET -NewName "$BASENAME.old$EXT"
    Write-Host "Arquivo renomeado para: $BASENAME.old$EXT"

    Start-Process $SHORTCUT -Verb RunAs
    Pause

    if (Test-Path $TARGET) {
        Write-Host "Novo arquivo atualizado encontrado."
        Start-Process $TARGET -Verb RunAs
    } else {
        Write-Host "Novo arquivo não foi encontrado. A atualização pode ter falhado."
    }
}

switch ($opcao) {
    "1" {
        $TARGET = "$basePath\DoctorVet\DoctorVet.exe"
        Write-Host "`nAtualizando DoctorVet:"
        Atualizar-Exe $TARGET
    }
    "2" {
        $TARGET = "$basePath\Finantec\Finantec.exe"
        Write-Host "`nAtualizando Finantec:"
        Atualizar-Exe $TARGET
    }
    "3" {
        $TARGET_DoctorVet = "$basePath\DoctorVet\DoctorVet.exe"
        $TARGET_Finantec = "$basePath\Finantec\Finantec.exe"

        Write-Host "`nAtualizando DoctorVet:"
        Atualizar-Exe $TARGET_DoctorVet

        Write-Host "`nAtualizando Finantec:"
        Atualizar-Exe $TARGET_Finantec
    }
    default {
        Write-Host "Opção inválida. Encerrando."
        exit
    }
}
