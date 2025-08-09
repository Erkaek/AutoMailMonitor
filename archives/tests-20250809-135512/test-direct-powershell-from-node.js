// Test d'exécution directe du script PowerShell qui fonctionne
const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

async function testDirectPowerShell() {
    try {
        console.log("🧪 Test PowerShell direct depuis Node.js");
        
        // Script PowerShell simplifié sans caractères spéciaux
        const script = `
try {
  $outlook = New-Object -ComObject Outlook.Application
  $namespace = $outlook.GetNamespace("MAPI")
  
  # Fonction pour naviguer vers un dossier specifique
  function Find-OutlookFolder {
    param([string]$FolderPath, [object]$Namespace)
    
    # Extraire compte et chemin
    if ($FolderPath -match '^([^\\\\]+)\\\\(.+)$') {
      $accountName = $matches[1]
      $folderPath = $matches[2]
      
      # Chercher le store/compte
      $targetStore = $null
      foreach ($store in $Namespace.Stores) {
        if ($store.DisplayName -like "*$accountName*" -or $store.DisplayName -eq $accountName) {
          $targetStore = $store
          break
        }
      }
      
      if (-not $targetStore) {
        $targetStore = $Namespace.DefaultStore
      }
      
      # Naviguer dans l'arborescence
      $currentFolder = $targetStore.GetRootFolder()
      $pathParts = $folderPath -split '\\\\'
      
      foreach ($part in $pathParts) {
        if ($part -and $part.Trim() -ne "") {
          $found = $false
          $folders = $currentFolder.Folders
          
          # Recherche exacte par nom
          for ($i = 1; $i -le $folders.Count; $i++) {
            $subfolder = $folders.Item($i)
            if ($subfolder.Name -eq $part) {
              $currentFolder = $subfolder
              $found = $true
              break
            }
          }
          
          # Si pas trouve, recherche par pattern pour "Boîte de réception"
          if (-not $found -and ($part -match "réception" -or $part -match "Boîte")) {
            for ($i = 1; $i -le $folders.Count; $i++) {
              $subfolder = $folders.Item($i)
              if ($subfolder.Name -match "réception") {
                $currentFolder = $subfolder
                $found = $true
                break
              }
            }
          }
          
          if (-not $found) {
            return $null
          }
        }
      }
      
      return $currentFolder
    } else {
      return $null
    }
  }
  
  # Test pour "test"
  $testPath = "erkaekanon@outlook.com\\Boîte de réception\\test"
  $targetFolder = Find-OutlookFolder -FolderPath $testPath -Namespace $namespace
  
  if ($targetFolder) {
    $items = $targetFolder.Items
    $count = $items.Count
    Write-Output "SUCCES: Dossier test trouve avec $count elements"
    
    if ($items.Count -gt 0) {
      $items.Sort("[ReceivedTime]", $true)
      $limit = [Math]::Min(3, $items.Count)
      for ($i = 1; $i -le $limit; $i++) {
        $mail = $items.Item($i)
        if ($mail.Class -eq 43) {
          $subject = $mail.Subject
          Write-Output "Email: $subject"
        }
      }
    }
  } else {
    Write-Output "ECHEC: Dossier test non trouve"
  }
  
  # Test pour "testA"
  $testAPath = "erkaekanon@outlook.com\\Boîte de réception\\testA"
  $targetFolderA = Find-OutlookFolder -FolderPath $testAPath -Namespace $namespace
  
  if ($targetFolderA) {
    $itemsA = $targetFolderA.Items
    $countA = $itemsA.Count
    Write-Output "SUCCES: Dossier testA trouve avec $countA elements"
  } else {
    Write-Output "ECHEC: Dossier testA non trouve"
  }
  
} catch {
  $errorMsg = $_.Exception.Message
  Write-Output "ERREUR: $errorMsg"
}
`;
        
        // Créer fichier temporaire avec encodage UTF-8 BOM
        const tempDir = os.tmpdir();
        const tempFile = path.join(tempDir, `test_direct_${Date.now()}.ps1`);
        
        // UTF-8 BOM
        const BOM = '\uFEFF';
        fs.writeFileSync(tempFile, BOM + script, { encoding: 'utf8' });
        console.log(`📄 Script temporaire: ${tempFile}`);
        
        // Exécuter
        const result = await new Promise((resolve, reject) => {
            exec(`powershell -ExecutionPolicy Bypass -File "${tempFile}"`, (error, stdout, stderr) => {
                // Nettoyer
                try { fs.unlinkSync(tempFile); } catch(e) {}
                
                if (error) {
                    reject(error);
                } else {
                    resolve({ stdout, stderr });
                }
            });
        });
        
        console.log("📋 Résultat:");
        console.log(result.stdout);
        if (result.stderr) {
            console.log("⚠️ Stderr:");
            console.log(result.stderr);
        }
        
    } catch (error) {
        console.error("❌ Erreur:", error.message);
    }
}

testDirectPowerShell();
