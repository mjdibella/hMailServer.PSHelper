function Connect-hMailServer {
    Param(
        [Parameter(Mandatory=$true)][string]$username,
        [Parameter(Mandatory=$true)][string]$password
    )
    $hMailServer.application = New-Object -ComObject hMailServer.Application
    $hMailServer.identity = $hMailServer.application.Authenticate($username, $password)
    if ($hMailServer.identity -eq $null) {
        throw "Authentication to hMailServer failed.   Check passed credentials and user's Administration level."
    } else {
        New-Item -Path $hMailServer.registryURL -Force | Out-null
        New-ItemProperty -Path $hMailServer.registryURL -Name username -Value $username -Force | Out-Null
        New-ItemProperty -Path $hMailServer.registryURL -Name password -Value $password -Force | Out-Null
        write-host "Connected to hMailServer`n"
    }
}

function Disconnect-hMailServer {
    Remove-ItemProperty -Path $hMailServer.registryURL -Name username | Out-Null
    Remove-ItemProperty -Path $hMailServer.registryURL -Name password | Out-Null
    Remove-Variable -Name hMailServer
}

function Get-hMailServer {
    $hMailServer
}

function Receive-hMailRawMessage {
    [cmdletbinding()]
    param(
        [Parameter(ValueFromPipeline)][string[]]$messageStrings,    
        [Parameter(Mandatory=$true)][string[]]$envelopeRecipients,
        [Parameter(Mandatory=$false)][string[]]$envelopeSender
    )
    begin {
    }
    process {
        if ($_) {
            foreach ($messageString in $messageStrings) {
                $messageObject = New-Object -ComObject hMailServer.Message
                $messageFile = $messageObject.FileName
                $messageString | Out-File -filePath $messageFile -encoding ASCII
                $messageObject.RefreshContent()
                if (-not $envelopeSender) {
                    $originalSender = $messageObject.HeaderValue("Return-Path")
                    if ($originalSender -eq "") {
                        $originalSender = $messageObject.HeaderValue("From")
                    }
                    $messageObject.FromAddress = $originalSender
                } else {
                    $messageObject.FromAddress = $envelopeSender
                }
		        $originalTo = $messageObject.HeaderValue("To")
		        $originalCC = $messageObject.HeaderValue("CC")
		        $messageObject.ClearRecipients()
        		foreach ($envelopeRecipient in $envelopeRecipients) {
    				$messageObject.AddRecipient("", $envelopeRecipient)
                }
    			$messageObject.HeaderValue("To") = $originalTo
                $messageObject.HeaderValue("CC") = $originalCC
    			$messageObject.Save()
            }
        }
    }
    end {
    }
}

function Write-hMailRawMessageToMailbox {
    [cmdletbinding()]
    param(
        [Parameter(ValueFromPipeline)][string[]]$messageStrings,    
        [Parameter(Mandatory=$true)][string[]]$mailbox,
        [Parameter(Mandatory=$false)][string[]]$folder = "INBOX"
    )
    begin {
        $domain = $($mailbox.split("@"))[1]
        $domainObject = $hMailServer.application.domains.itemByName($domain)
        $mailboxObject = $domainObject.accounts.itemByAddress($mailbox)
        $folderObject = $mailboxObject.IMAPFolders.itemByName($folder)
        $count = 0
    }
    process {
        if ($_) {
            foreach ($messageString in $messageStrings) {
                $messageObject = $folderObject.Messages.Add()
                $messageObject.Save()
                $messageFile = $messageObject.FileName
                $messageString | Out-File -filePath $messageFile -encoding ASCII
        		$messageObject.RefreshContent()
    			$messageObject.Save()
                $messageObject.Copy($folderObject.ID)
                $count++
                $resultObject = New-Object PSObject
                $resultObject | Add-Member Noteproperty Count $count
                $resultObject | Add-Member Noteproperty Message-Id $messageObject.Headers.ItemByName("Message-ID").Value
                $resultObject | Add-Member Noteproperty Date $messageObject.Headers.ItemByName("Date").Value
                $resultObject | Add-Member Noteproperty From $messageObject.Headers.ItemByName("From").Value
                $resultObject | Add-Member Noteproperty Subject $messageObject.Headers.ItemByName("Subject").Value
                $resultObject
            	$folderObject.Messages.DeleteByDBID($messageObject.ID)            
            }
        }
    }
    end {
    }
}

$hMailServer = [ordered]@{
    registryURL = "HKCU:\Software\hMailServer\hMailServer.PSHelper"
    application = $null
    identity = $null
}
New-Variable -Name hMailServer -Value $hMailServer -Scope script -Force
$registryKey = (Get-ItemProperty -Path $hMailServer.registryURL -ErrorAction SilentlyContinue)
if ($registryKey -eq $null) {
    Write-Warning "Autoconnect failed.  API key not found in registry.  Use Connect-hMailServer to connect manually."
} else {

    $hMailServer.application = New-Object -ComObject hMailServer.Application
    $hMailServer.identity = $hMailServer.application.Authenticate($registryKey.username, $registryKey.password)
    if ($hMailServer.identity -eq $null) {
        throw "Authentication to hMailServer failed.   Check passed credentials and user's Administration level."
    } else {
        write-host "Connected to hMailServer`n"
    }
}
Write-host "Cmdlets added:`n$(Get-Command | where {$_.ModuleName -eq 'hMailServer.PSHelper'})`n"