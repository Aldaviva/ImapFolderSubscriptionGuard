$binaryPathName = Resolve-Path ".\ImapFolderSubscriptionGuard\bin\Release\ImapFolderSubscriptionGuard.exe"

New-Service -Name "ImapFolderSubscriptionGuard" -DisplayName "IMAP Folder Subscription Guard" -Description "Automatically unsubscribe from and delete IMAP folders that opinionated clients keep trying to create." -BinaryPathName $binaryPathName.Path