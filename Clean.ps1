#!/usr/bin/env pwsh
remove-item -rec .paket -erroraction silentlycontinue
remove-item -rec paket-files -erroraction silentlycontinue
gci -recurse . -directory | where name -match "^(bin|obj)$" | select fullname | foreach { remove-item -rec -force $_.FullName}
