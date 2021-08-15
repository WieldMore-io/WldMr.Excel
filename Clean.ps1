gci -recurse . -directory | where name -match "^(bin|obj)$" | select fullname | foreach { rm -rec $_.FullName}
