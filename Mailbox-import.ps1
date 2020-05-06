import-Csv .\2015-05-13.csv | Foreach-Object{
   $user = New-MailboximportRequest -Mailbox $_.Login_ID_VH -FilePath "\\Phpukdagnode02\pst\$($_.Login_ID_VH).pst" -BadItemLimit 10 -BatchName 2015-05-13 -Name $_.Login_ID_VH
  }