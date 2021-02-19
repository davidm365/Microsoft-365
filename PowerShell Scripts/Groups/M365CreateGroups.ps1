New-UnifiedGroup -DisplayName "-2020OperationsGoals" -Alias -2020OperationsGoals
Set-UnifiedGroup -Identity “-2020OperationsGoals” -AccessType Private
Set-UnifiedGroup -Identity -2020OperationsGoals@nexion-health.com -HiddenFromAddressListsEnabled $true
Set-UnifiedGroup -Identity “-2020OperationsGoals” -UnifiedGroupWelcomeMessageEnabled:$False
Set-UnifiedGroup -Identity “-2020OperationsGoals” -HiddenFromExchangeClientsEnabled:$True
Set-UnifiedGroup -Identity -2020OperationsGoals@nexion-health.com -AutoSubscribeNewMembers

New-UnifiedGroup -DisplayName "-All-UsersNexion-Health" -Alias -AllUsersNexionHealth
Set-UnifiedGroup -Identity “-AllUsersNexionHealth” -AccessType Private
Set-UnifiedGroup -Identity -AllUsersNexionHealth@nexion-health.com -HiddenFromAddressListsEnabled $true
Set-UnifiedGroup -Identity “-AllUsersNexionHealth” -UnifiedGroupWelcomeMessageEnabled:$False
Set-UnifiedGroup -Identity “-AllUsersNexionHealth” -HiddenFromExchangeClientsEnabled:$True
Set-UnifiedGroup -Identity -AllUsersNexionHealth@nexion-health.com -AutoSubscribeNewMembers

New-UnifiedGroup -DisplayName "-All-UsersNexion-Health" -Alias -AllUsersNexionHealth
Set-UnifiedGroup -Identity “-AllUsersNexionHealth” -AccessType Private
Set-UnifiedGroup -Identity -AllUsersNexionHealth@nexion-health.com -HiddenFromAddressListsEnabled $true
Set-UnifiedGroup -Identity “-AllUsersNexionHealth” -UnifiedGroupWelcomeMessageEnabled:$False
Set-UnifiedGroup -Identity “-AllUsersNexionHealth” -HiddenFromExchangeClientsEnabled:$True
Set-UnifiedGroup -Identity -AllUsersNexionHealth@nexion-health.com -AutoSubscribeNewMembers