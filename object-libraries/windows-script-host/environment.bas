option explicit

sub wsh_environments() ' {

    dim wsh     as new wshShell

    showEnv wsh, "User"     ' Values as stored in the Registry under HKCU\Environment
    showEnv wsh, "System"   ' Values as stored in the Registry under HKLM\System\CurrentControlSet\Control\Session Manager\Environment
    showEnv wsh, "Volatile" '
    showEnv wsh, "Process"  '

end sub ' }

sub showEnv(wsh as wshShell, env as string) ' {

    debug.print "Show values for environment " & env

    dim envItem as variant
    for each envItem in wsh.environment(env)
      '
      ' envItem is NAME=VALUE, for example
      ' TEMP=%USERPROFILE%\AppData\Local\Temp
      '
        debug.print "  " & envItem
    next envItem

end sub ' }
