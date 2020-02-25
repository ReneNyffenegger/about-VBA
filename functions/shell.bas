option explicit

global PID as long

sub startTask() ' {

   PID  = shell("notepad.exe", vbNormalNoFocus)
   debug.print("process id = " & PID)

end sub ' }

sub stopTask() ' {

    shell "taskkill /pid " & PID

end sub ' }
