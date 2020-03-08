function searchWin()
  bag = false
  wow = false
  fgfg = false
  set bag = getobject("winmgmts:\\.\root\cimv2")
  set wow = bag.execquery("select * from win32_process where name='WowClassic.exe'")
  For each i in wow
    fg=true
  next
  set bag = nothing
  set qq = nothing
  searchWin = fg
end function

set ws = createObject("wscript.shell")
ws.appactivate "魔兽世界"
Do While true
  If searchWin()=false Then Exit Do
  n=rnd() 
  n = int((n * 1000) * 80)
  wscript.sleep n
  ws.sendkeys " "
  dim b
  b = datediff("s","1970-01-01 00:00:00",now)
  a = b Mod 100
  If a >= 0 And a <= 60  And (int(b / 100) Mod 6) = 0  Then 
    ws.sendkeys "n"
    wscript.sleep 1000
    ws.sendkeys "n"
    wscript.sleep 50000
  End if
Loop
msgbox("afkproof stopped")




