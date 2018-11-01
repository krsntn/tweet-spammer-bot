set shell = createobject("wscript.shell")

strtimes = "1000"

wscript.sleep(2000)

for i=1 to strtimes

for x=1 to 8
shell.sendkeys "{Tab}"
next

' shell.SendKeys "{Enter}"
' wscript.sleep(1000)
shell.SendKeys("Keygen: " + RandomString() + "-" + RandomString() + "-" + RandomString() + "-" + RandomString())
wscript.sleep(500)

for x=1 to 6
shell.sendkeys "{Tab}"
next

shell.SendKeys "{Enter}"
wscript.sleep(4000)
Shell.SendKeys "{F5}"
wscript.sleep(86000)
next

function RandomString()

Randomize()

dim CharacterSetArray
CharacterSetArray = Array(_
Array(4, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") _
)

dim i
dim j
dim Count
dim Chars
dim Index
dim Temp

for i = 0 to UBound(CharacterSetArray)

Count = CharacterSetArray(i)(0)
Chars = CharacterSetArray(i)(1)

for j = i to Count

Index = Int(Rnd() * Len(Chars)) + 1
Temp = Temp & Mid(Chars, Index, 1)

next

next

dim TempCopy

do until Len(Temp) = 0

Index = Int(Rnd() * Len(Temp)) + 1
TempCopy = TempCopy & Mid(Temp, Index, 1)
Temp = Mid(Temp, 1, Index - 1) & Mid(Temp, Index + 1)

loop

RandomString = TempCopy

end function