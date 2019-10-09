

MsgBox GetRandomChars


Public Function GetRandomChars ()

Set ObjFSO = CreateObject("Scripting.FileSystemObject")

upperlimit = 50000
lowerlimit = 1

Randomize
RndChrs = Int((upperlimit - lowerlimit + 1) * Rnd() + lowerlimit)
TmpName = Trim(Mid(ObjFSO.GetTempName, 4, 5))
strTmstp = Trim(Replace(Left(Time, 8), ":", ""))

GetRandomChars = (RndChrs & TmpName & strTmstp)

End Function