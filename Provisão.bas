Attribute VB_Name = "Provis�o"
' ----- Version -----
'        1.4.0
' -------------------

Sub SaveData(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim tblRow As ListRow
    Dim newID As String
    Dim userResponse As VbMsgBoxResult
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formul�rio")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' n�o encontrada!", vbExclamation
        Exit Sub
    End If
    
    newID = wsForm.OLEObjects("ComboBoxID").Object.Value
    
    ' If ComboBoxID is not empty, prompt the user
    If Trim(newID) <> "" Then
        userResponse = MsgBox("J� existe um aditivo com essa ID. Deseja sobreescrever?", vbYesNoCancel + vbQuestion, "Confirma��o")

        Select Case userResponse
            Case vbYes
                newID = Val(newID) ' Use ComboBoxID.Value as new ID
                ' Search for the ID in the first column of the table
                Set tblRow = dadosTable.ListRows(dadosTable.ListColumns(1).DataBodyRange.Find(What:=newID, LookAt:=xlWhole).Row - dadosTable.DataBodyRange.Row + 1)
            Case vbNo
                ' Proceed with generating new ID
                newID = Application.WorksheetFunction.Max(dadosTable.ListColumns(1).DataBodyRange) + 1
                wsForm.OLEObjects("ComboBoxID").Object.Value = newID
                ' Add a new row to the table
                Set tblRow = dadosTable.ListRows.Add
            Case vbCancel
                Exit Sub ' Exit without saving
        End Select
    Else
        newID = Application.WorksheetFunction.Max(dadosTable.ListColumns(1).DataBodyRange) + 1
        
        wsForm.OLEObjects("ComboBoxID").Object.Value = newID
        
        wsForm.OLEObjects("ComboBoxName").Object.Value = wsForm.Range("B6").Value & " - " & wsForm.Range("B10").Value & " - " & wsForm.Range("D6").Value
        
        ' Add a new row to the table
        Set tblRow = dadosTable.ListRows.Add
    End If
    
    ' Assign values to the new row
    With tblRow.Range
        ' Set new ID
        .Cells(1, colMap("ID")).Value = newID ' First column value
        
        ' Read column B values
        .Cells(1, colMap("Nome da Obra")).Value = wsForm.Range("B6").Value
        .Cells(1, colMap("Cliente")).Value = wsForm.Range("B10").Value
        .Cells(1, colMap("Tipo de Empreendimento")).Value = wsForm.Range("B14").Value
        .Cells(1, colMap("PM Respons�vel")).Value = wsForm.Range("B18").Value
        .Cells(1, colMap("PEP")).Value = wsForm.Range("B22").Value
        .Cells(1, colMap("DR Atividade")).Value = wsForm.Range("B28").Value
        .Cells(1, colMap("Valor MDS")).Value = wsForm.Range("B32").Value
        .Cells(1, colMap("Valor MDS L�quido")).Formula = "=[@[Valor MDS]]*0.9075"
        .Cells(1, colMap("Custo COT")).Value = wsForm.Range("B36").Value
        .Cells(1, colMap("Custo Atual Dispon�vel")).Value = wsForm.Range("B40").Value
    
        If .Cells(1, colMap("Impacto no COT")).Formula = "" Then
            .Cells(1, colMap("Impacto no COT")).Formula = "=IFERROR([@[Valor MDS]]/[@[Custo Atual Dispon�vel]];"")"
        End If
        If .Cells(1, colMap("Saldo Residual")).Formula = "" Then
            .Cells(1, colMap("Saldo Residual")).Formula = "=[@[Custo Atual Dispon�vel]]-[@[Valor MDS]]"
        End If
    
        ' Read column D values
        .Cells(1, colMap("Descri��o Breve do Aditivo")).Value = wsForm.Range("D6").Value
        .Cells(1, colMap("Justificativa do Aditivo")).Value = wsForm.Range("D10").Value
        
        If wsForm.Range("D14").Value < 0.4 Then
            .Cells(1, colMap("Est�gio da Obra")).Value = _
                Format(wsForm.Range("D14").Value, "##.00%") & " (Fase Inicial)"
        ElseIf wsForm.Range("D14").Value < 0.8 Then
            .Cells(1, colMap("Est�gio da Obra")).Value = _
                Format(wsForm.Range("D14").Value, "##.00%") & " (Fase Intermedi�ria)"
        Else
            .Cells(1, colMap("Est�gio da Obra")).Value = _
                Format(wsForm.Range("D14").Value, "##.00%") & " (Fase Final)"
        End If
        
        .Cells(1, colMap("Fase da Obra")).Formula = _
            "=IF([@[Est�gio da Obra]]<0.4,""Inicial"",IF([@[Est�gio da Obra]]<0.8,""Intermedi�rio"",""Final""))"
        .Cells(1, colMap("Fator Motivador")).Value = wsForm.Range("D18").Value
        .Cells(1, colMap("Detalhamento do Fator Motivador")).Value = wsForm.Range("D22").Value
        .Cells(1, colMap("Repasssar os custos ao cliente")).Value = wsForm.Range("D26").Value
        .Cells(1, colMap("Justificativa do n�o repasse")).Value = wsForm.Range("D30").Value
        .Cells(1, colMap("Prestador de Servi�o (Quem executou)")).Value = wsForm.Range("D34").Value
        .Cells(1, colMap("Outros Riscos")).Value = wsForm.Range("D38").Value
    
        ' Read column F values
        .Cells(1, colMap("Status")).Value = wsForm.Range("F6").Value
        .Cells(1, colMap("N�mero da RFP")).Value = wsForm.Range("F10").Value
        .Cells(1, colMap("Respons�vel Suprimentos")).Value = wsForm.Range("F14").Value
        .Cells(1, colMap("Pedido de Compra")).Value = wsForm.Range("F18").Value
        .Cells(1, colMap("Data da Solicita��o")).Value = ""  ' Clear date if overwritten in case an e-mail was already sent
        .Cells(1, colMap("Observa��es")).Value = wsForm.Range("F22").Value
    End With
    
    ' MsgBox "Dados salvos com sucesso!", vbInformation
End Sub

Sub RetrieveDataFromName(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchName As String
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formul�rio")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' n�o encontrada!", vbExclamation
        Exit Sub
    End If
    
    wsForm.OLEObjects("ComboBoxName").Top = wsForm.OLEObjects("ComboBoxID").Top + 38
    wsForm.OLEObjects("ComboBoxName").Left = wsForm.OLEObjects("ComboBoxID").Left
    
    ' Get the ID to search from ComboBox
    If wsForm.OLEObjects("ComboBoxName").Object.Value <> "" Then
        searchName = wsForm.OLEObjects("ComboBoxName").Object.Value
    Else
        'ClearForm
        Exit Sub
    End If
    
    ' Search for the matching row
    Set foundRow = Nothing
    For Each cell In dadosTable.ListColumns(2).DataBodyRange
        If cell.Value & " - " & cell.Cells(1, colMap("Cliente") - 1).Value & " - " & cell.Cells(1, colMap("Descri��o Breve do Aditivo") - 1).Value = searchName Then
            Set foundRow = cell.Offset(0, -1)
            Exit For
        End If
    Next cell
    
    ' If Name is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "Nenhuma obra encontrada!", vbExclamation
        Exit Sub
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        wsForm.OLEObjects("ComboBoxName").Object.Value = foundRow.Cells(1, colMap("Nome da Obra")).Value & " - " & foundRow.Cells(1, colMap("Cliente")).Value & " - " & foundRow.Cells(1, colMap("Descri��o Breve do Aditivo")).Value
        
        ' Read column B values
        .Range("B6").Value = foundRow.Cells(1, colMap("Nome da Obra")).Value
        .Range("B10").Value = foundRow.Cells(1, colMap("Cliente")).Value
        .Range("B14").Value = foundRow.Cells(1, colMap("Tipo de Empreendimento")).Value
        .Range("B18").Value = foundRow.Cells(1, colMap("PM Respons�vel")).Value
        .Range("B22").Value = foundRow.Cells(1, colMap("PEP")).Value
        .Range("B28").Value = foundRow.Cells(1, colMap("DR Atividade")).Value
        .Range("B32").Value = foundRow.Cells(1, colMap("Valor MDS")).Value
        .Range("B36").Value = foundRow.Cells(1, colMap("Custo COT")).Value
        .Range("B40").Value = foundRow.Cells(1, colMap("Custo Atual Dispon�vel")).Value
        
        ' Read column D values
        .Range("D6").Value = foundRow.Cells(1, colMap("Descri��o Breve do Aditivo")).Value
        .Range("D10").Value = foundRow.Cells(1, colMap("Justificativa do Aditivo")).Value
        .Range("D14").Value = foundRow.Cells(1, colMap("Est�gio da Obra")).Value
        .Range("D18").Value = foundRow.Cells(1, colMap("Fator Motivador")).Value
        .Range("D22").Value = foundRow.Cells(1, colMap("Detalhamento do Fator Motivador")).Value
        .Range("D26").Value = foundRow.Cells(1, colMap("Repasssar os custos ao cliente")).Value
        .Range("D30").Value = foundRow.Cells(1, colMap("Justificativa do n�o repasse")).Value
        .Range("D34").Value = foundRow.Cells(1, colMap("Prestador de Servi�o (Quem executou)")).Value
        .Range("D38").Value = foundRow.Cells(1, colMap("Outros Riscos")).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Cells(1, colMap("Status")).Value
        .Range("F10").Value = foundRow.Cells(1, colMap("N�mero da RFP")).Value
        .Range("F14").Value = foundRow.Cells(1, colMap("Respons�vel Suprimentos")).Value
        .Range("F18").Value = foundRow.Cells(1, colMap("Pedido de Compra")).Value
        .Range("F22").Value = foundRow.Cells(1, colMap("Observa��es")).Value
    End With
End Sub

Sub RetrieveDataFromID(Optional ShowOnMacroList As Boolean = False)

    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchID As Double
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formul�rio")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' n�o encontrada!", vbExclamation
        Exit Sub
    End If
    
    wsForm.OLEObjects("ComboBoxName").Top = wsForm.OLEObjects("ComboBoxID").Top + 38
    wsForm.OLEObjects("ComboBoxName").Left = wsForm.OLEObjects("ComboBoxID").Left
    
    ' Get the ID to search from ComboBox
    If wsForm.OLEObjects("ComboBoxID").Object.Value <> "" Then
        searchID = wsForm.OLEObjects("ComboBoxID").Object.Value
    Else
        'ClearForm
        Exit Sub
    End If
    
    ' Search for the ID in the first column of the table
    Set foundRow = Nothing
    On Error Resume Next
    Set foundRow = dadosTable.ListColumns(1).DataBodyRange.Find(What:=searchID, LookAt:=xlWhole)
    On Error GoTo 0
    
    ' If ID is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "ID n�o encontrado!", vbExclamation
        Exit Sub
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        wsForm.OLEObjects("ComboBoxName").Object.Value = foundRow.Cells(1, colMap("Nome da Obra")).Value & " - " & foundRow.Cells(1, colMap("Cliente")).Value & " - " & foundRow.Cells(1, colMap("Descri��o Breve do Aditivo")).Value
        
        ' Read column B values
        .Range("B6").Value = foundRow.Cells(1, colMap("Nome da Obra")).Value
        .Range("B10").Value = foundRow.Cells(1, colMap("Cliente")).Value
        .Range("B14").Value = foundRow.Cells(1, colMap("Tipo de Empreendimento")).Value
        .Range("B18").Value = foundRow.Cells(1, colMap("PM Respons�vel")).Value
        .Range("B22").Value = foundRow.Cells(1, colMap("PEP")).Value
        .Range("B28").Value = foundRow.Cells(1, colMap("DR Atividade")).Value
        .Range("B32").Value = foundRow.Cells(1, colMap("Valor MDS")).Value
        .Range("B36").Value = foundRow.Cells(1, colMap("Custo COT")).Value
        .Range("B40").Value = foundRow.Cells(1, colMap("Custo Atual Dispon�vel")).Value
        
        ' Read column D values
        .Range("D6").Value = foundRow.Cells(1, colMap("Descri��o Breve do Aditivo")).Value
        .Range("D10").Value = foundRow.Cells(1, colMap("Justificativa do Aditivo")).Value
        .Range("D14").Value = foundRow.Cells(1, colMap("Est�gio da Obra")).Value
        .Range("D18").Value = foundRow.Cells(1, colMap("Fator Motivador")).Value
        .Range("D22").Value = foundRow.Cells(1, colMap("Detalhamento do Fator Motivador")).Value
        .Range("D26").Value = foundRow.Cells(1, colMap("Repasssar os custos ao cliente")).Value
        .Range("D30").Value = foundRow.Cells(1, colMap("Justificativa do n�o repasse")).Value
        .Range("D34").Value = foundRow.Cells(1, colMap("Prestador de Servi�o (Quem executou)")).Value
        .Range("D38").Value = foundRow.Cells(1, colMap("Outros Riscos")).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Cells(1, colMap("Status")).Value
        .Range("F10").Value = foundRow.Cells(1, colMap("N�mero da RFP")).Value
        .Range("F14").Value = foundRow.Cells(1, colMap("Respons�vel Suprimentos")).Value
        .Range("F18").Value = foundRow.Cells(1, colMap("Pedido de Compra")).Value
        .Range("F22").Value = foundRow.Cells(1, colMap("Observa��es")).Value
    End With
End Sub

Sub EnviarParaAprova��o(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    '--- Variables for email content
    Dim HTMLbody As String
    Dim greeting As String
    Dim strSignature As String
    Dim faseObra As String
    
    '--- Create Outlook instance and a new mail item
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error GoTo 0
    
    If OutApp Is Nothing Then
        MsgBox "O Outlook n�o est� instalado nesse computador.", vbExclamation
        Exit Sub
    End If
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formul�rio")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' n�o encontrada!", vbExclamation
        Exit Sub
    End If
    
    ' Get the ID to search from ComboBox
    searchID = wsForm.OLEObjects("ComboBoxID").Object.Value
    
    ' Stop if data not saved
    If searchID = "" Then
        MsgBox "Desculpe, salve os dados antes de gerar o e-mail", vbInformation, "Aten��o"
        Exit Sub
    End If
    
    ' Search for the ID in the first column of the table
    Set foundRow = Nothing
    On Error Resume Next
    Set foundRow = dadosTable.ListColumns(1).DataBodyRange.Find(What:=searchID, LookAt:=xlWhole)
    On Error GoTo 0
    
    ' If ID is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "ID n�o encontrado!", vbExclamation
        Exit Sub
    End If
    
    If foundRow.Cells(1, colMap("Data da Solicita��o")).Value <> "" Then
        userResponse = MsgBox("O e-mail de aprova��o para esses dados j� foi enviado em " & foundRow.Cells(1, colMap("Data da Solicita��o")).Value & ". Deseja enviar novamente?", vbYesNo)
        If userResponse = vbNo Then
            MsgBox "Envio de e-mail cancelado!", vbInformation
            Exit Sub
        End If
    End If
    
    ' Decide between Bom dia or Boa tarde
    If Hour(Now) < 12 Then
        greeting = "Bom dia"
    Else
        greeting = "Boa tarde"
    End If
    
    ' Get user signature
    With OutMail
        .Display ' This opens the email and loads the default signature
        strSignature = .HTMLbody ' Capture the signature
    End With
    
    HTMLbody = ""
    HTMLbody = HTMLbody & "<p>" & greeting & ", Moretti</p>"
    HTMLbody = HTMLbody & "<p>Solicito sua confirma��o (�De acordo�) quanto aos valores abaixo, para que possamos dar continuidade � contrata��o da " & _
        foundRow.Cells(1, colMap("Prestador de Servi�o (Quem executou)")).Value & " para o servi�o descrito a seguir: " & foundRow.Cells(1, colMap("Descri��o Breve do Aditivo")).Value & " da " & foundRow.Cells(1, colMap("Nome da Obra")).Value & _
        " no valor de " & Format(foundRow.Cells(1, colMap("Valor MDS")).Value, "R$ #,##0.00") & ". Todos os valores apresentados abaixo foram analisados pela equipe de Implanta��o/Suprimentos e considerado procedentes." & "</p>"
    
    ' Start the table
    HTMLbody = HTMLbody & "<table border='1' style='border-collapse: collapse; font-size: 10pt;'>"
    
    ' Title row
    HTMLbody = HTMLbody & "<tr style='background-color:#d9d9d9;'>"
    HTMLbody = HTMLbody & "<td colspan='2'><b>Provis�o de Riscos" & " - " & foundRow.Cells(1, colMap("Nome da Obra")).Value & " - " & foundRow.Cells(1, colMap("Cliente")).Value & "</b></td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 1) VALOR MDS
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>VALOR MDS</b></td>"
    ' Example: reading from the "Dados" sheet. Adjust the range as needed.
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Valor MDS")).Value, "R$ #,##0.00") & " - que representa " & Format(foundRow.Cells(1, colMap("Impacto no COT")).Value, "#0.00%") & " do Custo Atual dispon�vel na Provis�o de Riscos</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 2) CUSTO COT DO DR
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>VALOR ORIGINAL DO DR (COT)</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Custo COT")).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 3) CUSTO ATUAL DISPON�VEL DO DR
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>SALDO ATUAL DO DR</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Custo Atual Dispon�vel")).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 4) SALDO RESIDUAL DO DR
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>SALDO RESIDUAL DO DR AP�S MDS</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Saldo Residual")).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 5) Inserido no DR/tarefa
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>INSERIDO NO DR/TAREFA:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("DR Atividade")).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 6) Justificativa
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>JUSTIFICATIVA:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("Justificativa do Aditivo")).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 7) Outros riscos j� mapeados
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>OUTROS RISCOS J� MAPEADOS:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("Outros Riscos")).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 8) Est�gio da Obra
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>EST�GIO DA OBRA:</b></td>"
    
    If IsNumeric(foundRow.Cells(1, colMap("Est�gio da Obra")).Value) Then
        If foundRow.Cells(1, colMap("Est�gio da Obra")).Value < 0.4 Then
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Est�gio da Obra")).Value, "##.00%") & " (Fase Inicial)" & "</td>"
        ElseIf foundRow.Cells(1, colMap("Est�gio da Obra")).Value < 0.8 Then
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Est�gio da Obra")).Value, "##.00%") & " (Fase Intermedi�ria)" & "</td>"
        Else
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Est�gio da Obra")).Value, "##.00%") & " (Fase Final)" & "</td>"
        End If
    Else
        HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("Est�gio da Obra")).Value & "</td>"
    End If

    HTMLbody = HTMLbody & "</tr>"
    
    ' 9) A��o necess�ria
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>A��O NECESS�RIA:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("Detalhamento do Fator Motivador")).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' Close the table
    HTMLbody = HTMLbody & "</table>"
    
    '-------------------------------------------------------------------------
    ' Configure and send the email
    '-------------------------------------------------------------------------
    With OutMail
        .To = "emoretti@weg.net"
        .CC = "matheusp@weg.net"
        .BCC = ""
        .Subject = "Aprova��o de Custos - Provis�o de Riscos - " & foundRow.Cells(1, colMap("Nome da Obra")).Value & " - " & foundRow.Cells(1, colMap("Cliente")).Value
        .HTMLbody = HTMLbody & strSignature
        .Display   'Use .Display to just open the email draft
        ' .Send       'Use .Send to send immediately
    End With
    
    '--- Cleanup
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    foundRow.Cells(1, colMap("Data da Solicita��o")).Value = Date
    
    MsgBox "Email """ & "Aprova��o de Custos - Provis�o de Riscos - " & foundRow.Cells(1, colMap("Nome da Obra")).Value & " - " & foundRow.Cells(1, colMap("Cliente")).Value & """ enviado com sucesso!", vbInformation
    
End Sub

Sub ClearForm(Optional ShowOnMacroList As Boolean = False)
    
    Dim wsForm As Worksheet
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formul�rio")
    
    If wsForm.OLEObjects("ComboBoxID").Object.Value = "" Then
        If MsgBox("Esses dados n�o foram salvos. Deseja limp�-los mesmo assim?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        .OLEObjects("ComboBoxID").Object.Value = ""
        .OLEObjects("ComboBoxName").Object.Value = ""
        .OLEObjects("ComboBoxName").Width = 123
        
        ' Read column B values
        .Range("B6").Value = ""
        .Range("B10").Value = ""
        .Range("B14").Value = ""
        .Range("B18").Value = ""
        .Range("B22").Value = ""
        .Range("B28").Value = ""
        .Range("B32").Value = ""
        .Range("B36").Value = ""
        .Range("B40").Value = ""
        
        ' Read column D values
        .Range("D6").Value = ""
        .Range("D10").Value = ""
        .Range("D14").Value = ""
        .Range("D18").Value = ""
        .Range("D22").Value = ""
        .Range("D26").Value = ""
        .Range("D30").Value = ""
        .Range("D34").Value = ""
        .Range("D38").Value = ""
        
        ' Read column F values
        .Range("F6").Value = ""
        .Range("F10").Value = ""
        .Range("F14").Value = ""
        .Range("F18").Value = ""
        .Range("F22").Value = ""
    End With
End Sub

Public Function GetColumnHeadersMapping() As Object
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    
    ' Add each header from the provided table to the dictionary,
    ' mapping it to its column position.
    headers.Add "ID", 1
    headers.Add "Nome da Obra", 2
    headers.Add "Cliente", 3
    headers.Add "Tipo de Empreendimento", 4
    headers.Add "PM Respons�vel", 5
    headers.Add "PEP", 6
    headers.Add "DR Atividade", 7
    headers.Add "Valor MDS", 8
    headers.Add "Valor MDS L�quido", 9
    headers.Add "Custo COT", 10
    headers.Add "Custo Atual Dispon�vel", 11
    headers.Add "Impacto no COT", 12
    headers.Add "Saldo Residual", 13
    headers.Add "Descri��o Breve do Aditivo", 14
    headers.Add "Justificativa do Aditivo", 15
    headers.Add "Est�gio da Obra", 16
    headers.Add "Fase da Obra", 17
    headers.Add "Fator Motivador", 18
    headers.Add "Detalhamento do Fator Motivador", 19
    headers.Add "Repasssar os custos ao cliente", 20
    headers.Add "Justificativa do n�o repasse", 21
    headers.Add "Prestador de Servi�o (Quem executou)", 22
    headers.Add "Outros Riscos", 23
    headers.Add "Status", 24
    headers.Add "N�mero da RFP", 25
    headers.Add "Respons�vel Suprimentos", 26
    headers.Add "Pedido de Compra", 27
    headers.Add "Data da Solicita��o", 28
    headers.Add "Observa��es", 29
    
    Set GetColumnHeadersMapping = headers
End Function
