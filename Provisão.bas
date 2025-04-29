Attribute VB_Name = "Provisão"
' ----- Version -----
'        1.2.1
' -------------------

Sub SaveData(Optional ShowOnMacroList As Boolean = False)
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim tblRow As ListRow
    Dim newID As String
    Dim userResponse As VbMsgBoxResult
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    newID = wsForm.OLEObjects("ComboBoxID").Object.Value
    
    ' If ComboBoxID is not empty, prompt the user
    If Trim(newID) <> "" Then
        userResponse = MsgBox("Esse aditivo já foi cadastrado. Deseja sobrescrever?", vbYesNoCancel + vbQuestion, "Confirmação")

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
        .Cells(1, 1).Value = newID ' First column value
        
        ' Read column B values
        .Cells(1, 2).Value = wsForm.Range("B6").Value
        .Cells(1, 3).Value = wsForm.Range("B10").Value
        .Cells(1, 4).Value = wsForm.Range("B14").Value
        .Cells(1, 5).Value = wsForm.Range("B18").Value
        .Cells(1, 6).Value = wsForm.Range("B24").Value
        .Cells(1, 7).Value = wsForm.Range("B28").Value
        .Cells(1, 8).Value = wsForm.Range("B32").Value
        .Cells(1, 9).Value = wsForm.Range("B36").Value
        If .Cells(1, 10).Formula = "" Then
            .Cells(1, 10).Formula = "=IFERROR([@[Valor MDS]]/[@[Custo Atual Disponível]];"")"
        End If
        If .Cells(1, 11).Formula = "" Then
            .Cells(1, 11).Formula = "=[@[Custo Atual Disponível]]-[@[Valor MDS]]"
        End If
        
        ' Read column D values
        .Cells(1, 12).Value = wsForm.Range("D6").Value
        .Cells(1, 13).Value = wsForm.Range("D10").Value
        .Cells(1, 14).Value = wsForm.Range("D14").Value
        .Cells(1, 15).Value = wsForm.Range("D18").Value
        .Cells(1, 16).Value = wsForm.Range("D22").Value
        .Cells(1, 17).Value = wsForm.Range("D26").Value
        .Cells(1, 18).Value = wsForm.Range("D30").Value
        .Cells(1, 19).Value = wsForm.Range("D34").Value
        .Cells(1, 20).Value = wsForm.Range("D38").Value
        
        ' Read column F values
        .Cells(1, 21).Value = wsForm.Range("F6").Value
        .Cells(1, 22).Value = wsForm.Range("F10").Value
        .Cells(1, 23).Value = wsForm.Range("F14").Value
        .Cells(1, 24).Value = wsForm.Range("F18").Value
        .Cells(1, 25).Value = "" 'Clear date if ovewriten in case an e-mail was already sent
        .Cells(1, 26).Value = wsForm.Range("F22").Value
        
    End With
    
    ' MsgBox "Dados salvos com sucesso!", vbInformation
End Sub

Sub RetrieveDataFromName(Optional ShowOnMacroList As Boolean = False)
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchName As String
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
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
        If cell.Value & " - " & cell.Offset(0, 1).Value & " - " & cell.Offset(0, 11).Value = searchName Then
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
        wsForm.OLEObjects("ComboBoxID").Object.Value = foundRow.Value
    
        ' Read column B values
        .Range("B6").Value = foundRow.Offset(0, 1).Value
        .Range("B10").Value = foundRow.Offset(0, 2).Value
        .Range("B14").Value = foundRow.Offset(0, 3).Value
        .Range("B18").Value = foundRow.Offset(0, 4).Value
        .Range("B24").Value = foundRow.Offset(0, 5).Value
        .Range("B28").Value = foundRow.Offset(0, 6).Value
        .Range("B32").Value = foundRow.Offset(0, 7).Value
        .Range("B36").Value = foundRow.Offset(0, 8).Value
        
        ' Read column D values
        .Range("D6").Value = foundRow.Offset(0, 11).Value
        .Range("D10").Value = foundRow.Offset(0, 12).Value
        .Range("D14").Value = foundRow.Offset(0, 13).Value
        .Range("D18").Value = foundRow.Offset(0, 14).Value
        .Range("D22").Value = foundRow.Offset(0, 15).Value
        .Range("D26").Value = foundRow.Offset(0, 16).Value
        .Range("D30").Value = foundRow.Offset(0, 17).Value
        .Range("D34").Value = foundRow.Offset(0, 18).Value
        .Range("D38").Value = foundRow.Offset(0, 19).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Offset(0, 20).Value
        .Range("F10").Value = foundRow.Offset(0, 21).Value
        .Range("F14").Value = foundRow.Offset(0, 22).Value
        .Range("F18").Value = foundRow.Offset(0, 23).Value
        .Range("F22").Value = foundRow.Offset(0, 25).Value
    End With
End Sub

Sub RetrieveDataFromID(Optional ShowOnMacroList As Boolean = False)
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchID As Double
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
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
        MsgBox "ID não encontrado!", vbExclamation
        Exit Sub
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        wsForm.OLEObjects("ComboBoxName").Object.Value = foundRow.Offset(0, 1).Value & " - " & foundRow.Offset(0, 2).Value & " - " & foundRow.Offset(0, 11).Value
        
        ' Read column B values
        .Range("B6").Value = foundRow.Offset(0, 1).Value
        .Range("B10").Value = foundRow.Offset(0, 2).Value
        .Range("B14").Value = foundRow.Offset(0, 3).Value
        .Range("B18").Value = foundRow.Offset(0, 4).Value
        .Range("B24").Value = foundRow.Offset(0, 5).Value
        .Range("B28").Value = foundRow.Offset(0, 6).Value
        .Range("B32").Value = foundRow.Offset(0, 7).Value
        .Range("B36").Value = foundRow.Offset(0, 8).Value
        
        ' Read column D values
        .Range("D6").Value = foundRow.Offset(0, 11).Value
        .Range("D10").Value = foundRow.Offset(0, 12).Value
        .Range("D14").Value = foundRow.Offset(0, 13).Value
        .Range("D18").Value = foundRow.Offset(0, 14).Value
        .Range("D22").Value = foundRow.Offset(0, 15).Value
        .Range("D26").Value = foundRow.Offset(0, 16).Value
        .Range("D30").Value = foundRow.Offset(0, 17).Value
        .Range("D34").Value = foundRow.Offset(0, 18).Value
        .Range("D38").Value = foundRow.Offset(0, 19).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Offset(0, 20).Value
        .Range("F10").Value = foundRow.Offset(0, 21).Value
        .Range("F14").Value = foundRow.Offset(0, 22).Value
        .Range("F18").Value = foundRow.Offset(0, 23).Value
        .Range("F22").Value = foundRow.Offset(0, 25).Value
    End With
End Sub

Sub EnviarParaAprovação(Optional ShowOnMacroList As Boolean = False)

    SaveData
    
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
        MsgBox "O Outlook não está instalado nesse computador.", vbExclamation
        Exit Sub
    End If
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    ' Get the ID to search from ComboBox
    searchID = wsForm.OLEObjects("ComboBoxID").Object.Value
    
    ' Search for the ID in the first column of the table
    Set foundRow = Nothing
    On Error Resume Next
    Set foundRow = dadosTable.ListColumns(1).DataBodyRange.Find(What:=searchID, LookAt:=xlWhole)
    On Error GoTo 0
    
    ' If ID is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "ID não encontrado!", vbExclamation
        Exit Sub
    End If
    
    If foundRow.Offset(0, 24).Value <> "" Then
        userResponse = MsgBox("O e-mail de aprovação para esses dados já foi enviado em " & foundRow.Offset(0, 25).Value & ". Deseja enviar novamente?", vbYesNo)
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
    HTMLbody = HTMLbody & "<p>Solicito sua confirmação (“De acordo”) quanto aos valores abaixo, para que possamos dar continuidade à contratação da " & _
        foundRow.Offset(0, 18).Value & " para o serviço descrito a seguir: " & foundRow.Offset(0, 11).Value & " da " & foundRow.Offset(0, 1).Value & _
        " no valor de " & Format(foundRow.Offset(0, 6).Value, "R$ #,##0.00") & ". Todos os valores apresentados abaixo foram analisados pela equipe de Implantação/Suprimentos e considerado procedentes." & "</p>"
    
    ' Start the table
    HTMLbody = HTMLbody & "<table border='1' style='border-collapse: collapse; font-size: 10pt;'>"
    
    ' Title row
    HTMLbody = HTMLbody & "<tr style='background-color:#d9d9d9;'>"
    HTMLbody = HTMLbody & "<td colspan='2'><b>Provisão de Riscos" & " - " & foundRow.Offset(0, 1).Value & " - " & foundRow.Offset(0, 2).Value & "</b></td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 1) VALOR MDS
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>VALOR MDS</b></td>"
    ' Example: reading from the "Dados" sheet. Adjust the range as needed.
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 6).Value, "R$ #,##0.00") & " - que representa " & Format(foundRow.Offset(0, 9).Value, "##,##%") & " do Custo Atual disponível na Provisão de Riscos</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 2) CUSTO COT DO DR
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>CUSTO COT</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 7).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 3) CUSTO ATUAL DISPONÍVEL DO DR
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>CUSTO ATUAL DISPONÍVEL</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 8).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 4) SALDO RESIDUAL DO DR
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>SALDO RESIDUAL DO DR</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 10).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 5) Inserido no DR/tarefa
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>INSERIDO NO DR/TAREFA:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 5).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 6) Justificativa
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>JUSTIFICATIVA:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 12).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 7) Outros riscos já mapeados
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>OUTROS RISCOS JÁ MAPEADOS:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 19).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 8) Estágio da Obra
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>ESTÁGIO DA OBRA:</b></td>"
    
    If foundRow.Offset(0, 13).Value < 0.4 Then
        faseObra = "(Fase Inicial)"
    ElseIf foundRow.Offset(0, 13).Value < 0.8 Then
        faseObra = "(Fase Intermediária)"
    Else
        faseObra = "(Fase Final)"
    End If

    HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 13).Value * 100 & "% " & faseObra & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 9) Ação necessária
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>AÇÃO NECESSÁRIA:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 15).Value & "</td>"
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
        .Subject = "Aprovação de Custos - Provisão de Riscos - " & foundRow.Offset(0, 1).Value & " - " & foundRow.Offset(0, 2).Value
        .HTMLbody = HTMLbody & strSignature
        .Display   'Use .Display to just open the email draft
        ' .Send       'Use .Send to send immediately
    End With
    
    '--- Cleanup
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    foundRow.Offset(0, 24).Value = Date
    
    MsgBox "Email """ & "Aprovação de Custos - Provisão de Riscos - " & foundRow.Offset(0, 1).Value & " - " & foundRow.Offset(0, 2).Value & """ enviado com sucesso!", vbInformation
    
End Sub

Sub ClearForm(Optional ShowOnMacroList As Boolean = False)
    
    Dim wsForm As Worksheet
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    
    If wsForm.OLEObjects("ComboBoxID").Object.Value = "" Then
        If MsgBox("Esses dados não foram salvos. Deseja limpá-los mesmo assim?", vbYesNo) = vbNo Then
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
        .Range("B24").Value = ""
        .Range("B28").Value = ""
        .Range("B32").Value = ""
        .Range("B36").Value = ""
        
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
