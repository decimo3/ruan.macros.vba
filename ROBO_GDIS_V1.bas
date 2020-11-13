Attribute VB_Name = "Módulo1"
Sub Gdis_ROBO()


Dim driver As WebDriver
    Dim lUltimaLinhaAtiva   As Long
    Dim lContador           As Long


'SHEETS(ABA)_______________________________________________________________________________________________________________________
Sheets("GDIS").Select


lUltimaLinhaAtiva = Worksheets("GDIS").Cells(Worksheets("GDIS").Rows.Count, 2).End(xlUp).Row

'Range("N2:AA2").Select
'    Selection.Copy
'    Range("n2:AA" & lUltimaLinhaAtiva).Select
 '     ActiveSheet.Paste

Set driver = New selenium.ChromeDriver 'PhantomJSDriver 'ChromeDriver


'PRIMEIRA PÁGINA__________________________________________________________________________________________________________________

driver.Get "https://gdis/gdis-gs-web/despacho/servico/andamento/create?numserv=" & "00000000" & "&idDespachanteFiltro=4400"

'LOGIN_____________________________________________________________________________________________________________________________
driver.FindElementByXPath("//*[@id='username']").SendKeys UserForm1.TextBox1

'SendKeys "4008576"

'SENHA______________________________________________________________________________________________________________________________
driver.FindElementByXPath("//*[@id='password']").SendKeys UserForm1.TextBox2
'SendKeys "GJtF0%pu"

'BOTÃO DE ACIONAMENTO (SUBMIT)______________________________________________________________________________________________________
driver.FindElementByXPath("//*[@id='fm1']/input[4]").Click

UserForm1.TextBox4 = lUltimaLinhaAtiva - 1

HoraAtual = TimeValue(Now)


'CONTADOR DE LINHAS, PONTO DE INICIO DA REPETIÇÃO____________________________________________________________________________________
For lContador = 2 To lUltimaLinhaAtiva

Range("g" & lContador).Select
UserForm1.TextBox3 = lContador - 1

UserForm1.TextBox5.Text = Format(TimeValue(Now) - HoraAtual, "hh:mm:ss")

'Range("g" & lContador).Select
On Error GoTo 0

If Range("g" & lContador) = "" Then

driver.Get "https://gdis/gdis-gs-web/ordemservico/detalhe_os?numeroServico=" & Range("f" & lContador)
If driver.FindElementByXPath("//*[@id='btnSituacaoServico']").Text = "Pendente" Then

'CARRINHO PENDENTE DE FINALIZAÇÃO(DIRECT LINK)_________________________________________________________________________________________


Range("N" & lContador).FormulaLocal = "=SE(DIA($L:$L)>9;"""";""0"")&DIA($L:$L)&SE(MÊS($L:$L)>9;""0"";""00"")&MÊS($L:$L)&""0""&ANO($L:$L)&SE(HORA(M:M)>9;""0"";""00"")&HORA(M:M)&"":""&SE(MINUTO(M:M)>9;"""";""0"")&MINUTO(M:M)&"":""&SE(SEGUNDO(M:M)>9;"""";""0"")&SEGUNDO(M:M)"
Range("O" & lContador).FormulaLocal = "=$M:$M-(""03:00:00"")"
Range("P" & lContador).FormulaLocal = "=SE(DIA($L:$L)>9;"""";""0"")&DIA($L:$L)&SE(MÊS($L:$L)>9;""0"";""00"")&MÊS($L:$L)&""0""&ANO($L:$L)&SE(HORA(O:O)>9;""0"";""00"")&HORA(O:O)&"":""&SE(MINUTO(O:O)>9;"""";""0"")&MINUTO(O:O)&"":""&SE(SEGUNDO(O:O)>9;"""";""0"")&SEGUNDO(O:O)"
Range("Q" & lContador).FormulaLocal = "=$M:$M-(""02:00:00"")"
Range("R" & lContador).FormulaLocal = "=SE(DIA($L:$L)>9;"""";""0"")&DIA($L:$L)&SE(MÊS($L:$L)>9;""0"";""00"")&MÊS($L:$L)&""0""&ANO($L:$L)&SE(HORA(Q:Q)>9;""0"";""00"")&HORA(Q:Q)&"":""&SE(MINUTO(Q:Q)>9;"""";""0"")&MINUTO(Q:Q)&"":""&SE(SEGUNDO(Q:Q)>9;"""";""0"")&SEGUNDO(Q:Q)"
Range("S" & lContador).FormulaLocal = "=$M:$M-(""01:00:00"")"
Range("T" & lContador).FormulaLocal = "=SE(DIA($L:$L)>9;"""";""0"")&DIA($L:$L)&SE(MÊS($L:$L)>9;""0"";""00"")&MÊS($L:$L)&""0""&ANO($L:$L)&SE(HORA(S:S)>9;""0"";""00"")&HORA(S:S)&"":""&SE(MINUTO(S:S)>9;"""";""0"")&MINUTO(S:S)&"":""&SE(SEGUNDO(S:S)>9;"""";""0"")&SEGUNDO(S:S)"
Range("U" & lContador).FormulaLocal = "=$M:$M-(""00:30:00"")"
Range("V" & lContador).FormulaLocal = "=SE(DIA($L:$L)>9;"""";""0"")&DIA($L:$L)&SE(MÊS($L:$L)>9;""0"";""00"")&MÊS($L:$L)&""0""&ANO($L:$L)&SE(HORA(U:U)>9;""0"";""00"")&HORA(U:U)&"":""&SE(MINUTO(U:U)>9;"""";""0"")&MINUTO(U:U)&"":""&SE(SEGUNDO(U:U)>9;"""";""0"")&SEGUNDO(U:U)"
Range("X" & lContador).FormulaLocal = "=ESQUERDA(H:H;4)"
Range("W" & lContador).FormulaLocal = "=SUBSTITUIR(X:X;""."";"""")"
Range("Z" & lContador).FormulaLocal = "=$M:$M-(""00:00:00"")"
Range("AA" & lContador).FormulaLocal = "=SE(DIA($L:$L)>9;"""";""0"")&DIA($L:$L)&SE(MÊS($L:$L)>9;""0"";""00"")&MÊS($L:$L)&""0""&ANO($L:$L)&SE(HORA(Z:Z)>9;""0"";""00"")&HORA(Z:Z)&"":""&SE(MINUTO(Z:Z)>9;"""";""0"")&MINUTO(Z:Z)&"":""&SE(SEGUNDO(Z:Z)>9;"""";""0"")&SEGUNDO(Z:Z)"



Range("N" & lContador & ":" & "AA" & lContador).Select
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False





driver.Get "https://gdis/gdis-gs-web/despacho/servico/andamento/create?numserv=" & Range("f" & lContador) & "&idDespachanteFiltro=4400"

On Error Resume Next

driver.FindElementByXPath("//*[@id='detalhesErroSpan']").Click

If Err.Number = "7" Then


On Error GoTo 0
On Error Resume Next

'CARRO_______________________________________________________________________________
driver.FindElementByXPath("//*[@id='formFechamento']/div[1]/div/div[2]/div/div[1]/div/div/button").Click
driver.FindElementByXPath("//*[@id='formFechamento']/div[1]/div/div[2]/div/div[1]/div/div/div/div/input").SendKeys Range("b" & lContador)
driver.FindElementByXPath("//*[@id='formFechamento']/div[1]/div/div[2]/div/div[1]/div/div/div/ul").Click

'MARCAÇÃO DE VIA VOZ____________________________________________________________________________________________________________________
driver.FindElementByXPath("//*[@id='viaVoz1']").Click
sng = Timer
Do While sng + 3 > Timer
Loop
'DESIGNAÇÃO_____________________________________________________________________________________________________________________________

driver.FindElementByXPath("//*[@id='servicoFTO.dhDesigna']").SendKeys Range("P" & lContador)
driver.FindElementByXPath("//*[@id='btnDesignar']").Click
sng = Timer
Do While sng + 3 > Timer
Loop
'ACIONAMENTO____________________________________________________________________________________________________________________________
driver.FindElementByXPath("//*[@id='servicoFTO.dhAciona']").SendKeys Range("R" & lContador)
driver.FindElementByXPath("//*[@id='btnAcionar']/i").Click
sng = Timer
Do While sng + 3 > Timer
Loop
'LOCALIZAÇÃO_____________________________________________________________________________________________________________________________
driver.FindElementByXPath("//*[@id='servicoFTO.dhLocali']").Clear
sng = Timer
Do While sng + 3 > Timer
Loop
driver.FindElementByXPath("//*[@id='servicoFTO.dhLocali']").SendKeys Range("T" & lContador)
driver.FindElementByXPath("//*[@id='btnLocalizar']/i").Click
sng = Timer
Do While sng + 3 > Timer
Loop
'PREVISÃO________________________________________________________________________________________________________________________________
driver.FindElementByXPath("//*[@id='servicoFTO.dhTerprev']").Clear
driver.FindElementByXPath("//*[@id='servicoFTO.dhTerprev']").SendKeys Range("V" & lContador)
driver.FindElementByXPath("//*[@id='dataAndamentoServico']/div/div[5]/div/dl/dd/div/div/button/i").Click
sng = Timer
Do While sng + 3 > Timer
Loop

'TERMINO REAL________________________________________________________________________________________________________________________________
driver.FindElementByXPath("//*[@id='servicoFTO.dhTerreal']").Clear

driver.FindElementByXPath("//*[@id='servicoFTO.dhTerreal']").SendKeys Range("AA" & lContador)
'driver.FindElementByXPath("//*[@id='btnAcionarFinalizacao']/i").Click

'//*[@id="dataAndamentoServico"]/div/div[5]/div/dl/dd/div/div/button/i

'//*[@id="servicoFTO.dhTerreal"]
'//*[@id="btnAcionarFinalizacao"]/i

On Error GoTo 0



'CODIGO___________________________________________________________________________________________________________________________________
driver.FindElementByXPath("//*[@id='findCodFech']").SendKeys Range("X" & lContador)




driver.FindElementByXPath("//*[@id='" & Range("W" & lContador) & "']").Click
'MATERIAL UTILIZADO________________________________________________________________________________________________________________________
On Error Resume Next

driver.FindElementByXPath("//*[@id='tab-materiais_usados']/a").Click
If Err.Number <> 7 Then
driver.FindElementByXPath("//*[@id='link-show-equipamentos']/i").Click

'TEMPORIZADOR______________________________________________________________________________________________________________________________
sng = Timer
Do While sng + 2 > Timer
Loop

'INICIO CONDICIONAL (VERDADEIRO)____________________________________________________________________________________________________________

driver.FindElementByXPath("//*[@id='tblMateriais_filter']/label/input").SendKeys Range("j" & lContador)


'PAGINA DE MATERIAIS UTILIZADOS______________________________________________________________________________________________________________
driver.FindElementByXPath("//*[@id='tblMateriais']/tbody/tr[1]/td[3]/button").Click


'# QUANT USADA (ZEREI, POIS TRATA-SE DE FINALIZAÇÃO DE CORTE)
driver.FindElementByXPath("//*[@id='txt-qtdeUsado']").SendKeys 0

'# QUANT RETIRADA (QUANTIDADE INFORMADA NA PLANILHA DO GFORMS)
driver.FindElementByXPath("//*[@id='txt-qtdeRetirado']").SendKeys Range("K" & lContador).Value

'# BOTÃO ADICIONAR
driver.FindElementByXPath("//*[@id='btn-adicionar']").Click

Else
End If



On Error GoTo 0
'DADOS ADICONAIS______________________________________________________________________________________________________________
driver.FindElementByXPath("//*[@id='tab-abaDadosAdicionais']/a").Click

driver.FindElementByXPath("//*[@id='txtObservacao']").SendKeys Range("C" & lContador) & " - " & Range("D" & lContador) & "; " & Range("I" & lContador)
driver.FindElementByXPath("//*[@id='btnFinalizar']").Click
Range("G" & lContador) = "FINALIZADA PELO ROBO, " & DateValue(Now) & "."







'FIM CONDICIONAL ___________________________________________________________________________________________________________________________



Else
On Error GoTo 0
Range("G" & lContador) = "ERRO NO NUMERO DA OS"

End If
Else
UserForm1.TextBox3 = lContador - 1

UserForm1.TextBox5.Text = Format(TimeValue(Now) - HoraAtual, "hh:mm:ss")
Range("G" & lContador) = "Não finalizado pelo robô, status consultado no GDIS: " & driver.FindElementByXPath("//*[@id='btnSituacaoServico']").Text
End If
Else
End If
'FIM DO CONTADOR (REINICIO)__________________________________________________________________________________________________________________

UserForm1.TextBox5.Text = Format(TimeValue(Now) - HoraAtual, "hh:mm:ss")
Range("g" & lContador).Select







Next lContador


'MENSAGEM DE FINALIZAÇÃO DO PROCESSO__________________________________________________________________________________________________________

driver.Quit

UserForm1.TextBox5.Text = Format(TimeValue(Now) - HoraAtual, "hh:mm:ss")

MsgBox "O relatório foi concluído em exatos: " & Format(TimeValue(Now) - HoraAtual, "hh:mm:ss") & ". " & Chr(13) & "OSC trabalhando com inteligência.", vbInformation, ""


End Sub
