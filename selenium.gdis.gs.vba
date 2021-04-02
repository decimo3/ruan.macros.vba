Sub autoConsulta()

    Dim driver As WebDriver
    Dim lUltimaLinhaAtiva   As Long
    Dim lContador           As Long
    Dim ordemServiço        As String

    lUltimaLinhaAtiva = Worksheets("Sheet1").Cells(Worksheets("Sheet1").Rows.Count, 1).End(xlUp).Row

    Set driver = New Selenium.ChromeDriver

    driver.Get "https://gdis/gdis/portal/acesso/index/"
    driver.FindElementByXPath("//*[@id='details-button']").Click
    driver.FindElementByXPath("//*[@id='proceed-link']").Click
    driver.FindElementByXPath("//*[@id='username']").SendKeys 'You username
    driver.FindElementByXPath("//*[@id='password']").SendKeys 'You password
    driver.FindElementByXPath("//*[@id='fm1']/input[4]").Click

    For lContador = 2 To lUltimaLinhaAtiva
        If Range("F" & lContador).Value = "POSB" Then ' Or Range("F" & lContador).Value = "ENVI"
            driver.Get "https://gdis/ACWeb/pages/jsp/solicitacao/cadastros/recupera.jsp?IdNsCcs=" & Range("A" & lContador).Value
            ordemServiço = driver.FindElementByXPath("/html/body/div[1]/table/tbody/tr/td[5]").Text
            driver.Get "https://gdis/gdis-gs-web/ordemservico/detalhe_os?numeroServico=" & ordemServiço
            If Not driver.FindElementByXPath("//*[@id='btnSituacaoServico']").Text = "Cancelado" Or driver.FindElementByXPath("//*[@id='btnSituacaoServico']").Text = "Devolvido" Then
                Range("I" & lContador).Value = driver.FindElementByXPath("//*[@id='geral']/div[1]/div[11]/dl/dd").Text
                driver.FindElementByXPath("//*[@id='codigosFechamentoTab']").Click
                Range("H" & lContador).Value = driver.FindElementByXPath("//*[@id='codigosFechamento']/div/table/tbody/tr/td[1]/span").Text
                Range("G" & lContador).Value = driver.FindElementByXPath("//*[@id='codigosFechamento']/div/table/tbody/tr/td[2]/span").Text
            Else
                Range("F" & lContador).Value = "CANC"
            End If
        End If
    Next lContador
    
    driver.Quit
    MsgBox ("Acabou!")

End Sub
Sub observacao()

    Dim driver As WebDriver
    Dim lUltimaLinhaAtiva   As Long
    Dim lContador           As Long
    Dim ordemServiço        As String

    lUltimaLinhaAtiva = Worksheets("Planilha1").Cells(Worksheets("Planilha1").Rows.Count, 1).End(xlUp).Row

    Set driver = New Selenium.ChromeDriver

    driver.Get "https://gdis/gdis/portal/acesso/index/"
    driver.FindElementByXPath("//*[@id='details-button']").Click
    driver.FindElementByXPath("//*[@id='proceed-link']").Click
    driver.FindElementByXPath("//*[@id='username']").SendKeys "2258038"
    driver.FindElementByXPath("//*[@id='password']").SendKeys "Light@24"
    driver.FindElementByXPath("//*[@id='fm1']/input[4]").Click

    For lContador = 2 To lUltimaLinhaAtiva
            driver.Get "https://gdis/gdis-gs-web/ordemservico/detalhe_os?numeroServico=" & Range("B" & lContador).Value
            driver.FindElementByXPath("//*[@id='retornoExecucaoTab']").Click
            Range("C" & lContador).Value = driver.FindElementByXPath("//*[@id='retornoExecucao']/div[2]/div[5]/dl/dd/span").Text
    Next lContador
    
    driver.Quit
    MsgBox ("Acabou!")

End Sub
