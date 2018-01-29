# Macro-FCA: MOPAR
Desenvolvimento de uma simples macro de importação  de dados 
' Macro utilizada para fazer a importação dos dados.
Sub Importar()

Dim linhas As Long          ' Será usado para encontrar a quantidade de linhas do arquivo de importação.
Dim planAtual As String     ' Guarda o nome da planilha atual que estamos trabalhando.
Dim planImport As String    ' Guarda o nome do arquivo de importação que estaremos copiando as informações.

' Executa a limpeza de todos os dados atuais.
Call Limpeza.Limpeza

' Grava o nome da planilha atual.
planAtual = ActiveWorkbook.Name

' Apresenta a janela para selecionar o arquivo a ser aberto.
nomeArquivo = Application.GetOpenFilename(Title:="Escolha o arquivo de importacao")

' Se nenhum arquivo for selecionado então ele sai da macro.
If nomeArquivo = False Then Exit Sub

' Caso o arquivo seja selecionado corretamente, então o arquivo é aberto.
Workbooks.Open Filename:=nomeArquivo

' Aguarda 3 segundos só para ter certeza que o arquivo abriu.
Application.Wait Now + TimeValue("00:00:03")

' Grava o nome da planilha contendo os dados a serem importados.
planImport = ActiveWorkbook.Name

' Encontra quantas linhas contem o arquivo.
Range("A2").Select
linhas = Selection.End(xlDown).Row

' Copia as informações do chassi.
Range("B2:B" & linhas).Copy

' Volta para a planilha e cola as informações do chassi.
Workbooks(planAtual).Activate
Range("A3").PasteSpecial xlPasteValues

' Copia as informações de KM.
Workbooks(planImport).Activate
Range("V2:V" & linhas).Copy

' Volta para a planilha e cola as informações de KM.
Workbooks(planAtual).Activate
Range("B3").PasteSpecial xlPasteValues

' Copia as informações do tipo de garantia, código local e status.
Workbooks(planImport).Activate
Range("S2:U" & linhas).Copy

' Volta para a planilha e cola as informações do tipo de garantia, código local e status.
Workbooks(planAtual).Activate
Range("C3").PasteSpecial xlPasteValues

' Copia as informações do código A01. ##### ATENÇÃO: NÃO SEI SE ESTOU PEGANDO A COLUNA CORRETA ####
Workbooks(planImport).Activate
Range("W2:W" & linhas).Copy

' Volta para a planilha e cola as informações do código A01.
Workbooks(planAtual).Activate
Range("F3").PasteSpecial xlPasteValues

' Copia as informações do histórico de garantia ##### ATENÇÃO: NÃO SEI SE ESTOU PEGANDO A COLUNA CORRETA ####
Workbooks(planImport).Activate
Range("R2:R" & linhas).Copy

' Volta para a planilha e cola as informações do histórico de garantia.
Workbooks(planAtual).Activate
Range("G3").PasteSpecial xlPasteValues

' Copia as informações do conc reparadora.
Workbooks(planImport).Activate
Range("Q2:Q" & linhas).Copy

' Volta para a planilha e cola as informações do conc reparadora.
Workbooks(planAtual).Activate
Range("H3").PasteSpecial xlPasteValues

Range("A1").Select

' Encerra com a cópia da coluna para poder fechar a planilha corretamente
Application.CutCopyMode = False

' Fecha o arquivo de importação sem salvar.
Workbooks(planImport).Close savechanges:=False

End Sub
