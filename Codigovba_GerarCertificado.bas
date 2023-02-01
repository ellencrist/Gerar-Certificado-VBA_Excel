Attribute VB_Name = "Módulo1"
Sub Gerador_Certificados()
'declarando as dimensões com dim = dimensão
Dim Nome As String
Dim UL As Integer
Dim i As Integer

Dim Caminho As String
Dim Arquivo As String

'local para salvar o certificado
Caminho = ThisWorkbook.Path & "\Certificados\"
'Caminho = "C:\Users\usuario\OneDrive\Documentos\Documentos\Certificados\"

'ul = ultima linha
UL = abaAlunos.Range("L5")

'loop da linha 3 até a ultima linha
For i = 3 To UL

'pegar o nome da celula B3 até a ultima linha
Nome = abaAlunos.Range("B" & i)

'pegar valor da variavel nome e colocar na celula L4
abaAlunos.Range("L4") = Nome

'usado como suporte ao salvar o arquivo, pegar o nome e adicionar.pdf no final
Arquivo = Nome & ".pdf"

'salvar o pdf
abaCertificado.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Caminho & Arquivo

Next i

MsgBox "Certificado gerado com sucesso!"








End Sub

