<p align="center">
  <img alt="Repository size" src="https://img.shields.io/github/directory-file-count/marelps/excel-agendamento-de-consulta?style=flat-square">
  <a href="https://twitter.com/piterparquinho">
    <img alt="Siga no Twitter" src="https://img.shields.io/twitter/url?style=social&url=https%3A%2F%2Ftwitter.com%2Fpiterparquinho">
  </a>
  <img alt="Github last commit" src="https://img.shields.io/github/last-commit/marelps/excel-agendamento-de-consulta?style=flat-square">
   <img alt="License" src="https://img.shields.io/badge/license-MIT-brightgreen">
  <a href="https://rocketseat.com.br">
    <img alt="Feito por Vitória" src="https://img.shields.io/badge/feito%20por-Vitória-%237519C1">
  </a>

# Planilha de cadastro de prontuários em Excel Macro
<h4 align="center"> 
	✅ Planilha Concluída ✅
</h4>

<p align="center">
 <a href="#objetivo">Objetivo</a> •
 <a href="#como-usar">Como Usar</a> •  
 <a href="#autor">Autor</a> • 
  <a href="#licença">Licença</a> • 
 <a href="#readme">Versões do README</a>
</p>

## Objetivo
Essa planilha foi criada na época em que eu estagiava em um centro de infecctologia e enfrentava um problema desde o começo que era o cadastro de prontuário na planilha do Excel.

O problema do cadastro é que ele era feito por muitas pessoas com diferentes níveis de conhecimento em informática e excel, por isso, muitas das vezes a planilha sofria alguns problemas de digitação ou ordem das cores. Por conta disso criei um pequeno formulário para auxiliar no cadastro, onde é possível visualizar as cores de respectiva doença e a ordem da sua prioridade. 

Também configurei o preenchimento automático de sua respectiva cor de acordo com sua prioridade, através da formatação da condicional.

<p align="center">
<img src="imgs/condicional.png" alt="Formatação Condicional">
</p>

Com isso, a chance do preenchimento da cor errada se torna menor do que como era feita anteriormente e também se torna ainda mais rápido.


 ## Como Usar
 Através de um botão visto no topo da planilha, é possível abrir esse formulário e é aqui que era possível cadastrar os prontuários na recepção que eram repassados para nós pelos responsáveis pelo registro dos novos casos. 
 
 Após o preenchimento, o cadastro já aparecia pronto e formatado da maneira correta na planilha e com as cores já preenchidas também corretamente.

<p align="center">
   <img src="imgs/form.png" alt="Formulário">
</p>

Também é possível alterar a situação dos prontuário. Aqueles prontuários que não mexeriamos mais por algum motivo e seria levado para uma outra sala onde seria arquivado, estes prontuários ficam com a linha em evidência na planilha.

<p align="center">
<img src="imgs/planilha.png" alt="Planilha">
</p>

 ### Macros
 Macros utilizadas no formulário. Textos escritos com um ' no começo da linha, são alguns comentários que fiz para me localizar.
 ```
'Identifica o tipo do objeto e insere se for um dos tipos definidos
Private Sub lsInserir(ByRef lTextBox As Variant, ByVal Plan1 As String, ByVal lColunaCodigo As Long, ByVal lUltimaLinha As Long)
    If (TypeOf lTextBox Is MSForms.TextBox) Or (TypeOf lTextBox Is MSForms.ComboBox) Then
        Sheets(Plan1).Range(lTextBox.Tag & lUltimaLinha).Value = lTextBox.Text
    Else
        If TypeOf lTextBox Is MSForms.OptionButton Then
            If lTextBox.Value = True Then
                Sheets(Plan1).Range(lTextBox.Tag & lUltimaLinha).Value = lTextBox.Caption
            End If
        End If
    End If
End Sub

'Loop por todos os componentes da tela
'frmProntuario = Nome do UserForm atual
'Plan1 = Nome da planilha aonde irão ser inseridos os valores
'lColunaCodigo = Coluna de referência para a inserção dos dados
Public Function lsInserirTextBox(frmProntuario As UserForm, ByVal Plan1 As String, ByVal lColunaCodigo As Long)
    Dim controle            As Control
    Dim lUltimaLinhaAtiva   As Long
    
    lUltimaLinhaAtiva = Worksheets(Plan1).Cells(Worksheets(Plan1).Rows.Count, lColunaCodigo).End(xlUp).Row + 1
    
    For Each controle In frmProntuario.Controls
        lsInserir controle, Plan1, lColunaCodigo, lUltimaLinhaAtiva
    Next
End Function

'Limpa todos os objetos TextBox da tela
Public Function lsLimparTextBox(frmProntuario As UserForm)
    Dim controle            As Control
    
    For Each controle In frmProntuario.Controls
        If TypeOf controle Is MSForms.TextBox Then
            controle.Text = ""
        End If
    Next
End Function

'Aciona o botão de limpar
Private Sub CommandButton1_Click()
    lsLimparTextBox frmProntuario
    
    TextBox1.SetFocus
End Sub

'Aciona o botão de inserir
Private Sub CommandButton2_Click()
    lsInserirTextBox frmProntuario, "PRONTUARIO", 2
    
    lsLimparTextBox frmProntuario
    
    TextBox1.SetFocus
End Sub

Private Sub TextBox1_Change()
    TextBox1 = UCase(TextBox1)
    'Ucase = Upper case
End Sub

Private Sub TextBox2_Change()
    TextBox2 = UCase(TextBox2) 'Ucase = Upper case
End Sub

Private Sub TextBox3_Change()
    TextBox3 = UCase(TextBox3)
    'Ucase = Upper case
End Sub
 ```
 ***
Macro utilizada para chamar o formulário no botão localizado no topo da planilha

<p align="center">
   <img src="imgs/button.png" alt="Botão no topo da planilha">
</p>

```
Sub ChamarFormProntuario()
    frmProntuario.Show
End Sub
```

## Autor
<p align="center">
 <img style="border-radius: 50%;" src="https://avatars.githubusercontent.com/u/48718646?v=4" width="100px;" alt="Autora do projeto"/>
 <br />
 <sub><b>Vitória Garrucho</b></br> Feito com ❤️</sub></p>

<p align="center">Entre em contato através das minhas redes sociais!<br>
<a href="https://twitter.com/piterparquinho" target="_blank"><img src="https://img.shields.io/badge/-@piterparquinho-1ca0f1?style=flat-square&labelColor=1ca0f1&logo=twitter&logoColor=white&link=https://twitter.com/piterparquinho" alt="Twitter Badge"></a>
<a href="https://www.linkedin.com/in/vitoriagarrucho/" target="_blank"><img src="https://img.shields.io/badge/-Vitória-blue?style=flat-square&logo=Linkedin&logoColor=white&link=https://www.linkedin.com/in/vitoriagarrucho/" alt="Linkedin Badge"></a>
<a href="mailto:vitoriagarrucho@gmail.com" target="_blank"><img src="https://img.shields.io/badge/-vitoriagarrucho@gmail.com-c14438?style=flat-square&logo=Gmail&logoColor=white&link=mailto:vitoriagarrucho@gmail.com" alt="Gmail Badge"></a>
 </p>

## Licença

Este projeto esta sobe a licença [MIT](./LICENSE).

Feito com ❤️ por Vitória Garrucho

<a href="https://www.linkedin.com/in/vitoriagarrucho/" target="_blank">Entre em contato!</a>

## README
[Português](./README.md)  |  [English](./README-en.md)