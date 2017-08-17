# Movimentação

Este é um Google App voltado para movimentações financeiras gerais do Ramo Estudantil IEEE UFABC. Ele faz parte de um processo de automatização financeira que está sendo desenvolvido.

Suas funções são:
- Capturar os dados e o arquivo de recibo fornecidos em um formulário e depositar em um local pré determinado.
  - Os arquivos são renomeados e depositados numa pasta do Google Drive.
  - Os dados são organizados em uma Planilha no próprio Google
- Enviar um e-mail para a pessoa que submeteu o documento e para o Ramo Estudantil avisando que algo foi submetido.
  - O link do arquivo é fornecido via e-mail.

## Todo V1:

---------------------------------------------------------------------------

*Versão atual: RC1*

---------------------------------------------------------------------------

- [ ] Adicionar links de pastas de cada capitulo - Verificar necessidade - mov_code.gs

---------------------------------------------------------------------------

- [ ] Trocar as cores do formulário - mov_form.html - V1.0
- [ ] Definir e corrigir os links do botão flutuante - mov_form.html - V1.0

---------------------------------------------------------------------------

- [ ] Corrigir corpo do e-mail - mail_template.html - V1.0
- [ ] Fazer assinatura do Ramo no e-mail - mail_template.html - V1.0
  - [x] Deixar uma pasta publica no Google Drive 
  - [x] Adicionar a imagem e o link para assinatura
  - [ ] Encontrar tutorial para montar assinatura
    ![Tutorial](https://drive.google.com/uc?id=0B8CcpExpMKFlX1hhTldISF9jU0E)

- [ ] Revisão ortográfica. - Código completo, documentação e comentários 

---------------------------------------------------------------------------

- [ ] Fazer documentação para modificação - Novo arquivo - 1.0
- [ ] Fazer documentação de uso - Novo arquivo - 1.0

---------------------------------------------------------------------------

### V2.X
- [ ] **Adicionar Transação de Transferência**
  - [ ] Ocultar campo de Upload
  - [ ] Trocar campo unidade por unidade de Saída
  - [ ] Adicionar campo unidade de entrada
  - [ ] Adicionar código para adicionar duas linhas: uma de entrada outra de saída
  - [ ] Adicionar no código o envio de e-mail para as duas unidades
  - [ ] Adicionar um aviso de que esta transação deve ser feita apenas por tesoureiros das respectivas unidades
- [ ] Separar e-mail dos capítulo e e-mail da pessoa que enviou - Novo arquivo
- [ ] Adicionar valor de caixa atualizado no e-mail dos capitulos - mov_code.gs
- [ ] **Adicionar o código de criação da planilha e do formulário de recibos**
  - [ ] Adicionar e nomear Pastas
  - [ ] Organizar campos
  - [ ] Adicionar fórmulas
  - [ ] Definir formatação das células
---------------------------------------------------------------------------

### Melhorias Programadas

- [ ] Separação do código em funções.
---------------------------------------------------------------------------

## Controle de versões

### **Alpha**

- [x] Definição dos campos do formulário e funcionamento dos mesmos.

### **Beta 1**
- [x] Troca do Theme para materializecss.
- [x] Melhorias Visuais.
- [x] Adicionar suporte a arquivos de câmera de celular.

### **Beta 2**

- [x] Renomear arquivo automaticamente.

```O Nome do arquivo CPmmT0000yy
CP - unidade em código numérico
mm - mês da movimentação
t - entrada ou saída
0000 - numero da mov
yy - ano da mov
```

### **Beta 3**

- [x] Implementar boas práticas no código - mov_code.gs.
- [x] Reorganização do algorítimo - mov_code.gs.

### **Beta 4**

- [x] Fazer laço if para cada capitulo - mov_code.gs.

### **Beta 5**

- [x] Melhorar visualização do valor financeiro nos textos.
- [x] Revisar e comentar código.
- [x] Limpeza e organização do código.

### **Beta 6**

- [x] Modificar o nome das variáveis no html e no código.

### **Beta 7**

- [x] Fazer corpo html em arquivo separado - mov_code.gs.
- [x] Criação do arquivo - mail_template.html.

### **Beta 8**

- [x] Eliminados dos scripts inutilizáveis - javaScript.html.
- [x] Atualizadas css do tema - mov_form.html.
- [x] Atualizadas js do tema - javaScript.html.
- [x] Adicionado Envio apenas se todas questões preenchidas.

### **Beta 9**

- [x] Responsividade no formulário.

### **Beta 10**

- [x] Hiperlinks nos arquivos de mail_template.html e Mov_thanks.html - Email e url.
- [x] Troca do botão type=submit para type=button - mov_form.html - Para evitar criação nova aba.
- [x] Adicionada mensagem de erro no campo email - mov_form.html.
- [x] Remoção de função desnecessária - JavaScript.html.
- [x] Modificação na permissão da URL - mov_code.gs - agora visualização.

### **Beta 11** 

- [x] 1 correção ortográfica - Código completo e comentários. (Obrigado Juh!)

### **Beta 12**

- [x] Limpeza dos códigos - Todos arquivos.
- [x] Comentar procedimentos - Todos arquivos .
- [x] Adicionar pattern no Campo E-mail - mov_form.html.
- [x] Adicionado possibilidade de caixas "modal".
- [x] Botão flutuante com menu vertical (antes era horizontal).
- [x] Corrigido bug do calendário (dropdown não aparecia).
- [x] Corrigido bug do campo valor (tamanho menor que restantes).
- [x] Adicionadas referências.