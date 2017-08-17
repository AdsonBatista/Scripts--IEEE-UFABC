# Recibos

Este é um Google Script voltado para emissão de recibos gerais do Ramo Estudantil IEEE UFABC. Ele faz parte de um processo de automatização financeira que está sendo desenvolvido.

Suas funções são:
- Produzir e enviar e o arquivo de recibo em duas copias. Uma para a pessoa e outra para o ramo.
  - Os arquivos são nomeados com um padrão
  - Os arquivos são depositados numa pasta do Google Drive.
  - Os dados são organizados em uma Planilha no próprio Google
- Enviar um e-mail para a pessoa efetuou o pagamento e um para o Ramo Estudantil como backup.

## Todo V-1:

---------------------------------------------------------------------------

*Versão atual: Beta 5*

---------------------------------------------------------------------------

- [ ] Adicionar links de pastas de cada capitulo - Verificar necessidade.
- [ ] Trocar as cores do formulário - Verificar necessidade. 

---------------------------------------------------------------------------
- [ ] Separar corpo do e-mail em um arquivo diferente.
- [ ] Melhorar texto corpo do e-mail.
- [ ] Fazer assinatura do Ramo no e-mail - mail_template.html - V1.0
  - [x] Deixar uma pasta publica no Google Drive.
  - [ ] Adicionar a imagem e o link para assinatura.
  - [ ] Encontrar tutorial para montar assinatura.
    ![Tutorial](https://drive.google.com/uc?id=0B8CcpExpMKFlX1hhTldISF9jU0E)
- [ ] Revisão ortográfica. - Código completo, documentação e comentários.

---------------------------------------------------------------------------

- [ ] Fazer documentação para modificação - Novo arquivo.
- [ ] Fazer documentação de uso - Novo arquivo.

---------------------------------------------------------------------------

### Melhorias Programadas

- [ ] Separação do código em funções.
---------------------------------------------------------------------------

## Controle de versões

### **Alpha**

- [x] Definição dos campos do formulário e funcionamento dos mesmos.

### **Beta 1**
- [x] Definições das Expressões regulares de verificação.
- [x] Criação do modelo de recibo.

### **Beta 2**

- [x] Definir padrão do nome do arquivo. (verificar código)

```O Nome do arquivo CPmmT0000yy
CP - unidade em código numérico
mm - mês da movimentação
t - entrada ou saída
0000 - numero da mov
yy - ano da mov
```

- [x] Melhoria na expressão regular do CPF. (Obrigado Junior Campos!)
- [x] Troca do padrão do recibo. (A4 para A6)
- [x] Melhorias no recibo.
  - [x] Formatação do CPF. (Obrigado Junior Campos!)
  - [x] Transcrição do valor do recibo.

### **Beta 3**

- [x] Criação do ID do recibo
- [x] Acrecentado do ID do recibo no recibo.

### **Beta 4**

- [x] Limpeza dos códigos.
- [x] Comentar procedimentos.

### **Beta 5**

- [x] Correção de Bugs
- [x] Melhorias no Recibo