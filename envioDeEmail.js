
/**
 * Variaveis de scopo global para acesso as guias do Sheets
 */
 var planilha = SpreadsheetApp.getActiveSpreadsheet();
 var guiaapontamento = planilha.getSheetByName("apontamento");
 var guiadados = planilha.getSheetByName("dados");
 var guiaemail = planilha.getSheetByName("lista_email");
 
 /**
  * Enviar e-mail a partir da guia dados
  *  @return {void}
  */
 function enviaEmail() {
   let horario = Browser.inputBox("Insira o corpo do e-mail aqui 😉")
   filtroEmail()
   let ultimaLinhaPreenchida = pegarObjetoUltimalinha()
   for (let i=8; i <= ultimaLinhaPreenchida['guiaDadosUltimaLinha']; i++){
     let conteudoEmail = bodyEmail()
     let email = deleteEnviado()
     let mensagem = {
       to: email,
       subject: "Correção de Processos ⏰",
       htmlBody: horario + conteudoEmail,
       name: "Apontamento",
     }
     MailApp.sendEmail(mensagem);
   }
 }
 
 /**
  * Formula para buscar os emails da guiaemail
  * @return{void}
  */
 function filtroEmail() {
   let ultimaLinha = pegarObjetoUltimalinha()
   for (let i = 8; i <= ultimaLinha['guiaDadosUltimaLinha']; i++) {
     guiadados.getRange("Q" + i).setFormula('=IFERROR(VLOOKUP(C:C;lista_email!A2:C400;2;FALSE);"")');
   }
 }
 /**
  * Deletar emails que ja foram enviados
  * @param{String} emailDeletar
  */
 function deleteEnviado(emailDeletar = guiadados.getRange("Q8").getValue()) {
   let ultimaLinhaPreenchida = pegarObjetoUltimalinha()
   for (let i = 8; i <= ultimaLinhaPreenchida['guiaDadosUltimaLinha']; i++) {
     if (guiadados.getRange("Q" + i).getValue() === emailDeletar) {
       guiadados.deleteRows(i)
       i = i - 1
     }
   }
   console.log(emailDeletar)
   return emailDeletar
 }
 
 /**
  *  Define as informações do corpo do email
  */
 function bodyEmail() {
   let ultimaLinhaPreenchida = pegarObjetoUltimalinha()
 
   let body = Array()
   const linhaSelect = 1
   const colunaSelect = 16
   const colunaInicial = 1
   let primeiroEmailListaDados = guiadados.getRange("Q8").getValue()
   for (let i = 8; i <= ultimaLinhaPreenchida['guiaDadosUltimaLinha']; i++) {
     if (guiadados.getRange("Q" + i).getValue() === primeiroEmailListaDados) {
       body.push(guiadados.getSheetValues(i, colunaInicial, linhaSelect, colunaSelect))
     }
   }
   body = body.map((item) => {
     console.log(item)
     return item[0]
   })
   let htmlBody = `<!DOCTYPE html>
   <html lang="en">
 
   <head>
     <meta charset="UTF-8">
     <meta http-equiv="X-UA-Compatible" content="IE=edge">
     <meta name="viewport" content="widtd=device-widtd, initial-scale=1.0">
     <title>Document</title>
     <style>
         @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700;900&display=swap');
         * {
             padding: 0;
             margin: 0;
             font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;
         }
     </style>
     </head>
 
   <body style="background-color: rgb(182, 182, 182);  font-size: 1.1em;">
     <div style="overflow: auto; width: 90%; background-color: #fff; margin: 0 auto;">
         <table style="font-size: .6559em; border: 1px solid rgba(0,0,0,.1)">
             <thead style="background-color: rgb(0, 51, 77); color: #fff">
                 <tr>
                     <td style="padding:10px 20px; border:1px solid;">ETIQUETA</td>
                     <td style="padding:10px 20px;">LÍDER</td>
                     <td style="padding:10px 20px;">LOGIN</td>
                     <td style="padding:10px 20px;">USUARIO</td>
                     <td style="padding:10px 20px;">OBJETO DA AÇÃO</td>
                     <td style="padding:10px 20px;">DIV. ASSUNTO X OBJETO DA AÇÃO</td>
                     <td style="padding:10px 20px;">DIV. RESULTADO DA AÇÃO X OBJETO</td>
                     <td style="padding:10px 20px;">DATAS</td>
                     <td style="padding:10px 20px;">CLASSE X COMPETENCIA</td>
                     <td style="padding:10px 20px;">POLO ATIVO</td>
                     <td style="padding:10px 20px;">POLO PASSIVO</td>
                     <td style="padding:10px 20px;">ADVOGADO POLO ATIVO</td>
                     <td style="padding:10px 20px;">ADVOGADO POLO PASSIVO</td>
                     <td style="padding:10px 20px;">CAMPOS OBRIGATORIOS</td>
                     <td style="padding:10px 20px;">CLASSE X COMPETENCIA</td>
                     <td style="padding:10px 20px;">RESULTADO DA AÇÃO X OBJETO</td>
                 </tr>
             </thead>
             <tbody>`
 
   body.forEach(function (item, key) {
     htmlBody += "<tr>"
     item.forEach(function (colunaArray) {
       htmlBody += `<td style="border:1px solid rgba(0,0,0,.1)">${colunaArray}</td>`
     })
     htmlBody += "</tr>"
   })
   htmlBody += "</tbody></table></div></body></html>"
   console.log(htmlBody)
   return htmlBody
 }
 
 /**
  * Identificar ultima linha preenchida das guias email e dados.
  * @return {objeto}
  */
 function pegarObjetoUltimalinha() {
   let objetoultimalinha = {
     'guiaEmailUltimaLinha': guiaemail.getLastRow() - 1, // "-1 para nao incluir a linha de cabeçalho"
     'guiaDadosUltimaLinha': guiadados.getLastRow()
   }
   return objetoultimalinha
 };
 
 /**
  * Ocultar e mostrar guias
  * Ocultar para as outras guias não sairem no arquivo PDF do email.
  *  @param{boolean} exibir
  */
 function exibirOcultarGuias(exibir = true) {
   if (exibir) {
     planilha.getSheetByName("lista_email").activate();
     planilha.getSheetByName("dados").activate();
   } else {
     planilha.getSheetByName("lista_email").hideSheet();
     planilha.getSheetByName("dados").hideSheet();
   }
 }