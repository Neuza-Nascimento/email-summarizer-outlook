Office.onReady(() => {
    document.getElementById("summarize-button").onclick = handleSummarize;
    document.getElementById("button-suggestion").onclick= gerarResposta;
    document.getElementById("button-new-suggestion").onclick =gerarResposta;
   
  });

  async function handleSummarize() {
    
    const summaryDiv = document.getElementById("summary");
    const gerarResposta = document.getElementById("gerar-resposta");
    const summarizeButton = document.getElementById("summarize-button");

    summaryDiv.textContent = "Lendo o email...";
   
  
    try {
      Office.context.mailbox.item.body.getAsync("text", async (result) => {
        let emailText;
    
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          emailText = result.value; 
     
          const itemId = Office.context.mailbox.item.itemId;

        
          const cacheKey = `resumo_${itemId}`;
          const resumoArmazenado = localStorage.getItem(cacheKey);

          if (resumoArmazenado) {
            summaryDiv.textContent = "Resumo (armazenado): " + resumoArmazenado;
          
          } else {
    
      
        summaryDiv.textContent = "Resumindo com IA...";
    
        const res = await fetch("https://localhost:3000/resumir", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text: emailText })
        });
    
        const data = await res.json();
        const resumo = data.resumo; 
        localStorage.setItem(cacheKey, resumo);

        summaryDiv.textContent = "Resumo: " + resumo;
        summarizeButton.style,display = "none";

      };

        const remetente = Office.context.mailbox.item.from.emailAddress;
       
      if (remetente) {
        
        if (remetente.toLowerCase().includes("no-reply") || remetente.toLowerCase().includes("noreply")) {
          console.log("Este é um email no-reply. Ignorar."); 

        } else {

              gerarResposta.style.display = "block";
     
        }

      } else {
        console.log("Item de email não disponível.");
      }
    }});

    } catch (err) {
      summaryDiv.textContent = "Erro ao resumir: " + err.message;
    }
  }    
  
 async function gerarResposta() {
   
    const loading= document.getElementById("loading-reply");
    const replyBox = document.getElementById("reply-box");
    const replyText = document.getElementById("reply-text");
    const replyActions = document.getElementById("reply-actions");
    const styleSelect = document.getElementById('reply-style-select');
    const selectedStyle = styleSelect ? styleSelect.value : 'À maneira da IA'
    
    const remetente = Office.context.mailbox.item.from.emailAddress;
     const assunto = Office.context.mailbox.item;

    loading.style.display= "block";
    
  try {
      Office.context.mailbox.item.body.getAsync("text", async (result) => {
        let emailText;
    
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          emailText = result.value;
    
        const res = await fetch("https://localhost:3000/gerar-resposta", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text: emailText, style: selectedStyle })
        });
    
        const data = await res.json();
        replyText.textContent = "Resposta: " +  data.resumo
        
        loading.style.display= "none";

        if (replyText.textContent) {
          replyBox.style.display = "block";
        }

        if (replyBox.style.display === "block") {
          replyActions.style.display = "block";
        }
      
         document.getElementById("button-open-reply").onclick = function enviarEmail() {
          
          const respostaFormatada = data.resumo.replace(/\n/g, "<br>");

            

             /* Office.context.mailbox.displayNewMessageForm({
              toRecipients: [remetente],
              subject: "Resposta: " + assunto.subject,
              htmlBody: respostaFormatada,
            });*/

              Office.context.mailbox.item.displayReplyForm({
              htmlBody: respostaFormatada
            });

              
          replyBox.style.display= "none";

          const summaryDiv = document.getElementById("summary");
          const gerarResposta = document.getElementById("gerar-resposta");
          const geraResumo =  document.getElementById("summarize-button");
   
          summaryDiv.textContent = "Enviado com sucesso!";
          gerarResposta.style.display = "none";
          geraResumo.style.display ="none";

  }
    }});
  
    } catch (err) {
      replyText.textContent = "Erro ao gerar resposta: " + err.message;

       if (replyText.textContent) {
          replyBox.style.display = "block";
        }
    }
  }

 