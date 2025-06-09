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


  
  Office.context.mailbox.item.body.getAsync("text", async (result) => {

    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      summaryDiv.textContent = "Erro ao ler o email.";
      return;
    }

    const emailText = result.value;
    const itemId = Office.context.mailbox.item.itemId;
    const remetente = Office.context.mailbox.item.from?.emailAddress || "";
    const cacheKey = `resumo_${itemId}`;

    try {

      const resumoArmazenado = await buscarNoIndexedDB('ResumoDB', 'resumos', cacheKey);

      if (resumoArmazenado) {
        summaryDiv.textContent = "Resumo (armazenado): " + resumoArmazenado;
      } else {
        summaryDiv.textContent = "Resumindo com IA...";

        const res = await fetch("https://localhost:3000/resumir", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text: emailText }),
        });

        const data = await res.json();
        const resumo = data.resumo;

       await salvarNoIndexedDB('ResumoDB', 'resumos', cacheKey, resumo);

      

        summaryDiv.textContent = "Resumo: " + resumo;
        summarizeButton.style.display = "none";
      }
    } catch (error) {
      console.error("Erro ao acessar IndexedDB ou resumir:", error);
      summaryDiv.textContent = "Erro ao resumir o email.";
    }

    if (remetente) {
      if (
        remetente.toLowerCase().includes("no-reply") ||
        remetente.toLowerCase().includes("noreply")
      ) {
        console.log("Este é um email no-reply. Ignorar.");
      } else {
        gerarResposta.style.display = "block";
      }
    } else {
      console.log("Item de email não disponível.");
      gerarResposta.style.display = "none";
    }
  });
}

  
async function gerarResposta() {
  const loading = document.getElementById("loading-reply");
  const replyBox = document.getElementById("reply-box");
  const replyText = document.getElementById("reply-text");
  const replyActions = document.getElementById("reply-actions");
  const styleSelect = document.getElementById('reply-style-select');
  const selectedStyle = styleSelect ? styleSelect.value : 'À maneira da IA';

  const remetente = Office.context.mailbox.item.from.emailAddress;
  const assunto = Office.context.mailbox.item;
  const itemId = Office.context.mailbox.item.itemId;
  const cacheKey = `resposta_${itemId}_${selectedStyle}`;

  loading.style.display = "block";

  Office.context.mailbox.item.body.getAsync("text", async (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      replyText.textContent = "Erro ao ler o email.";
      loading.style.display = "none";
      return;
    }

    const emailText = result.value;

    try {
     
      const respostaArmazenada =  await buscarNoIndexedDB('RespostaDB', 'respostas', cacheKey);


      let respostaFinal;

      if (respostaArmazenada) {
        respostaFinal = "Resposta (armazenada):" + respostaArmazenada;
      } else {
        const res = await fetch("https://localhost:3000/gerar-resposta", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text: emailText, style: selectedStyle })
        });

        const data = await res.json();
        respostaFinal = data.resumo;

        await salvarNoIndexedDB('RespostaDB', 'respostas', cacheKey, respostaFinal);

      }

      replyText.textContent = "Resposta: " + respostaFinal;
      loading.style.display = "none";

      if (replyText.textContent) replyBox.style.display = "block";
      if (replyBox.style.display === "block") replyActions.style.display = "block";

      document.getElementById("button-open-reply").onclick = function enviarEmail() {
        const respostaFormatada = respostaFinal.replace(/\n/g, "<br>");

        Office.context.mailbox.displayNewMessageForm({
          toRecipients: [remetente],
          subject: "Resposta: " + assunto.subject,
          htmlBody: respostaFormatada,
        });

        replyBox.style.display = "none";

        const summaryDiv = document.getElementById("summary");
        const gerarResposta = document.getElementById("gerar-resposta");
        const geraResumo = document.getElementById("summarize-button");

        summaryDiv.textContent = "Enviado com sucesso!";
        gerarResposta.style.display = "none";
        geraResumo.style.display = "none";
      };

    } catch (err) {
      replyText.textContent = "Erro ao gerar resposta: " + err.message;
      replyBox.style.display = "block";
      loading.style.display = "none";
    }
  });
}



function abrirBanco(nomeDB, storeName) {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(nomeDB, 1);

    request.onupgradeneeded = (event) => {
      const db = event.target.result;
      if (!db.objectStoreNames.contains(storeName)) {
        db.createObjectStore(storeName, { keyPath: 'key' });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(`Erro ao abrir o banco ${nomeDB}`);
  });
}

async function salvarNoIndexedDB(nomeDB, storeName, cacheKey, valor) {
  const db = await abrirBanco(nomeDB, storeName);
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readwrite');
    const store = tx.objectStore(storeName);
    const request = store.put({ key: cacheKey, value: valor });

    request.onsuccess = () => {
      db.close();
      resolve();
    };
    request.onerror = () => {
      db.close();
      reject(`Erro ao salvar em ${storeName}`);
    };
  });
}

async function buscarNoIndexedDB(nomeDB, storeName, cacheKey) {
  const db = await abrirBanco(nomeDB, storeName);
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readonly');
    const store = tx.objectStore(storeName);
    const request = store.get(cacheKey);

    request.onsuccess = () => {
      db.close();
      if (request.result) resolve(request.result.value);
      else resolve(null);
    };
    request.onerror = () => {
      db.close();
      reject(`Erro ao buscar em ${storeName}`);
    };
  });
}




 