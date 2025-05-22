const express = require("express");
const fetch = require("node-fetch");
const cors = require("cors");
const dotenv = require("dotenv");
const path = require("path");
const https = require("https");
const fs = require("fs");

dotenv.config();
const app = express();
const publicDir = path.resolve(__dirname, "../public");


app.use(express.static(publicDir));
app.use(cors());
app.use(express.json())


app.post("/resumir", async (req, res) => {
  console.log("Requisição recebida para resumir e-mail.");
  const { text } = req.body;

  if (!text) {
    console.error("Texto do e-mail não fornecido.");
    return res.status(400).json({ erro: "Texto do e-mail não fornecido" });
  }

  try {
    const resposta = await fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + process.env.GEMINI_API_KEY,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          contents: [
            {
              parts: [
                {
                  text: `Resuma o seguinte e-mail em menos de 50 palavras de forma clara e objetiva:\n\n${text}`
                }
              ]
            }
          ]
        })
      }
    );

    const data = await resposta.json();
    const resumo = data.candidates?.[0]?.content?.parts?.[0]?.text;

    if (!resumo) {
      console.error("Resumo não encontrado na resposta da API:", data);
      return res.status(500).json({ erro: "Resumo não encontrado na resposta da API" });
    }

    res.json({ resumo });
  } catch (err) {
    console.error("Erro ao resumir com Gemini:", err);
    res.status(500).json({ erro: "Falha ao gerar resumo com Gemini" });
  }
});

app.post("/gerar-resposta", async (req, res) => {
  console.log("Requisição recebida para dar resposta ao e-mail.");
  const { text, style } = req.body;

  let promptStyle = "";
  if (style === "casual") {
    promptStyle = "Responda de maneira casual e amigável.";
  } else if (style === "criativo") {
    promptStyle = "Responda de maneira divertida e criativa, como uma IA.";
  } else {
    promptStyle = "Responda de maneira formal e profissional.";
  }

  try {
    const resposta = await fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + process.env.GEMINI_API_KEY2,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          contents: [
            {
              parts: [
                {
                  text: `${promptStyle}\n\nEmail recebido:\n${text}\n\n Não coloque assinatura nem despeça no fim.`
                }
              ]
            }
          ]
        })
      }
    );

    const data = await resposta.json();
    const resumo = data.candidates?.[0]?.content?.parts?.[0]?.text;

    if (!resumo) {
      console.error("Resumo não encontrado na resposta da API:", data);
      return res.status(500).json({ erro: "Resumo não encontrado na resposta da API" });
    }

    res.json({ resumo });
  } catch (err) {
    console.error("Erro ao resumir com Gemini:", err);
    res.status(500).json({ erro: "Falha ao gerar resumo com Gemini" });
  }
});


const httpsOptions = {
  key: fs.readFileSync("localhost-key.pem"),
  cert: fs.readFileSync("localhost.pem"),
};

https.createServer(httpsOptions, app).listen(3000, () => {
  console.log("✅ Servidor rodando em https://localhost:3000")
});
