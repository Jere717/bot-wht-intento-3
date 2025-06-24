import express from 'express';
import cors from 'cors';
import qrcode from 'qrcode';
import { Client, LocalAuth } from 'whatsapp-web.js';
import { MistralClient } from '@mistralai/mistralai';
import dotenv from 'dotenv';
dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

const sessions = {};
const mistral = new MistralClient({ apiKey: process.env.MISTRAL_API_KEY });

function createSession(sessionId) {
  if (sessions[sessionId]) return sessions[sessionId];

  const client = new Client({
    authStrategy: new LocalAuth({ dataPath: `./sessions/${sessionId}` }),
    puppeteer: {
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    }
  });

  sessions[sessionId] = {
    client,
    qr: null,
    connected: false
  };

  client.on('qr', async (qr) => {
    const base64 = await qrcode.toDataURL(qr);
    sessions[sessionId].qr = base64;
    sessions[sessionId].connected = false;
    console.log(`🔐 QR generado para sesión ${sessionId}`);
  });

  client.on('ready', () => {
    sessions[sessionId].qr = null;
    sessions[sessionId].connected = true;
    console.log(`✅ Sesión ${sessionId} conectada`);
  });

  client.on('message', async msg => {
    const text = msg.body?.trim();
    if (!text) return;
    const reply = await handleIA(text);
    msg.reply(reply);
  });

  client.initialize();
}

async function handleIA(texto) {
  try {
    const chat = await mistral.chat({
      model: 'mistral-small',
      messages: [{ role: 'user', content: texto }]
    });
    return chat.choices?.[0]?.message?.content || "🤖 No entendí, intentá de nuevo.";
  } catch (e) {
    console.error("Mistral error:", e.message);
    return "⚠️ Error al usar la IA.";
  }
}

// Endpoint de QR (usado por Apps Script)
app.post('/getqr', (req, res) => {
  const sessionId = req.body.sheet_id;
  if (!sessionId) {
    return res.status(400).json({ status: "-1", message: "Falta el ID del spreadsheet (sheet_id)." });
  }

  createSession(sessionId);

  const s = sessions[sessionId];
  if (s.connected) {
    return res.json({ status: "0", message: "CONECTADO" });
  } else if (s.qr) {
    return res.json({ status: "0", message: "Esperando escaneo", qr: encodeURIComponent(s.qr) });
  } else {
    return res.json({ status: "-1", message: "Inicializando sesión. Intenta nuevamente en unos segundos." });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Servidor activo en puerto ${PORT}`));