import express from 'express';
import cors from 'cors';
import qrcode from 'qrcode';
import pkg from 'whatsapp-web.js';
const { Client, LocalAuth } = pkg;
import MistralClient from '@mistralai/mistralai';
import dotenv from 'dotenv';
import fs from 'fs';
dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

const sessions = {};
const mistral = new MistralClient({ apiKey: process.env.MISTRAL_API_KEY });

function createSession(sessionId) {
  // Limpieza de sesión fallida antes de crearla
  const sessionPath = `./sessions/${sessionId}`;
  if (fs.existsSync(sessionPath)) {
    fs.rmSync(sessionPath, { recursive: true });
    console.log(`🧹 Sesión eliminada para ${sessionId}`);
  }

  const client = new Client({
    authStrategy: new LocalAuth({ dataPath: `./sessions/${sessionId}` }),
      puppeteer: {
        headless: true,
        args: [
          '--no-sandbox',
          '--disable-setuid-sandbox',
          '--disable-dev-shm-usage',
          '--disable-accelerated-2d-canvas',
          '--disable-gpu',
          '--window-size=1920x1080'
        ]
      }
  });

  sessions[sessionId] = {
    client,
    qr: null,
    connected: false
  };

  // En el evento 'qr' (dentro de createSession)
  client.on('qr', async (qr) => {
    const base64 = await qrcode.toDataURL(qr);
    sessions[sessionId].qr = base64;
    sessions[sessionId].connected = false;
    
    // Notificar a GAS vía POST
    const payload = {
      op: 'qr',
      qr: base64
    };
    
    fetch(sessions[sessionId].appScript, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).catch(err => console.error('Error notificando a GAS:', err));
  });

  // En el evento 'ready' (dentro de createSession)
  client.on('ready', () => {
    sessions[sessionId].qr = null;
    sessions[sessionId].connected = true;

    // Notificar a GAS que la sesión está lista
    const payload = {
      op: 'qr',
      qr: 'CONECTADO',
      session: sessionId,
      numero: client.info.wid.user
    };
    
    fetch(sessions[sessionId].appScript, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).catch(err => console.error('Error notificando a GAS (ready):', err));
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

// Endpoint para verificar el estado de las sesiones
app.get('/status', (req, res) => {
  const count = Object.keys(sessions).length;
  res.json({ status: 'ok', sesiones_activas: count });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Servidor activo en puerto ${PORT}`));