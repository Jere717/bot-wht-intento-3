// --- 1. IMPORTAR HERRAMIENTAS ---
const express = require('express');
const cors = require('cors');
const qrcode = require('qrcode');
const { Client, LocalAuth } = require('whatsapp-web.js');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const MistralClient = require('@mistralai/mistralai');

// --- 2. CONFIGURACIÓN INICIAL ---
const app = express();
app.use(cors());
app.use(express.json());

let qrData = null;
let isConnected = false;

const client = new Client({
  authStrategy: new LocalAuth({ dataPath: './session' }),
  puppeteer: {
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  }
});

client.on('qr', async (qr) => {
  qrData = await qrcode.toDataURL(qr);
  isConnected = false;
  console.log('QR generado y listo para escanear.');
});

client.on('ready', () => {
  console.log('✅ ¡Cliente de WhatsApp conectado y listo!');
  qrData = null;
  isConnected = true;
});

client.initialize();

const ID_DE_TU_GOOGLE_SHEET = process.env.GOOGLE_SHEET_ID || '1jjROnAY1TobjiDYwjHv8YD2i-D_LlwTtw79XZ1i-1Oo';
const creds = process.env.GOOGLE_CREDENTIALS ? JSON.parse(process.env.GOOGLE_CREDENTIALS) : null;

// --- 3. ENDPOINT PARA GOOGLE APPS SCRIPT ---
app.post('/getqr', (req, res) => {
  if (qrData) {
    res.json({ qr: qrData, status: 'Esperando escaneo' });
  } else if (isConnected) {
    res.json({ status: 'CONECTADO' });
  } else {
    res.json({ status: 'NO CONECTADO' });
  }
});

// --- ENDPOINT PARA PROCESAR MENSAJES DESDE APPS SCRIPT ---
app.post('/registermessage', async (req, res) => {
  try {
    const { mensaje, numero } = req.body;
    if (!mensaje || !numero) {
      return res.status(400).json({ error: 'Faltan parámetros: mensaje y numero' });
    }
    // --- Conexión con Google Sheets ---
    const doc = new GoogleSpreadsheet(ID_DE_TU_GOOGLE_SHEET, {
      apiKey: null,
      access_token: null,
      service_account_auth: creds
    });
    await doc.loadInfo();
    const configSheet = doc.sheetsByTitle['Configuracion'];
    await configSheet.loadCells('B9:D8');
    const mistralApiKey = configSheet.getCellByA1('D8').value;
    const botPrompt = configSheet.getCellByA1('B9').value;
    if (!mistralApiKey || !botPrompt) {
      return res.status(500).json({ error: 'No se encontró la API Key de Mistral o el Prompt en la hoja de configuración.' });
    }
    // --- Conexión con Mistral AI ---
    const mistralClient = new MistralClient(mistralApiKey);
    const chatResponse = await mistralClient.chat({
      model: 'mistral-large-latest',
      messages: [{ role: 'user', content: `${botPrompt}\n\nCliente: ${mensaje}` }],
    });
    const aiResponse = chatResponse.choices[0].message.content;
    // --- Guardar en la hoja de 'Solicitudes' ---
    const requestsSheet = doc.sheetsByTitle['Solicitudes'];
    await requestsSheet.addRow({
      Fecha: new Date().toLocaleString(),
      Usuario: numero,
      Mensaje: mensaje,
      Respuesta_IA: aiResponse
    });
    // --- Responder al Apps Script ---
    res.json({ respuesta: aiResponse });
  } catch (error) {
    console.error('Error procesando mensaje:', error);
    res.status(500).json({ error: 'Error procesando el mensaje.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor escuchando en puerto ${PORT}`));
