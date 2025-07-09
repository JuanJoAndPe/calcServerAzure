const express = require('express');
const cors = require('cors');
const axios = require('axios');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const mongoose = require('mongoose');
const enviarCorreoGraph = require('./enviarCorreoGraph');
const qs = require('qs');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

// Usuarios
const users = [
  {
    id: 1,
    username: process.env.ADMIN_USER,
    passwordHash: bcrypt.hashSync(process.env.ADMIN_PASSWORD, 10),
    role: 'admin'
  },
  {
    id: 2,
    username: process.env.ANALYST_USER,
    passwordHash: bcrypt.hashSync(process.env.ANALYST_PASSWORD, 10),
    role: 'analyst'
  }
];

// Middleware JWT
const authenticateJWT = (req, res, next) => {
  const authHeader = req.headers.authorization;

  if (authHeader) {
    const token = authHeader.split(' ')[1];
    jwt.verify(token, process.env.JWT_SECRET, (err, user) => {
      if (err) return res.sendStatus(403);
      req.user = user;
      next();
    });
  } else {
    res.sendStatus(401);
  }
};

// Login
app.post('/login', (req, res) => {
  const { username, password } = req.body;
  const user = users.find(u => u.username === username);
  if (!user || !bcrypt.compareSync(password, user.passwordHash)) {
    return res.status(401).json({ error: 'Credenciales inválidas' });
  }

  const token = jwt.sign(
    { userId: user.id, username: user.username, role: user.role },
    process.env.JWT_SECRET,
    { expiresIn: '24h' }
  );
  res.json({ token });
});

// Proxy autenticado
app.post('/proxy', authenticateJWT, async (req, res) => {
  try {
    const response = await axios.post(process.env.AVAL_URL, req.body, {
      headers: {
        'Authorization': 'Basic ' + Buffer.from('WS-TAQTICA:&jg4I(iKGA').toString('base64'),
        'Content-Type': 'application/json',
        'User-Agent': 'Mozilla/5.0'
      }
    });
    res.json(response.data);
  } catch (error) {
    console.error('Error en /proxy:', error.response?.data || error.message);
    res.status(500).json({ error: error.message });
  }
});

// MongoDB
mongoose.connect(process.env.MONGO_URI)
  .then(() => console.log('Conectado a MongoDB'))
  .catch(err => console.error('Error al conectar a MongoDB:', err));

const AnalisisSchema = new mongoose.Schema({
  cedulaDeudor: String,
  nombreDeudor: String,
  scoreDeudor: Number,
  evaluacionIntegralDeudor: String,
  cedulaConyuge: String,
  nombreConyuge: String,
  scoreConyuge: Number,
  evaluacionIntegralConyuge: String,
  patrimonio: Number,
  ingresos: Number,
  gastos: Number,
  marca: String,
  modelo: String,
  valorVehiculo: Number,
  entrada: Number,
  gtosLegales: Number,
  dispositivo: Number,
  seguroDesgravamen: Number,
  seguroVehicular: Number,
  montoFinanciar: String,
  cuotaMensual: String,
  plazo: Number,
  indicadorEndeudamiento: String,
  decisionFinal: String,
  fecha: Date
});

const Analisis = mongoose.model('Analisis', AnalisisSchema);

app.post('/guardarAnalisis', async (req, res) => {
  try {
    const nuevoAnalisis = new Analisis(req.body);
    await nuevoAnalisis.save();
    res.json({ mensaje: 'Análisis guardado correctamente' });
  } catch (err) {
    console.error('Error al guardar en MongoDB:', err);
    res.status(500).json({ error: 'Error al guardar análisis en MongoDB' });
  }
});

// Enviar correo con Graph
app.post('/enviarCorreo', async (req, res) => {
  const { pdfBase64, nombreArchivo, destinatarios } = req.body;
  try {
    await enviarCorreoGraph(destinatarios, pdfBase64, nombreArchivo);
    res.json({ mensaje: 'Correo enviado con Microsoft Graph' });
  } catch (err) {
    console.error('Error al enviar correo:', err);
    res.status(500).json({ error: 'Error al enviar con Graph' });
  }
});

// Enviar a Excel - Versión actualizada para manejar tablas existentes
app.post('/guardarExcel', authenticateJWT, async (req, res) => {
  try {
    const username = req.user.username;
    const tokenRes = await axios.post(
      `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
      qs.stringify({
        grant_type: 'client_credentials',
        client_id: process.env.CLIENT_ID2,
        client_secret: process.env.CLIENT_SECRET2,
        scope: 'https://graph.microsoft.com/.default',
      }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );
    const accessToken = tokenRes.data.access_token;

    const userPrincipal = 'pmantilla@tactiqaec.com';
    const filePath = 'ORIGINACION/ANÁLISIS DE CRÉDITOS/BASE_AUTOMATICA.xlsx';

    const file = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${userPrincipal}/drive/root:/${filePath}:`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const fileId = file.data.id;

    const sheetRes = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${userPrincipal}/drive/items/${fileId}/workbook/worksheets`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const sheetName = sheetRes.data.value[0].name;

    const tablesRes = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${userPrincipal}/drive/items/${fileId}/workbook/worksheets('${sheetName}')/tables`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    let tableId;
    if (tablesRes.data.value.length > 0) {
      tableId = tablesRes.data.value[0].id;
    } else {
      const tableRes = await axios.post(
        `https://graph.microsoft.com/v1.0/users/${userPrincipal}/drive/items/${fileId}/workbook/worksheets('${sheetName}')/tables/add`,
        { address: 'A1:V1', hasHeaders: true },
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
      tableId = tableRes.data.id;

      await axios.post(
        `https://graph.microsoft.com/v1.0/users/${userPrincipal}/drive/items/${fileId}/workbook/tables/${tableId}/rows/add`,
        {
          values: [["FECHA", "CEDULA", "APELLIDOS_NOMBRES", "CEDULA_CYG", "APELLIDOS_NOMBRES_CYG", "CONCESIONARIO", "LOCAL", "ASESOR", "MARCA", "MODELO", "VALOR", "ENTRADA", "PORCENTAJE", "SEG_DESGRAVAMEN", "SEG_VEHICULAR", "FIDEICOMISO", "DISPOSITIVO", "MONTO_FINANCIAR", "PLAZO", "SCORE", "SCORE_CYG", "DECISION"]]
        },
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
    }

    await axios.post(
      `https://graph.microsoft.com/v1.0/users/${userPrincipal}/drive/items/${fileId}/workbook/tables/${tableId}/rows/add`,
      {
        values: [[
          req.body.fecha,
          req.body.cedula,
          req.body.nombre,
          req.body.cedula_cyg,
          req.body.conyuge,
          req.body.concesionario,
          req.body.local,
          username,
          req.body.marca,
          req.body.modelo,
          req.body.valor,
          req.body.entrada,
          req.body.porcentaje,
          req.body.seg_desgravamen,
          req.body.seg_vehicular,
          req.body.fideicomiso,
          req.body.dispositivo,
          req.body.monto_financiar,
          req.body.plazo,
          req.body.score,
          req.body.score_cyg,
          req.body.decision,
        ]]
      },
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    res.json({ mensaje: 'Datos guardados correctamente en Excel' });

  } catch (err) {
    console.error('Error al guardar en Excel:', err.response?.data || err.message);
    res.status(500).json({ error: 'No se pudo guardar en Excel' });
  }
});

app.listen(3000, () => console.log('Servidor corriendo en puerto 3000'));