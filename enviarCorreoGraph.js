const axios = require('axios');
require('dotenv').config();

async function enviarCorreoGraph(destinatarios, pdfBase64, nombreArchivo) {
  try {
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      })
    );

    const accessToken = tokenResponse.data.access_token;
    const usuarioEmisor = 'jandrade@tactiqaec.com';

    // Verificación del base64
    console.log('Primeros caracteres de pdfBase64:', pdfBase64.slice(0, 50));
    let contentBytes = pdfBase64;
    if (pdfBase64.startsWith('data:application/pdf')) {
      contentBytes = pdfBase64.replace(/^data:application\/pdf;.*base64,/, '');
    }
    console.log('Primeros caracteres de contentBytes:', contentBytes.slice(0, 50));

    const emailData = {
      message: {
        subject: 'Análisis Crediticio PDF',
        body: {
          contentType: 'Text',
          content: 'Adjunto se encuentra el informe de análisis crediticio.'
        },
        toRecipients: destinatarios.map(email => ({
          emailAddress: { address: email }
        })),
        attachments: [
          {
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: nombreArchivo,
            contentBytes,
            contentType: 'application/pdf'
          }
        ]
      },
      saveToSentItems: 'false'
    };

    await axios.post(
      `https://graph.microsoft.com/v1.0/users/${usuarioEmisor}/sendMail`,
      emailData,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      }
    );

    console.log('Correo enviado exitosamente.');
  } catch (error) {
    console.error('Error al enviar correo:', error.response?.data || error);
  }
}

module.exports = enviarCorreoGraph;
