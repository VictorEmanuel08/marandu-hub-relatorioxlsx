module.exports = async (req, res) => {
  // Configura os headers CORS sempre no topo
  //   res.setHeader("Access-Control-Allow-Origin", "https://marandu-hub.flutterflow.app");
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  // Responde a requisições OPTIONS (pré-flight)
  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  try {
    const ExcelJS = require("exceljs");
    const reservas = req.body;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Reservas");

    sheet.columns = [
      { header: "ID Reserva", key: "uid" },
      { header: "ID Usuário", key: "idUser" },
      { header: "Sala", key: "idRoom" },
      { header: "Data", key: "date" },
      { header: "Hora Início", key: "inicialHour" },
      { header: "Hora Fim", key: "finalHour" },
      { header: "Duração (h)", key: "duracao" },
      { header: "Convidados", key: "numberGuests" },
      { header: "Status", key: "status" },
      { header: "Descrição", key: "desc" },
    ];

    for (const r of reservas) {
      const start = new Date(r.inicialHour._seconds * 1000);
      const end = new Date(r.finalHour._seconds * 1000);
      const duracao = ((end - start) / (1000 * 60 * 60)).toFixed(2);
      sheet.addRow({
        ...r,
        date: new Date(r.date._seconds * 1000).toISOString().slice(0, 10),
        inicialHour: start.toTimeString().slice(0, 5),
        finalHour: end.toTimeString().slice(0, 5),
        duracao,
      });
    }

    // Gera o arquivo em memória
    const buffer = await workbook.xlsx.writeBuffer();

    // Retorna como base64
    const base64 = Buffer.from(buffer).toString("base64");

    res.status(200).json({ fileBase64: base64 });
  } catch (err) {
    console.error("Erro ao gerar Excel:", err);
    res.status(500).json({ error: "Erro interno ao gerar Excel" });
  }
};
