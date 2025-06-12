module.exports = async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

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

    // Ordenar reservas por: data, depois hora de início, depois hora de fim
    reservas.sort((a, b) => {
      const aDate = new Date(a.date._seconds * 1000);
      const bDate = new Date(b.date._seconds * 1000);
      const aStart = new Date(a.inicialHour._seconds * 1000);
      const bStart = new Date(b.inicialHour._seconds * 1000);
      const aEnd = new Date(a.finalHour._seconds * 1000);
      const bEnd = new Date(b.finalHour._seconds * 1000);

      return aDate - bDate || aStart - bStart || aEnd - bEnd;
    });

    for (const r of reservas) {
      const dateObj = new Date(r.date._seconds * 1000);
      const start = new Date(r.inicialHour._seconds * 1000);
      const end = new Date(r.finalHour._seconds * 1000);
      const duracao = ((end - start) / (1000 * 60 * 60)).toFixed(2);

      const dateFormatada = `${String(dateObj.getDate()).padStart(
        2,
        "0"
      )}/${String(dateObj.getMonth() + 1).padStart(
        2,
        "0"
      )}/${dateObj.getFullYear()}`;

      sheet.addRow({
        ...r,
        date: dateFormatada,
        inicialHour: start.toTimeString().slice(0, 5),
        finalHour: end.toTimeString().slice(0, 5),
        duracao,
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const base64 = Buffer.from(buffer).toString("base64");

    res.status(200).json({ fileBase64: base64 });
  } catch (err) {
    console.error("Erro ao gerar Excel:", err);
    res.status(500).json({ error: "Erro interno ao gerar Excel" });
  }
};
