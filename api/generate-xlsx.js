import ExcelJS from "exceljs";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Método não permitido" });
  }

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

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader(
    "Content-Disposition",
    'attachment; filename="relatorio_reservas.xlsx"'
  );

  await workbook.xlsx.write(res);
  res.end();
}
