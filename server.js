const express = require("express");
const ExcelJS = require("exceljs");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.json());

app.post("/guardar", async (req, res) => {
    const { fecha, nombre, apellido, correo, cantidad, pelicula } = req.body;

    const workbook = new ExcelJS.Workbook();
    const file = "compras.xlsx";

    // Si el archivo ya existe, lo abrimos. Si no, creamos uno nuevo.
    try {
        await workbook.xlsx.readFile(file);
    } catch {
        workbook.addWorksheet("Compras");
        const sheet = workbook.getWorksheet("Compras");
        sheet.addRow(["Fecha", "Nombre", "Apellido", "Correo", "Cantidad", "Película"]);
    }

    const sheet = workbook.getWorksheet("Compras");
    sheet.addRow([fecha, nombre, apellido, correo, cantidad, pelicula]);

    await workbook.xlsx.writeFile(file);

    res.json({ mensaje: "Compra guardada correctamente en Excel" });
});

app.listen(3000, () => console.log("Servidor corriendo en http://localhost:3000"));