const express = require("express");
const path = require("path");
const excel = require("exceljs");
const fs = require("fs");
const wordDoc = require("docxtemplater");
const multer = require("multer");
const { default: Docxtemplater } = require("docxtemplater");
const port = 3000;

const app = express();
const upload = multer({ dest: "uploads/" });

app.get("", (req, res) => {
	res.sendFile(path.join(__dirname, "../FRONTEND/index.html"));
});

app.post("/BACKEND/uploads", upload.single("excelFile"), (req, res) => {
	const users = new excel.Workbook();
	const data = [];
	users.xlsx
		.readFile(req.file.path)
		.then(function () {
			const userEntry = users.getWorksheet(1);
			userEntry.eachRow(function (row, rowNumber) {
				if (rowNumber === 1) return;
				data.push = {
					name: row.getCell(1).value,
					degree: row.getCell(2).value,
					distinction: row.getCell(3).value,
					date: row.getCell(4).value,
					courseInstructor: row.getCell(5).value,
				};
				console.log(row.getCell(1).value);
			});

			try {
				const templatePath = path.join(__dirname, "../template2.docx");
				const content = fs.readFileSync(templatePath, "binary");
				docx = new wordDoc();
				docx.loadZip(content);
				//docx.setData({students: data});

				try {
					docx.render();

					const outputpath = path.join(__dirname, "generated.docx");
					const buffer = docx
						.getZip()
						.generate({ type: "nodebuffer" });
					fs.writeFileSync(outputpath, buffer);

					res.sendFile(outputpath);
				} catch (error) {
					console.error(error);
					res.status(500).send("File could not be generated");
				}
			} catch (error) {
				console.error(error);
				res.status(500).send("Loading failed");
			}
		})
		.catch((readFileError) => {
			console.error("Error reading Excel file:", readFileError);
			res.status(500).send("Error reading Excel file");
		});
});

app.listen(port);
