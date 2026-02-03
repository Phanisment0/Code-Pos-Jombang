const XLSX = require("xlsx");
const fs = require("fs");

function processExcel(inputFile, outputFile) {
	const workbook = XLSX.readFile(inputFile);
	const sheetName = workbook.SheetNames[0];
	const sheet = workbook.Sheets[sheetName];

	const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

	for (let i = 1; i < data.length; i++) {
		const row = data[i];
		const cellJ = row[9];

		if (typeof cellJ === "string") {
			const commaCount = (cellJ.match(/,/g) || []).length;

			if (commaCount === 2) {
				const parts = cellJ.split(",");
				const lastValue = parts[parts.length - 1].trim();

				let result;

				switch (lastValue) {
					case "JOMBANG":
						result = 61415;
						break;
				
					case "TEMBELANG":
						result = 61452;
						break;
				
					case "PLOSO":
						result = 61453;
						break;
				
					case "KUDU":
						result = 61454;
						break;
				
					case "KABUH":
						result = 61455;
						break;
				
					case "PLANDAAN":
						result = 61456;
						break;
				
					case "MEGALUH":
						result = 61457;
						break;
				
					case "NGUSIKAN":
						result = 61458;
						break;
				
					case "PERAK":
						result = 61461;
						break;
				
					case "BANDARKEDUNGMULYO":
						result = 61462;
						break;
				  
					case "BANDAR KEDUNG MULYO":
						result = 61462;
						break;

					case "GUDO":
						result = 61463;
						break;
				
					case "DIWEK":
						result = 61471;
						break;
				
					case "NGORO":
						result = 61473;
						break;
				
					case "BARENG":
						result = 61474;
						break;
				
					case "MOJOWARNO":
						result = 61475;
						break;
				
					case "WONOSALAM":
						result = 61476;
						break;
				
					case "PETERONGAN":
						result = 61481;
						break;
				
					case "MOJOAGUNG":
						result = 61482;
						break;
				
					case "SUMOBITO":
						result = 61483;
						break;
				
					case "KESAMBEN":
						result = 61484;
						break;
				
					case "JOGOROTO":
						result = 61485;
						break;
				
					default:
						result = 0;
				}

				row[10] = result;
			}
		}
	}

	const newSheet = XLSX.utils.aoa_to_sheet(data);
	workbook.Sheets[sheetName] = newSheet;

	XLSX.writeFile(workbook, outputFile);
}

processExcel(
	"/storage/emulated/0/Download/input.xlsx",
	"/storage/emulated/0/Download/output.xlsx"
);
