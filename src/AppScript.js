/**
 * Desarrollador: Erlin Olaya
 * Fecha : 02/03/2023
 * Descripcion: Script encargado de cruzar información y comparar por referencias, fecha, nombre empresa y valor de hoja de google sheett vs mismos valores en un CSV y si encuentra filas que coincidan se actualiza el sheet con información que viene en el CSV
*/
function partidasIdentificar() {

	// Se inicializa cuerpo de respuesta
	const body = {}
	let response = {}

	try {
		Logger.log("Inicia ejecución de AppScript");
		// Se busca archivo CSV
		response = searchFileCSV();

		const csvFile = response.body.csvFile;
		const csvData = response.body.csvData;

		// Se verifica si el archivo y el contenido fue encontrado
		Logger.log("Se valida si archivo CSV existe en la ruta");

		if (!csvFile || !csvData) {
			return buildResponse(ResponseStatus.Successful, response.body);
		}

		// Se obtiene los valores del CSV
		Logger.log("Se obtiene data del CSV");

		response = getSheet();
		const sheetValuesPartidas = response.body.sheetValuesPartidas;
		const sheetGoogle = response.body.sheetPartidas;

		// Se verifica si el proceso fue exitoso
		if (response.status == ResponseStatus.Unsuccessful) {
			return response;
		}

		//Se itera por los registros del CSV
		Logger.log("Inicia loop para array de la data del archivo CSV");
		for (let x = 0; x <= csvData.length - 1; x++) {

			//Valores a comparar en sheet google
			let fechaCSV = csvData[x][0]
			let valorCSV = csvData[x][1]
			let referenciaCSV = csvData[x][2];
			let nombreEmpresaCSV = csvData[x][3];

			//Valores a registrar en sheet si hay coincidencia
			let fechaPartidasCSV = csvData[x][4];
			let usuarioPartidasSCV = csvData[x][5];

			Logger.log("FechaCSV: " + fechaCSV + " valorCSV: " + valorCSV + " referenciaCSV: " + referenciaCSV + " nombreEmpresaCSV: " + nombreEmpresaCSV);

			Logger.log("Inicia loop para array de la data del Sheet");
			//Se itera por los registros del Sheet
			for (let y = 0; y <= sheetValuesPartidas.length - 1; y++) {

				let dateSheetAux = sheetValuesPartidas[y][0];

				//Cambiar formato de fecha
				let dateSheet = Utilities.formatDate(dateSheetAux, Session.getScriptTimeZone(), "dd-MM-yyyy");

				let valueSheet = sheetValuesPartidas[y][2]
				let referenciaSheet = sheetValuesPartidas[y][3];
				let nombreEmpresaSheet = sheetValuesPartidas[y][5];

				Logger.log("\nFechaCSV" + fechaCSV + "  == FechaSheet" + dateSheet
					+ "\nvalorCSV" + valorCSV + "== valueSheet" + valueSheet
					+ "\nreferenciaCSV" + referenciaCSV + " == referenciaSheet" + referenciaSheet
					+ "\nnombreEmpresaCSV" + nombreEmpresaCSV + " == nombreEmpresaSheet" + nombreEmpresaSheet);

				if (fechaCSV == dateSheet && valorCSV == valueSheet && referenciaCSV == referenciaSheet && nombreEmpresaCSV == nombreEmpresaSheet) {

					Logger.log("Se encontró coindidencia de referencias y de valor");
					let fechaPartidas = sheetGoogle.getRange(y + 2, Params.UPDATE_PARTIDAS.COLUMN_SHEET_PARTIDAS_FECHA);
					fechaPartidas.setValue(fechaPartidasCSV);
					let usuarioPartidas = sheetGoogle.getRange(y + 2, Params.UPDATE_PARTIDAS.COLUMN_SHEET_PARTIDAS_USUARIO);
					usuarioPartidas.setValue(usuarioPartidasSCV);
				}

			}
		}

		//Mover archivo CSV a carpeta Backup
		backupLastUpdate();

		body.message = 'Ejecución realizada exitosamente';
		return buildResponse(ResponseStatus.Successful, body)

	} catch (error) {
		body.message = 'Error en la funcion principal de partidas' + `${error.stack} `;
		Logger.log(`ERROR - ${body.message} `);
		criticalCheckPoint(body.message);
		return buildResponse(ResponseStatus.Unsuccessful, body);
	}
}

/**
* Mueve los archivos insumo generados en anteriores ejecuciones
* a la carpeta de backup
*/
const backupLastUpdate = () => {
	// Se obtiene carpeta
	const folder = DriveApp.getFolderById(Params.UPDATE_PARTIDAS.CSV_FOLDER_PARTIDAS);
	// Se obtiene carpeta de respaldo de CSV
	const backupFolder = DriveApp.getFolderById(Params.UPDATE_PARTIDAS.BACKUP_CSV_FOLDER);

	Logger.log("INFO - Buscando archivos a mover dentro de la carpeta de salida");

	// Se busca archivos de insumo
	const csvsFilter = `title contains '${Params.UPDATE_PARTIDAS.CSV_NAME}' `;
	const inputDbFiles = folder.searchFiles(csvsFilter);

	// Move all output files to backup folder
	while (inputDbFiles.hasNext()) {
		const outputFile = inputDbFiles.next();
		outputFile.moveTo(backupFolder);

		Logger.log(
			`INFO - Archivo ${outputFile.getName()} movido a la carpeta de backup`
		);
	}
};


//Obtiene el archivo csv
const searchFileCSV = () => {
	// Se inicializa cuerpo de respuesta
	const body = {}

	Logger.log("Inicia función searchFileCSV() para encontrar archivo CSV y data respectiva");
	// Se obtiene la carpeta del archivo CSV
	const csvFolder = DriveApp.getFolderById(Params.UPDATE_PARTIDAS.CSV_FOLDER_PARTIDAS);

	// Se busca Archivo CSV

	//Fecha Actual
	let newDate = new Date();
	let dateSheetAux = Utilities.formatDate(newDate, Session.getScriptTimeZone(), "dd-MM-yyyy");
	let date = dateSheetAux.replace(/-/g, "");

	const nameFile = Params.UPDATE_PARTIDAS.CSV_NAME;
	const filePartidasCSV = nameFile.concat(date);


	Logger.log("filePartidasCSV: " + filePartidasCSV)

	const files = csvFolder.getFilesByName(filePartidasCSV + ".csv")

	//const files = csvFolder.searchFiles(filePartidasCSV);

	// Se obtiene archivo y su contenido si existe alguno
	if (files.hasNext()) {

		Logger.log("INFO - La carpeta contiene un archivo");

		// Se obtiene contenido del archivo 
		const file = files.next();
		const content = file.getBlob().getDataAsString();

		// Se convierte la informacion a un array (2D)
		// y se eliminan encabezados
		const csvData = Utilities.parseCsv(content);
		csvData.splice(0, 1);

		body.message = 'Achivo CSV e información obtenida';
		body.csvFile = file;
		body.csvData = csvData;

		Logger.log(`INFO - ${body.message} `);
		return buildResponse(ResponseStatus.Successful, body)
	}
	// La carpeta no tiene archivos
	body.message = 'No se encontraron archivos de CSV para procesar';
	Logger.log(`INFO - ${body.message} `);
	return buildResponse(ResponseStatus.Successful, body);
}

//Obtiene el archivo de google sheet
const getSheet = () => {

	// Se inicializa cuerpo de respuesta
	const body = {
		sheetValuesPartidas: "",
		sheetPartidas: ""
	}

	try {
		// Se obtiene archivo e información de la Sheet de google
		Logger.log("Inicia función getSheet() para obtener data del Sheet");

		let fileSheet = SpreadsheetApp.openById(Params.UPDATE_PARTIDAS.ID_SHEET_PARTIDAS);
		//var hoja = fileSheet.getSheetByName(Params.UPDATE_PARTIDAS.NAME_SHEET);
		let sheetGoogle = fileSheet.getSheetByName(Params.UPDATE_PARTIDAS.NAME_SHEET);
		let lastColumn = sheetGoogle.getLastColumn();
		let lastRow = sheetGoogle.getLastRow();
		const valores = sheetGoogle.getRange(2, 1, lastRow - 1, lastColumn).getValues();

		body.message = 'Archivo e información obtenida';
		body.sheetValuesPartidas = valores;
		body.sheetPartidas = sheetGoogle;

		return buildResponse(ResponseStatus.Successful, body);

	} catch (error) {
		body.message = 'Se genero un error en la busqueda de la sheet';
		Logger.log(`Error - ${body.message}.Description: ${error.stack} `);
		return buildResponse(ResponseStatus.Unsuccessful, body);
	}
}
