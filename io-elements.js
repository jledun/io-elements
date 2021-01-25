/* Code sous licence MIT :
 *
 * Copyright 2021 Julien Ledun <j.ledun@iosystems.fr>
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software
 * and associated documentation files (the "Software"), to deal in the Software without restriction, 
 * including without limitation the rights to use, copy, modify, merge, publish, distribute, 
 * sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is 
 * furnished to do so, subject to the following conditions:
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
 * LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN 
 * NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
 * WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH 
 * THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 *
 * */ 
 
 // APPLICATION DE CREATION DE CHEMINS A PARTIR D'UNE LISTE D'ELEMENTS PARAMETRES DANS UN TABLEAU EXCEL
 
'use strict';

// data tools
const ExcelJS = require('exceljs');
const ADODB = require('node-adodb');

// cli tools
const { Confirm, Select } = require('enquirer');
const loading =  require('loading-cli');
const chalk = require('chalk');
const figlet = require('figlet');

// filesystem tool
const path = require('path');

// constants
const TYPES_ELEMENTS_COMPLET = ["Moteur", "Pendulaire", "Boite2D", "Boite3D", "Trappe", "Elevateur", "Contenant"];
const TYPES_ELEMENTS_SIMPLE = ["Moteur", "Elevateur", "Contenant"];
const PROMPT_MODES = [
	{
		exec: "createAsservissementSheet", 
		message: "Créer ou mettre à jour la feuille 'Asservissements' et initialiser la feuille 'Cycles' dans le fichier 'Elements.xlsx'",
		confirmation: "ATTENTION : cette action est destructrice !!! Etes-vous sûr de vouloir supprimer la feuille 'Cycles' dans le fichier 'Elements.xlsx' ?"
	},
	{
		exec: "readAsservissement", 
		message: "Générer les chemins dans la feuille 'Cycles' de 'Elements.xlsx' et mettre à jour la base de données 'Cycles.mdb'",
		confirmation: "ATTENTION : cette action est destructrice !!! Etes-vous sûr de vouloir mettre à jour les données dans 'Cycles.mdb' ?"
	}
];

const getColumnRef = (col) => {
	const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ', 'DA', 'DB', 'DC', 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL', 'DM', 'DN', 'DO', 'DP', 'DQ', 'DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ', 'EA', 'EB', 'EC', 'ED', 'EE', 'EF', 'EG', 'EH', 'EI', 'EJ', 'EK', 'EL', 'EM', 'EN', 'EO', 'EP', 'EQ', 'ER', 'ES', 'ET', 'EU', 'EV', 'EW', 'EX', 'EY', 'EZ', 'FA', 'FB', 'FC', 'FD', 'FE', 'FF', 'FG', 'FH', 'FI', 'FJ', 'FK', 'FL', 'FM', 'FN', 'FO', 'FP', 'FQ', 'FR', 'FS', 'FT', 'FU', 'FV', 'FW', 'FX', 'FY', 'FZ', 'GA', 'GB', 'GC', 'GD', 'GE', 'GF', 'GG', 'GH', 'GI', 'GJ', 'GK', 'GL', 'GM', 'GN', 'GO', 'GP', 'GQ', 'GR', 'GS', 'GT', 'GU', 'GV', 'GW', 'GX', 'GY', 'GZ', 'HA', 'HB', 'HC', 'HD', 'HE', 'HF', 'HG', 'HH', 'HI', 'HJ', 'HK', 'HL', 'HM', 'HN', 'HO', 'HP', 'HQ', 'HR', 'HS', 'HT', 'HU', 'HV', 'HW', 'HX', 'HY', 'HZ'];
	return cols[col - 1];
}

const createAsservissementSheet = () => {
	return new Promise(async (resolve, reject) => {
		const workbook = new ExcelJS.Workbook();
		let elements = [];
		
		// Lecture fichier Excel "Elements.xlsx"
		let load = loading({
			"text": "Ouverture fichier 'Elements.xlsx'",
			"color": "white",
			"interval": 100,
			"frames": ["◰", "◳", "◲", "◱"]
		}).start();
		try{
			await workbook.xlsx.readFile(path.join(".", "Elements.xlsx"));
			load.stop();
			console.log(chalk.green("Fichier 'Elements.xlsx' ouvert."));
		}catch(e){
			load.stop();
			return reject(e);
		}

		// Lecture de la feuille "Elements"
		console.log(chalk.white("Recherche de la feuille de calcul 'Elements'"));
		const elemSheet = workbook.getWorksheet('Elements');
		if (!elemSheet) {
			return reject(new Error("La feuille de calcul 'Elements' n'existe pas, abandon."));
		}
		console.log(chalk.green("Feuille de calcul 'Elements' trouvée"));

		// parcours de chaque ligne
		load.text = "Parcours de la liste des éléments...";
		load.start();
		for (let i = 3; i < 219; i++) {
			elements.push({
				row: i,
				numero: elemSheet.getCell('A'.concat(i)).value,
				mnemo: elemSheet.getCell('B'.concat(i)).value,
				type: elemSheet.getCell('C'.concat(i)).value,
				origine: elemSheet.getCell('D'.concat(i)).value,
				destination: elemSheet.getCell('E'.concat(i)).value,
				isOrigine: elemSheet.getCell('D'.concat(i)).value != null,
				isDestination: elemSheet.getCell('E'.concat(i)).value != null
			});
		}
		load.stop();
		console.log(chalk.green("Liste des éléments lue."));
		
		// génération de la feuille "Asservissements"
		console.log(chalk.white("Recherche de la feuille de calcul 'Asservissements'"));
		let asserSheet = workbook.getWorksheet('Asservissements');
		if (!asserSheet) {
			console.log(chalk.white("La feuille de calcul 'Asservissements' n'existe pas, création..."));
			asserSheet = workbook.addWorksheet('Asservissements');
			console.log(chalk.green("Feuille de calcul 'Asservissements' créée"));
		}
		console.log(chalk.green("Feuille de calcul 'Asservissements' trouvée"));
		
		load.text = "Recopie de la liste des éléments dans les lignes et colonnes de la feuille de calcul 'Asservissements'...";
		load.start();
		for (let elem of elements) {
			asserSheet.getCell('A'.concat(elem.row)).value = elem.numero;
			asserSheet.getCell('B'.concat(elem.row)).value = elem.mnemo;
			asserSheet.getCell(getColumnRef(elem.row).concat('1')).value = elem. numero;
			asserSheet.getCell(getColumnRef(elem.row).concat('2')).value = elem. mnemo;
			asserSheet.getCell(getColumnRef(elem.row).concat(elem.row)).value = "X";
		}
		load.stop();
		console.log(chalk.green("Liste des éléments de la feuille de calcul 'Asservissements' mis à jour."));
		
		console.log(chalk.white(`Suppression de la feuille de calcul 'Cycles' si elle existe...`));
		let cyclesSheet = workbook.getWorksheet('Cycles');
		if (cyclesSheet) {
			workbook.removeWorksheet(cyclesSheet.id);
			console.log(chalk.green(`Feuille de calcul 'Cycles' supprimée...`));
		}
		console.log(chalk.white(`Création d'une feuille de calcul 'Cycles' vierge pour stocker les chemins...`));
		cyclesSheet = workbook.addWorksheet('Cycles');
		console.log(chalk.green(`Feuille de calcul 'Cycles' créée...`));
		
		console.log(chalk.white(`Génération et saisie des titres des colonnes dans la feuille de calcul 'Cycles'...`));
		let columns = [];
		// colonnes 00 à 79
		for (let i = 0; i < 80; i++) columns.push((i.toString().length < 2) ? "0".concat(i.toString()) : i.toString());
		// colonnes Départ, Cycle_Simple, Cycle_Complet, Destination
		columns.push("Depart");
		columns.push("Cycle_Simple");
		columns.push("Cycle_Complet");
		columns.push("Destination");
		columns.forEach((col, i) => cyclesSheet.getCell(`${getColumnRef(i + 1)}1`).value = col);
		console.log(chalk.green(`Feuille de calcul 'Cycles' complète.`));
		
		load.text = "Enregistrement du classeur 'Elements.xlsx'...";
		load.start();
		await workbook.xlsx.writeFile(path.join(".", "Elements.xlsx"));
		load.stop();
		console.log(chalk.green("Enregistrement du classeur 'Elements.xlsx' terminé."));
		resolve("Fichier 'Elements.xlsx' prêt, libre à vous de renseigner les éléments aval dans la feuille de calcul 'Asservissement' :-)");
	});
}

const readAsservissement = () => {
	return new Promise(async (resolve, reject) => {
		const workbook = new ExcelJS.Workbook();
		let elements = [];
		
		// Lecture fichier Excel "Elements.xlsx"
		let load = loading({
			"text": "Ouverture du classeur 'Elements.xlsx'",
			"color": "white",
			"interval": 100,
			"frames": ["◰", "◳", "◲", "◱"]
		}).start();
		try{
			await workbook.xlsx.readFile(path.join(".", "Elements.xlsx"));
			load.stop();
			console.log(chalk.green("Fichier 'Elements.xlsx' ouvert."));
		}catch(e){
			load.stop();
			return reject(e);
		}

		// Lecture de la feuille "Elements"
		console.log(chalk.white(`Lecture de la feuille de calcul 'Elements'`));
		const elemSheet = workbook.getWorksheet('Elements');
		if (!elemSheet) {
			return reject(new Error("La feuille de calcul 'Elements' n'existe pas, abandon."));
		}
		console.log(chalk.green("Feuille de calcul 'Elements' trouvée"));

		// lecture de la liste des éléments
		load.text = `Lecture de la liste des éléments : numéro, désignation, type, nature origine et/ou destination`;
		load.start();
		for (let i = 3; i < 219; i++) {
			elements.push({
				row: i,
				numero: elemSheet.getCell('A'.concat(i)).value,
				mnemo: elemSheet.getCell('B'.concat(i)).value,
				type: elemSheet.getCell('C'.concat(i)).value,
				origine: elemSheet.getCell('D'.concat(i)).value,
				destination: elemSheet.getCell('E'.concat(i)).value,
				isOrigine: elemSheet.getCell('D'.concat(i)).value != null,
				isDestination: elemSheet.getCell('E'.concat(i)).value != null
			});
		}
		load.stop();
		console.log(chalk.green("Liste des éléments lue."));

		// Lecture de la feuille "ASSERVISSEMENTS"
		console.log(chalk.white(`Lecture de la feuille de calcul 'Asservissements'`));
		const asserSheet = workbook.getWorksheet('Asservissements');
		if (!asserSheet) {
			return reject(new Error(`La feuille de calcul "Asservissements" n'existe pas, abandon.`));
		}
		console.log(chalk.green("Feuille de calcul 'Asservissements' lue."));
		
		// Comparaison des éléments présents dans la feuille Element et dans la feuille Asservissements
		console.log(chalk.white(`Comparaison des éléments présents dans la feuille 'Element' et dans la feuille 'Asservissements'`));
		const ctl1 = [];
		for (let i = 3; i < 219; i++) {
			ctl1.push({
				row: i,
				numero: asserSheet.getCell('A'.concat(i)).value,
				mnemo: asserSheet.getCell('B'.concat(i)).value
			});
		}
		const ctl2 = elements.map(elem => ({
			row: elem.row,
			numero: elem.numero,
			mnemo: elem.mnemo
		}));
		if (ctl1.length != ctl2.length) {
			return reject(new Error(`Le nombre d'éléments ne correspond pas : la feuille de calcul 'Elements' contient ${ctl2.length} éléments alors qu'il y a ${ctl1.length} dans 'Asservissements'. Abandon.`));
		}
		ctl1.forEach((c1, i) => {
			if (JSON.stringify(c1) !== JSON.stringify(ctl2[i])) {
				console.log(`Elements différents détectés :`);
				console.log(`Dans la feuille de calcul 'Elements' : `, ctl2[i]);
				console.log(`Dans la feuille de calcul 'Asservissements' : `, c1);
			}
		});
		if (JSON.stringify(ctl1) !== JSON.stringify(ctl2)) {
			return reject(new Error(`La liste des éléments ne correspond pas avec la table d'asservissements, abandon.`));
		}
		
		console.log(chalk.green("Tout est bon, c'est parti !"));
		
		console.log(chalk.white(`Sélection et contrôle des éléments d'origines...`));
		const origines = elements.filter(elem => elem.isOrigine);
		if (origines.length <= 0) {
			return reject(new Error(`Aucun emplacement d'origine n'a été saisi, abandon.`));
		}else{
			console.log(chalk.green(`Enregistrement des éléments d'origines : ${origines.length} origine(s) définie(s)`));
		}
		console.log(chalk.white(`Sélection des éléments de destination...`));
		const destinations = elements.filter(elem => elem.isDestination);
		if (destinations.length <= 0) {
			return reject(new Error(`Aucun emplacement de destination n'a été saisi, abandon.`));
		}else{
			console.log(chalk.green(`Enregistrement des éléments de destination : ${destinations.length} destination(s) définie(s)`));
		}
		
		// modélisation des chemins
		console.log(chalk.white(`Parcours en profondeur des enchaînements d'éléments en partant des origines connues pour modéliser tous les chemins possibles...`));
		load.text = "Modélisation en cours...";
		load.start();
		const chemins = [];
		// init des chemins à explorer : saisie des éléments en position 0
		const lifo = origines.map(orig => ([{...orig}]));
		while (lifo.length > 0) {
			const chemin = lifo.pop();
			if (chemin.length > 0) {
				explorer(chemins, lifo, chemin, elements, asserSheet);
			}
		}
		const maxLen = chemins.reduce((acc, chemin) => {
			return (chemin.length > acc) ? chemin.length : acc;
		}, 0);
		load.stop();
		console.log(chalk.green(`${chemins.length} chemins modélisés, longueur maxi des chemins modélisés : ${maxLen} :-)`));
		
		
		console.log(chalk.white(`Démarrage enregistrement dans la feuille de calcul Chemins`));
		// Mise en forme des données pour enregistrement 
		console.log(chalk.white(`Mise en forme des données pour correspondre à la structure de la table.`));
		const dataToInsert = chemins.map(chemin => {
			let insertString  = [];
			// colonnes 00 à 79
			for (let i = 0; i < 80; i++) {
				insertString.push((chemin[i] && chemin[i].direction) ? chemin[i].direction.toString() : "0" );
			}
			// colonnes Départ, Cycle_Simple, Cycle_Complet, Destination
			insertString.push(chemin[0].mnemo);
			insertString.push(
				chemin.filter(elem => TYPES_ELEMENTS_SIMPLE.indexOf(elem.type) >= 0)
				.map(elem => elem.mnemo)
				.join(" - ")
			);
			insertString.push(
				chemin.map(elem => elem.mnemo)
				.join(" - ")
			);
			insertString.push(chemin[chemin.length - 1].mnemo);
			return insertString; // "(\"" + insertString.join("\", \"") + "\")";
		});
		
		console.log(chalk.white(`Initialisation de la feuille de calcul 'Cycles'...`));
		let cyclesSheet = workbook.getWorksheet('Cycles');
		if (cyclesSheet) workbook.removeWorksheet(cyclesSheet.id);
		cyclesSheet = workbook.addWorksheet('Cycles');
		
		console.log(chalk.white(`Génération et saisie des titres des colonnes dans la feuille de calcul 'Cycles'...`));
		let columns = [];
		// colonnes 00 à 79
		for (let i = 0; i < 80; i++) columns.push((i.toString().length < 2) ? "0".concat(i.toString()) : i.toString());
		// colonnes Départ, Cycle_Simple, Cycle_Complet, Destination
		columns.push("Depart");
		columns.push("Cycle_Simple");
		columns.push("Cycle_Complet");
		columns.push("Destination");
		columns.forEach((col, i) => cyclesSheet.getCell(`${getColumnRef(i + 1)}1`).value = col);
		console.log(chalk.green(`Feuille de calcul 'Cycles' prête.`));
		
		// insertion des données existantes
		console.log(chalk.white('Saisie des chemins dans la feuille de calcul Cycles...'));
		dataToInsert.forEach((chemin, i) => {
			chemin.forEach((elem, j) => cyclesSheet.getCell(`${getColumnRef(j + 1)}${i + 2}`).value = elem);
		});
		console.log(chalk.green("Tous les chemins sont enregistrés dans la feuille de calcul 'Cycles'"));
		
		console.log(chalk.white("Enregistrement du classeur 'Elements.xlsx'..."));
		load.text = 'Enregistrement du classeur...';
		load.start();
		await workbook.xlsx.writeFile(path.join(".", "Elements.xlsx"));
		load.stop();
		console.log(chalk.green(`Classeur 'Elements.xlsx' à jour.`));
		
		console.log(chalk.white("Ouverture de la base de données 'Cycles.mdb'"));
		const db = ADODB.open(`Provider=Microsoft.Jet.OLEDB.4.0;Data Source=${path.join(".", "Cycles.mdb")};`);
		if (!db) {
			return reject(new Error("Fichier 'Cycles.mdb' non trouvé"));
		}
		console.log(chalk.green("Fichier 'Cycles.mdb' ouvert."));
		
		// nettoyage des données existantes
		try{
			console.log(chalk.white(`Nettoyage des données existantes...`));
			load.text = "Nettoyage en cours...";
			load.start();
			await db.execute('DELETE FROM `Cycles`;');
			load.stop();
			console.log(chalk.green(`Table 'Cycles' prête à recevoir les nouvelles données.`));
		}catch(e){
			load.stop();
			console.log(chalk.yellow("Erreur lors du nettoyage de la table Cycles, abandon."));
			return reject(e);
		}
		
		// insertion des données existantes
		console.log(chalk.white(`Génération des requêtes d'insertion des chemins dans la table 'Cycles'...`));	
		let requests = dataToInsert.map(line => {
			return `INSERT INTO \`Cycles\` (\`${columns.join("\`, \`")}\`) VALUES ('${line.join("', '")}');`; // "(\"" + insertString.join("\", \"") + "\")";
		});
		console.log(chalk.green('Requêtes prêtes'));
		
		console.log(chalk.white("Insertion des chemins dans la table 'Cycles'"));
		load.text = "Insertion des données en cours...";
		load.start();
		try{
			for (let req of requests) {
				await db.execute(req);
			}
			load.stop();
			return resolve(`Données importées avec succès dans la base de données Access Cycles.mdb !!! Merci d'avoir utilisé cette incroyable moulinette ;-)`);
		}catch(e){
			load.stop();
			console.log(chalk.yellow("Erreur lors de l'insertion dans la table 'Cycles', abandon."));
			return reject(e);
		}
	});
}

const getAvalElements = (source, elements, feuille) => {
	const tmp = [];
	for (let elem of elements) {
		const cellValue = feuille.getCell(getColumnRef(elem.row).concat(source.row)).value;
		if (typeof cellValue === "number") {
			tmp.push({
				direction: cellValue + elem.numero,
				element: Object.assign({}, elem)
			});
		}
	}
	return tmp.sort((a, b) => (a.direction - b.direction));
}

const explorer = (chemins, lifo, chemin, elements, feuille) => {
	if (chemin.length > 1 && chemin[chemin.length - 1].isDestination) {
		// console.log(`Destination atteinte.`);
		chemins.push(chemin);
		return;
	}
	const aval = getAvalElements(chemin[chemin.length - 1], elements, feuille);
	if (aval.length > 0) {
		aval.forEach(elementAval => {
			// recherche d'une double utilisation d'une élément dans un chemin
			if (chemin.findIndex(elementChemin => elementChemin.numero === elementAval.element.numero) < 0) {
				chemin[chemin.length - 1].direction = elementAval.direction;
				lifo.push(chemin.concat(elementAval.element));
			// }else{
				// console.log(`Elément déjà utilisé dans le chemin, arrêt de la poursuite de ce chemin...`);
			}
		});
	// }else{
		// console.log(`Pas d'élément aval, arrêt de la poursuite...`);
	}
}

const fcts = {
	createAsservissementSheet: createAsservissementSheet,
	readAsservissement: readAsservissement
};

const run = async () => {
	console.log(
		figlet.textSync(`IO Systems ${new Date().getFullYear()}`, {})
	);
	let answer, confirm, fct;
	console.log(chalk.yellow("Application développée sous licence MIT :"));
	console.log(chalk.yellow(`
Copyright 2021 Julien Ledun <j.ledun@iosystems.fr>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.`));
	const licenceAgree = new Confirm({
		name: "confirm",
		message: "Etes-vous d'accord avec les termes de la licence ?"
	});
	try{
		confirm = await licenceAgree.run();
	}catch(e) {
		console.log(chalk.yellow("Erreur accord de licence"));
		console.log(e);
		return;
	}
	if (!confirm) {
		console.log(chalk.yellow("Ok, licence refusée, pas de chemins trouvés :-|"));
		console.log(chalk.green("Abandon"));
		return;
	}
	console.log(chalk.yellow("licence MIT acceptée"));
	
	const modePrompt = new Select({
		name: "mode",
		message: "Choisissez un mode d'exécution du script",
		choices: PROMPT_MODES.map(p => p.message)
	});
	try{
		answer = await modePrompt.run();
		fct = PROMPT_MODES.find(p => p.message === answer).exec;
	}catch(e) {
		console.log(chalk.yellow("Erreur dans la sélection du mode de fonctionnement"));
		console.log(e);
		return;
	}
	const modeConfirm = new Confirm({
		name: "confirm",
		message: "Etes-vous sûr ?"
	});
	try{
		confirm = await modeConfirm.run();
	}catch(e) {
		console.log(chalk.yellow("Erreur de confirmation"));
		console.log(e);
		return;
	}
	if (confirm) {
		try{
			const result = await fcts[fct]();
			if (result) console.log(chalk.green(result));
		}catch(e){
			console.log(chalk.yellow("Une erreur s'est produite..."));
			console.log(e);
		}
	}
	console.log(chalk.green("Terminé :-)"));
}

run();