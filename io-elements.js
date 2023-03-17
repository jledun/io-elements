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
const TYPES_ELEMENTS_SIMPLE = [TYPES_ELEMENTS_COMPLET[0], TYPES_ELEMENTS_COMPLET[5], TYPES_ELEMENTS_COMPLET[6]];
const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ', 'DA', 'DB', 'DC', 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL', 'DM', 'DN', 'DO', 'DP', 'DQ', 'DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ', 'EA', 'EB', 'EC', 'ED', 'EE', 'EF', 'EG', 'EH', 'EI', 'EJ', 'EK', 'EL', 'EM', 'EN', 'EO', 'EP', 'EQ', 'ER', 'ES', 'ET', 'EU', 'EV', 'EW', 'EX', 'EY', 'EZ', 'FA', 'FB', 'FC', 'FD', 'FE', 'FF', 'FG', 'FH', 'FI', 'FJ', 'FK', 'FL', 'FM', 'FN', 'FO', 'FP', 'FQ', 'FR', 'FS', 'FT', 'FU', 'FV', 'FW', 'FX', 'FY', 'FZ', 'GA', 'GB', 'GC', 'GD', 'GE', 'GF', 'GG', 'GH', 'GI', 'GJ', 'GK', 'GL', 'GM', 'GN', 'GO', 'GP', 'GQ', 'GR', 'GS', 'GT', 'GU', 'GV', 'GW', 'GX', 'GY', 'GZ', 'HA', 'HB', 'HC', 'HD', 'HE', 'HF', 'HG', 'HH', 'HI', 'HJ', 'HK', 'HL', 'HM', 'HN', 'HO', 'HP', 'HQ', 'HR', 'HS', 'HT', 'HU', 'HV', 'HW', 'HX', 'HY', 'HZ'];
const minLineNumber = 3;
const maxLineNumber = 400;
// paramétrage des version de la supervision
const SPV_VERSIONS = [{ // générique
		spvVersion: "IWS",
		elemSheetName: "Elements",
		asserSheetName: 'Asservissements',
		cycleSheetName: 'Cycles',
		dataDef: [],
	}, {
		spvVersion: "IWS", // IWS
		elemSheetName: "Elements IWS",
		asserSheetName: 'Asservissements IWS',
		cycleSheetName: 'Cycles IWS',
		dataDef: [],
	}, {
		spvVersion: "InTouch", // InTouch
		elemSheetName: "Elements InTouch",
		asserSheetName: 'Asservissements InTouch',
		cycleSheetName: 'Cycles InTouch',
		dataDef: [],
	},
];
// dataDef générique
for (let i = 0; i < 80; i++) {
	SPV_VERSIONS[0].dataDef.push({
		id: i,
		ref: i.toString().padStart(2, '0'),
		sqlType: 'TEXT',
	});
}
SPV_VERSIONS[0].dataDef.push({
		id: 80,
		ref: 'Depart',
		sqlType: 'TEXT',
});
SPV_VERSIONS[0].dataDef.push({
		id: 81,
		ref: 'Cycle_Simple',
		sqlType: 'TEXT',
});
SPV_VERSIONS[0].dataDef.push({
		id: 82,
		ref: 'Cycle_Complet',
		sqlType: 'TEXT',
});
SPV_VERSIONS[0].dataDef.push({
		id: 83,
		ref: 'Destination',
		sqlType: 'TEXT',
});
// dataDef IWS
SPV_VERSIONS[1].dataDef = [].concat(SPV_VERSIONS[0].dataDef);
// dataDef InTouch
for (let i = 0; i < 15; i++) {
	SPV_VERSIONS[2].dataDef.push({
		id: i,
		ref: `MW${900 + i}`,
		sqlType: 'INTEGER',
	});
}
SPV_VERSIONS[2].dataDef.push({
	id: 15,
	ref: 'Origine',
	sqlType: 'TEXT',
});
SPV_VERSIONS[2].dataDef.push({
	id: 16,
	ref: 'Cycle',
	sqlType: 'TEXT',
});
SPV_VERSIONS[2].dataDef.push({
	id: 17,
	ref: 'Description',
	sqlType: 'TEXT',
});
SPV_VERSIONS[2].dataDef.push({
	id: 18,
	ref: 'Destination',
	sqlType: 'TEXT',
});
SPV_VERSIONS[2].dataDef.push({
	id: 19,
	ref: 'Visible',
	sqlType: 'INTEGER',
});

let spvVersion = 0;
let columns = [];

const PROMPT_MODES = [
	{
		exec: "createAsservissementSheet", 
		message: "Créer ou initialiser la feuille 'Asservissements' et initialiser la feuille 'Cycles' dans le fichier 'Elements.xlsx'",
		confirmation: "ATTENTION : cette action est destructrice !!! Etes-vous sûr de vouloir supprimer les feuilles 'Asservissements' et 'Cycles' dans le fichier 'Elements.xlsx' ?"
	},
	{
		exec: "readAsservissement", 
		message: "Générer les chemins dans la feuille 'Cycles' de 'Elements.xlsx' et mettre à jour la base de données 'Cycles.mdb'",
		confirmation: "ATTENTION : cette action est destructrice !!! Etes-vous sûr de vouloir mettre à jour les données dans 'Cycles.mdb' ?"
	}
];

const getColumnRef = (col) => {
	return cols[col - 1];
}

const getMWColumnRef = (MW) => {
	// format attendu : "MW9xx"
	if (MW.length < 5) return false;
	if (MW.slice(0, 3) !== "MW9") return false;
	const num = Number(MW.slice(3));
	return cols[num];
}

const initCycleSheet = (workbook) => {
	return new Promise((resolve, reject) => {
		console.log(chalk.white(`Suppression de la feuille de calcul 'Cycles' si elle existe...`));
		let cyclesSheet;
		for (let i = 0; i < SPV_VERSIONS.length; i++) {
			cyclesSheet = workbook.getWorksheet(SPV_VERSIONS[i].cycleSheetName);
			if (cyclesSheet) {
				workbook.removeWorksheet(cyclesSheet.id);
				console.log(chalk.green(`Feuille de calcul '${SPV_VERSIONS[i].cycleSheetName}' supprimée...`));
			}
		}
		console.log(chalk.white(`Création d'une feuille de calcul '${SPV_VERSIONS[spvVersion].cycleSheetName}' vierge pour stocker les chemins...`));
		cyclesSheet = workbook.addWorksheet(SPV_VERSIONS[spvVersion].cycleSheetName);
		if (!cyclesSheet) {
			return reject(`La feuille de calcul ${SPV_VERSIONS[spvVersion].cycleSheetName} n'a pas pu être créée, abandon.`);
		}
		console.log(chalk.green(`Feuille de calcul '${SPV_VERSIONS[spvVersion].cycleSheetName}' créée...`));
		
		console.log(chalk.white(`Génération et saisie des titres des colonnes dans la feuille de calcul '${SPV_VERSIONS[spvVersion].cycleSheetName}'...`));
		columns = SPV_VERSIONS[spvVersion].dataDef.map(def => def.ref);
		columns.forEach((col, i) => cyclesSheet.getCell(`${getColumnRef(i + 1)}1`).value = col);
		console.log(chalk.green(`Feuille de calcul '${SPV_VERSIONS[spvVersion].cycleSheetName}' complète.`));
		return resolve(cyclesSheet);
	});
}

const readElementsFromElementSheet = (elementsSheet) => {
	// parcours de chaque ligne
	let load = loading({
		"text": "Lecture de la liste des éléments : numéro, désignation, type, nature origine et/ou destination...",
		"color": "white",
		"interval": 100,
		"frames": ["◰", "◳", "◲", "◱"]
	}).start();
	let elements = [];
	for (let i = minLineNumber; i < maxLineNumber; i++) {
		switch(spvVersion) {
			case 2: // InTouch
				elements.push({
					elemRow: i,
					asserRow: 0,
					foundInAsservissementSheet: false,
					numero: elementsSheet.getCell('A'.concat(i)).value,
					mnemo: elementsSheet.getCell('B'.concat(i)).value,
					description: elementsSheet.getCell('C'.concat(i)).value,
					type: elementsSheet.getCell('D'.concat(i)).value,
					origine: elementsSheet.getCell('E'.concat(i)).value,
					destination: elementsSheet.getCell('F'.concat(i)).value,
					isOrigine: elementsSheet.getCell('E'.concat(i)).value != null,
					isDestination: elementsSheet.getCell('F'.concat(i)).value != null
				});
				break;
				
			case 0: // Défaut
			case 1: // IWS
			default:
				elements.push({
					elemRow: i,
					asserRow: 0,
					foundInAsservissementSheet: false,
					numero: elementsSheet.getCell('A'.concat(i)).value,
					mnemo: elementsSheet.getCell('B'.concat(i)).value,
					type: elementsSheet.getCell('C'.concat(i)).value,
					origine: elementsSheet.getCell('D'.concat(i)).value,
					destination: elementsSheet.getCell('E'.concat(i)).value,
					isOrigine: elementsSheet.getCell('D'.concat(i)).value != null,
					isDestination: elementsSheet.getCell('E'.concat(i)).value != null
				});
				break;
		}
	}
	load.stop();
	console.log(chalk.green("Liste des éléments lue."));
	return elements.filter(elm => elm.mnemo);
}

const readMW900ParamSheet = (workbook, origines) => {
	return new Promise(async (resolve, reject) => {
		const sheet = workbook.getWorksheet("Param MW900");
		if (!sheet) {
			return reject(new Error(`Feuille de calcul "Param MW900" non trouvée, abandon`));
		}
		const newOrigines = [];
		for (let i = 3; i < 43; i++) {
			const numeroOrigine = sheet.getCell("B".concat(i)).value;
			if (!numeroOrigine) {
				continue;
			}
			const toUpgrade = origines.filter(elm => elm.origine === numeroOrigine);
			if (toUpgrade.length <= 0) {
				continue;
			}
			let mw900Params = SPV_VERSIONS[spvVersion].dataDef.map((ref, j) => {
				const col = getColumnRef(3 + j);
				if (j < 13) {
					// lecture des paramètres origines : MW900 à MW912
					return {
						...ref,
						value: 0,
						used: sheet.getCell(`${col}${i}`).value !== null,
						elevateur: sheet.getCell(`${col}3`).value === TYPES_ELEMENTS_COMPLET[5],
					};
				}else{
					// autres colonnes
					return {
						...ref,
						value: 0,
						used: true,
						elevateur: false,
					};
				}
			});
			
			// longueur maxi du chemin hors éléments d'origine et de destination
			const maxLen = mw900Params.filter((ref, i) => i >= 0 && i <=12 && ref.used).length;
			for (let elem of toUpgrade) {
				newOrigines.push({...elem, result: mw900Params, maxLen});
			}
		}
		if (newOrigines.length !== origines.length) {
			return reject(new Error(`Défaut de paramétrage MW900 : le nombre d'élements d'origines ne correspond pas.`));
		}
		return resolve(newOrigines);
	});
}

const createAsservissementSheet = () => {
	return new Promise(async (resolve, reject) => {
		const workbook = new ExcelJS.Workbook();
		
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
		console.log(chalk.white(`Recherche d'une feuille de calcul ${SPV_VERSIONS.map(spvVer => "'" + spvVer.elemSheetName + "'").join(', ')}...`));
		let elemSheet;
		for (let i = 0; i < SPV_VERSIONS.length; i++) {
			elemSheet = workbook.getWorksheet(SPV_VERSIONS[i].elemSheetName);
			if (elemSheet) {
				spvVersion = i;
				break;
			}
		}
		if (!elemSheet) {
			return reject(new Error(`Les feuilles de calcul ${SPV_VERSIONS.map(spvVer => "'" + spvVer.elemSheetName + "'").join(', ')} n'existent pas, abandon.`));
		}
		console.log(chalk.green(`Feuille de calcul '${SPV_VERSIONS[spvVersion].elemSheetName}' trouvée`));

		// Lecture de la liste des éléments dans la feuille de calcul 'Elements'
		let elements = readElementsFromElementSheet(elemSheet);
		
		// génération de la feuille "Asservissements"
		console.log(chalk.white(`Suppression de la feuille de calcul ${SPV_VERSIONS.map(ver => "'" + ver.asserSheetName + "'").join(", ")}, si elle existe`));
		let asserSheet;
		for (let i = 0; i < SPV_VERSIONS.length; i++) {
			asserSheet = workbook.getWorksheet(SPV_VERSIONS[i].asserSheetName);
			if (asserSheet) {
				workbook.removeWorksheet(asserSheet.id);
				console.log(chalk.green(`Feuille de calcul '${SPV_VERSIONS[i].asserSheetName}' supprimée...`));
			}
		}
		console.log(chalk.white(`Création de la feuille de calcul '${SPV_VERSIONS.map(ver => "'" + ver.asserSheetName + "'").join(", ")}'...`));
		asserSheet = workbook.addWorksheet(SPV_VERSIONS[spvVersion].asserSheetName);
		if (!asserSheet) {
			return reject(`La feuille de calcul ${SPV_VERSIONS[spvVersion].asserSheetName} n'a pas pu être créée, abandon.`);
		}
		console.log(chalk.green(`Feuille de calcul '${SPV_VERSIONS[spvVersion].asserSheetName}' créée`));
		
		asserSheet.getCell('B2').value = "Graphes éléments";
		asserSheet.fill = {
			type: 'pattern',
			pattern:'solid',
			fgColor:{argb:'Ffffffff'},
		};
		asserSheet.border = {
			top: {style:'thin'},
			left: {style:'thin'},
			bottom: {style:'thin'},
			right: {style:'thin'}
		};
		
		load.text = `Recopie de la liste des éléments dans les lignes et colonnes de la feuille de calcul '${SPV_VERSIONS[spvVersion].asserSheetName}'...`;
		load.start();
		let index = 3;
		for (let elem of elements) {
			// numéro d'élement et mnémo sur les 2 premières colonne
			asserSheet.getCell('A'.concat(index)).value = elem.numero;
			asserSheet.getCell('B'.concat(index)).value = elem.mnemo;
			// numéro d'élément et mnémo sur les 2 premières lignes
			asserSheet.getCell(getColumnRef(index).concat('1')).value = elem. numero;
			asserSheet.getCell(getColumnRef(index).concat('2')).value = elem. mnemo;
			// croix sur origine seule ou destination seule
			if (elem.isOrigine && !elem.isDestination) {
				asserSheet.getCell(getColumnRef(index).concat('2')).border = {
					diagonal: {up: true, down: true, style:'thin', color: {argb:'ff000000'}}
				};
			}else if (elem.isDestination && !elem.isOrigine) {
				asserSheet.getCell('B'.concat(index)).border = {
					diagonal: {up: true, down: true, style:'thin', color: {argb:'ff000000'}}
				};
			}
			// création de la croix = cas impossible
			asserSheet.getCell(getColumnRef(index).concat(index)).value = "X";
			asserSheet.getCell(getColumnRef(index).concat(index)).fill = {
				type: 'pattern',
				pattern:'solid',
				fgColor:{argb:'F0808000'},
			};
			index++;
		}
		load.stop();
		console.log(chalk.green(`Liste des éléments de la feuille de calcul '${SPV_VERSIONS[spvVersion].asserSheetName}' mis à jour.`));
		
		// initialisation de la feuille Cycles
		try{
			await initCycleSheet(workbook);
		}catch(e){
			return reject(e);
		}
		
		load.text = "Enregistrement du classeur 'Elements.xlsx'...";
		load.start();
		await workbook.xlsx.writeFile(path.join(".", "Elements.xlsx"));
		load.stop();
		console.log(chalk.green("Enregistrement du classeur 'Elements.xlsx' terminé."));
		resolve(`Fichier 'Elements.xlsx' prêt, libre à vous de renseigner les éléments aval dans la feuille de calcul '${SPV_VERSIONS[spvVersion].asserSheetName}' :-)`);
	});
}

const getOrigine = (origines, element) => {
	return origines.find(orig => orig.numero === element.numero);
}

const readAsservissement = () => {
	return new Promise(async (resolve, reject) => {
		const workbook = new ExcelJS.Workbook();
		
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
		console.log(chalk.white(`Recherche de la feuille de calcul ${SPV_VERSIONS.map(ver => "'" + ver.elemSheetName + "'").join(', ')}`));
		let elemSheet;
		for (let i = 0; i < SPV_VERSIONS.length; i++) {
			elemSheet = workbook.getWorksheet(SPV_VERSIONS[i].elemSheetName);
			if (elemSheet) {
				spvVersion = i;
				break;
			}
		}
		
		if (!elemSheet) {
			return reject(new Error(`Les feuilles de calcul ${SPV_VERSIONS.map(ver => "'" + ver.elemSheetName + "'").join(', ')} n'existent pas, abandon.`));
		}
		console.log(chalk.green(`Feuille de calcul '${SPV_VERSIONS[spvVersion].elemSheetName}' trouvée`));

		// lecture de la liste des éléments
		let elements = readElementsFromElementSheet(elemSheet);

		// Lecture de la feuille "ASSERVISSEMENTS"
		console.log(chalk.white(`Lecture de la feuille de calcul '${SPV_VERSIONS[spvVersion].asserSheetName}'`));
		const asserSheet = workbook.getWorksheet(SPV_VERSIONS[spvVersion].asserSheetName);
		if (!asserSheet) {
			return reject(new Error(`La feuille de calcul '${SPV_VERSIONS[spvVersion].asserSheetName}' n'existe pas, abandon.`));
		}
		console.log(chalk.green(`Feuille de calcul '${SPV_VERSIONS[spvVersion].asserSheetName}' lue.`));
		
		// Comparaison des éléments présents dans la feuille Element et dans la feuille Asservissements
		console.log(chalk.white(`Comparaison des éléments présents dans la feuille 'Element' et dans la feuille 'Asservissements'`));
		let ctl1 = [];
		for (let i = minLineNumber; i < maxLineNumber; i++) {
			let mnemo = asserSheet.getCell('B'.concat(i)).value;
			if (mnemo) {
				ctl1.push({
					numero: asserSheet.getCell('A'.concat(i)).value,
					mnemo: mnemo
				});
				// ajout à l'élément de la ligne correspondante dans la feuille asservissements
				let mnemos = elements.map((elem, i) => ({...elem, id: i})).filter(elem => elem.mnemo === mnemo && !elem.foundInAsservissementSheet);
				if (mnemos.length > 0) {
					mnemos[0].asserRow = i;
					mnemos[0].foundInAsservissementSheet = true;
					elements[mnemos[0].id] = mnemos[0];
				}
			}
		}
		const ctl2 = elements.map(elem => ({
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
		let origines = elements.filter(elem => elem.isOrigine);
		try{
			origines = await readMW900ParamSheet(workbook, origines);
		}catch(e){
			return reject(e);
		}
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
		
		console.log(chalk.white(`Démarrage enregistrement dans la feuille de calcul '${SPV_VERSIONS[spvVersion].cycleSheetName}'`));
		// Mise en forme des données pour enregistrement 
		console.log(chalk.white(`Mise en forme des données pour correspondre à la structure de la table.`));
		let dataToInsert = [];
		switch (spvVersion) {
			case 0: // générique
			case 1: // IWS
			default: // par défaut
				dataToInsert = chemins.map(chemin => {
					let insertString  = [];
					// colonnes 00 à 79
					for (let i = 0; i < 80; i++) {
						insertString.push((chemin[i] && chemin[i].direction) ? chemin[i].direction.toString() : "0" );
					}
					// colonne Départ
					insertString.push(chemin[0].mnemo);
					// colonne Cycle_simple
					insertString.push(
						chemin.filter(elem => TYPES_ELEMENTS_SIMPLE.indexOf(elem.type) >= 0)
						.filter((elem, i, ar) => (i > 0 && i < ar.length -1))
						.map(elem => elem.mnemo)
						.join(" - ")
					);
					// colonne Cycle_complet
					insertString.push(
						chemin.filter((elem, i, ar) => (i > 0 && i < ar.length -1))
						.map(elem => elem.mnemo)
						.join(" - ")
					);
					// colonne destination
					insertString.push(chemin[chemin.length - 1].mnemo);
					return insertString;
				});
				break;
			
			case 2: // InTouch
				dataToInsert = chemins.map(chemin => {
					// traitement cas impossible : avec cette méthode de programmation, il faut toujours au moins 1 élement entre l'origine et la destination
					if (chemin.length <= 2) {
						throw new Error(`Cas impossible : chemin vide, veuillez corriger les asservissements. Origine : ${chemin[0].origine} - ${chemin[0].mnemo}, Destination : ${chemin[chemin.length - 1].destination} - ${chemin[chemin.length - 1].mnemo}`);
					}
					// on recopie le chemin en retirant origine et destination
					let tmpChemin = chemin
						.filter((elm, i, ar) => i > 0 && i < ar.length - 1) // retrait origine et destination
						.map((elm, i) => ({...elm, indexChemin: i}));		// ajout de l'index élément dans le chemin au cas où
					// création de l'objet résultant
					const origineChemin = getOrigine(origines, chemin[0]);
					// recopie avec init des valeurs
					const result = origineChemin.result.map(res => ({...res, value: 0}));
					
					// correction immédiate des valeurs directement attribuables
					// numéro de destination
					result[13].value = chemin[chemin.length - 1].destination;
					// numéro d'origine
					result[14].value = chemin[0].origine;
					// mnémo élément d'origine
					result[15].value = chemin[0].mnemo;
					// description cycle simple
					result[16].value = tmpChemin.map(elm => elm.mnemo.replace(/\+/gi, '-')).join(" - ");
					// description cycle complète
					result[17].value = tmpChemin.map(elm => elm.description.replace(/\+/gi, '-')).join(" - ");
					// mnémo élément de destination
					result[18].value = chemin[chemin.length - 1].mnemo;
					// option visibilité
					result[19].value = 1;
					
					// traitement champs entre 900 et 912
					let t = {};
					switch (tmpChemin.length) {
						case 1:
							// cas particulier : 1 seul élément dans le chemin : selon paramétrage
							// on insère l'élément dans le premier emplacement utilisé ???
							// sauf si élément = élévateur : dans le premier emplacement élévateur utilisé ???
							t = tmpChemin.shift();
							for (let i = 0; i <= 12; i++) {
								if (result[i].used) {
									if (t.type === TYPES_ELEMENTS_COMPLET[5] && !result[i].elevateur) continue;
									result[i].value = t.numero;
									break;
								}
							}
							break;

						case 2:
							// cas particulier : 2 éléments dans le chemin : le premier en premier et de second en dernier
							// on insère l'élément dans le premier emplacement utilisé ???
							// sauf si élément = élévateur : dans le premier emplacement élévateur utilisé ???
							t = tmpChemin.shift();
							let nextI = 0;
							for (let i = 0; i <= 12; i++) {
								if (result[i].used) {
									if (t.type === TYPES_ELEMENTS_COMPLET[5] && !result[i].elevateur) continue;
									result[i].value = t.numero;
									nextI = i + 1;
									break;
								}
							}
							// on insère l'élément dans le dernier emplacement utilisé ??? 
							// sauf si élément est un élévateur : dans le premier emplacement élévateur utilisé ???
							t = tmpChemin.pop();
							if (t.type === TYPES_ELEMENTS_COMPLET[5]) {
								for (let i = nextI; i <= 12; i++) {
									if (result[i].used && result[i].elevateur) {
										result[i].value = t.numero;
										break;
									}
								}
							}else{
								for (let i = 12; i > 0; i--) {
									if (result[i].used) {
										result[i].value = t.numero;
										break;
									}
								}
							}
							break;

						default:
							// cas général : tous les éléments restant dans le chemin doivent être organisés selon les paramètres
							for (let i = 0; i <= 12; i++) {
								if (tmpChemin.length <= 0) break;
								if (tmpChemin[0].type === TYPES_ELEMENTS_COMPLET[5]) {
									if (result[i].used && result[i].elevateur) {
										result[i].value = tmpChemin.shift().numero;
										continue;
									}
								}else{
									if (result[i].used) {
										result[i].value = tmpChemin.shift().numero;
										continue;
									}
								}
							}
							// déplacement du dernier élément avant destination en MW912 si nécessaire et sauf si dernier élément est un élévateur qui doit rester à la place qui lui a été attribuée
							if (result[12].used && result[12].value <= 0) {
								for (let i = 12; i > 0; i--) {
									if (result[i].value > 0) {
										// recherche dans le chemin du type d'élément détecté
										const indexChemin = chemin.findIndex(chm => chm.numero === result[i].value);
										// et on annule l'opération si le dernier élément du chemin est un élévateur
										if (indexChemin >= 0 && chemin[indexChemin].type === TYPES_ELEMENTS_COMPLET[5]) break;
										result[12].value = result[i].value;
										result[i].value = 0;
										break;
									}
								}
							}
							break;
					}
					// arf, on a raté des trucs !!!
					if (tmpChemin.length > 0) {
						console.log(`Chemin abandonné - longueur restante : ${tmpChemin.length}, Origine: ${chemin[0].mnemo}, Destination: ${chemin[chemin.length - 1].mnemo}`);
						return null;
					}
					return result.map(elm => elm.value);
				}).filter(xlsxPhrase => xlsxPhrase); // retrait des éléments du tableau revenu avec une valeur nulle.
				break;
		}
		
		console.log(chalk.white(`Initialisation de la feuille de calcul 'Cycles'...`));
		
		// initialisation de la feuille Cycles
		let cyclesSheet;
		try{
			cyclesSheet = await initCycleSheet(workbook);
		}catch(e){
			return reject(e);
		}
		
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
		
		let confirm;
		const dbConfirm = new Confirm({
			name: "confirm",
			message: "Souhaitez-vous générer le fichier Microsoft Access Cycles.mdb ?"
		});
		try{
			confirm = await dbConfirm.run();
		}catch(e) {
			console.log(chalk.yellow("Erreur de confirmation"));
			return reject(e);
		}
		if (!confirm) {
			return resolve(`Ok, fichier Elements.xlsx à jour, reste à mettre à jour le fichier Cycles.mdb`);
		}
		
		console.log(chalk.white("Ouverture de la base de données 'Cycles.mdb'"));
		const db = ADODB.open(`Provider=Microsoft.Jet.OLEDB.4.0;Data Source=${path.join(".", "Cycles.mdb")};`);
		if (!db) {
			return reject(new Error("Fichier 'Cycles.mdb' non trouvé"));
		}
		console.log(chalk.green("Fichier 'Cycles.mdb' ouvert."));
		
		// paramétrage du nom de la table
		const tableName = SPV_VERSIONS[spvVersion].cycleSheetName;
		
		// requête création de la table ${tableName}
		let createRequest = `CREATE TABLE \`${tableName}\` (`;
		// inclusion de la définition des colonnes
		createRequest = createRequest.concat(
			SPV_VERSIONS[spvVersion].dataDef
				.map(def => `${def.ref} ${def.sqlType}`)
				.join(",\n")
		);
		createRequest = createRequest.concat(');');
		
		// nettoyage des données existantes
		try{
			console.log(chalk.white(`Nettoyage des données existantes...`));
			load.text = "Nettoyage en cours...";
			load.start();
			await db.execute(`DROP TABLE \`${tableName}\`;`);
		}catch(e){
			if (e.process.code !== -2147217865) {
				load.stop();
				console.log(chalk.yellow(`Erreur lors du nettoyage de la table '${tableName}', abandon.`));
				return reject(e);
			}
		}
		try{
			await db.execute(createRequest);
			load.stop();
			console.log(chalk.green(`Table '${tableName}' prête à recevoir les nouvelles données.`));
		}catch(e){
			load.stop();
			console.log(chalk.yellow(`Erreur lors de la création de la table '${tableName}', abandon.`));
			return reject(e);
		}
		
		// insertion des données existantes
		console.log(chalk.white(`Génération des requêtes d'insertion des chemins dans la table '${tableName}'...`));
		// Microsoft Access ne tolère pas l'insertion de plusieurs lignes sauf si c'est le résultat d'une requête select :-|
		let requests = dataToInsert.map(line => {
			return `INSERT INTO \`${tableName}\` (\`${columns.join("\`, \`")}\`) VALUES ('${line.join("', '")}');`;
		});
		console.log(chalk.green('Requêtes prêtes'));
		
		console.log(chalk.white(`Insertion des chemins dans la table '${tableName}'`));
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
			console.log(chalk.yellow(`Erreur lors de l'insertion dans la table '${tableName}', abandon.`));
			return reject(e);
		}
	});
}

const getAvalElements = (source, elements, feuille) => {
	const tmp = [];
	for (let elem of elements) {
		const cellValue = feuille.getCell(getColumnRef(elem.asserRow).concat(source.asserRow)).value;
		if (typeof cellValue === "number") {
			tmp.push({
				direction: cellValue + source.numero,
				element: Object.assign({}, elem)
			});
		}
	}
	return tmp.sort((a, b) => (a.direction - b.direction));
}

const explorer = (chemins, lifo, chemin, elements, feuille) => {
	
	// contrôles spécifiques selon version supervision
	switch(spvVersion) {
		case 2: // supervision InTouch
			// contrainte : pas plus de 3 élévateurs par chemin
			const elementsElevateurs = chemin.filter(elem => elem.type === TYPES_ELEMENTS_COMPLET[5]);
			if (elementsElevateurs.length > 3) return;
			
			// contrainte : pas plus de maxLen éléments par chemin (MW900 à MW912), maxLen défini dans readMW900ParamSheet
			// + 2 car l'élément d'origine et l'élément de destination ne sont pas comptés dans maxLen.
			if (chemin.length > chemin[0].maxLen + 2) return;
			break;
			
		case 0:	// générique
		case 1: // IWS
		default:// par défaut
			break;
	}
	
	// destination atteinte
	if (chemin.length > 1 && chemin[chemin.length - 1].isDestination) {
		chemin[chemin.length - 1].direction = chemin[chemin.length - 1].numero;
		chemins.push([].concat(chemin));
		return;
	}
	
	// lecture des éléments aval
	const aval = getAvalElements(chemin[chemin.length - 1], elements, feuille);
	if (aval.length > 0) {
		for (let elementAval of aval) {
			// recherche d'une double utilisation d'une élément dans un chemin sauf si l'élément en double est à l'origine, au tour suivant, cet élément en double devra forcément être une destination
			const indexDoublon = chemin.findIndex((elementChemin, i) => elementChemin.numero === elementAval.element.numero && i > 0);
			if (indexDoublon < 0) {
				// copie du chemin
				const tmp = [];
				for (let elm of chemin) {
					tmp.push(Object.assign({}, elm));
				}
				
				// affectation de la direction de l'élément
				tmp[tmp.length - 1].direction = elementAval.direction;
				
				// ajout de l'élément aval
				tmp.push(elementAval.element);
				
				// copie chemin temporaire dans la lifo pour recherche du chemin suivant au prochain batch
				lifo.push([].concat(tmp));
			}
		}
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
