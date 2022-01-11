const XLSX = require('xlsx');
const electron = require('electron').remote;

const EXTENSIONS = "xls|xlsx|xlsm|xlsb|xml|csv|txt|dif|sylk|slk|prn|ods|fods|htm|html".split("|");

const processWb = function(wb) {
	const HTMLOUT = document.getElementById('htmlout');
	const XPORT = document.getElementById('exportBtn');
	XPORT.disabled = false;
	HTMLOUT.innerHTML = "";
	var data = [];
	data.push(["ESTABELECIMENTO", "PROFISSIONAL", "CBO", "AMB CH"]);
	wb.SheetNames.forEach(function(sheetName) {
		var ws = wb.Sheets[sheetName];
		ws = Object.entries(ws).filter(item => item[0].charAt(0) !== 'S' && item[0] !== '!ref');
		ws = Object.fromEntries(ws);
		ws = Object.values(ws);
		ws.splice(0, 11);
		var equipe = '';
		var estabelecimento = '';
		for(var i = 0; i < ws.length;i++){
			var value = ws[i].v;
			var aux_equipe = null;
			if(value == "Estabelecimento :")
				aux_equipe = ws[i+5].v;			
			if(value == "Estabelecimento :" && aux_equipe != equipe){
				equipe = aux_equipe;
				estabelecimento = ws[i+1].v;
				i += 14;
				continue;
			} else if(value == "Total de Profissionais :"){
				i += 1;
				continue;
			} else {
				var profissional = ws[i].v;
				var cbo = ws[i+1];
				var amb_ch = ws[i+2];
				if(cbo != undefined && amb_ch != undefined){
					cbo = cbo.v;
					amb_ch = amb_ch.v;
					data.push([estabelecimento, profissional, cbo, amb_ch]);
					i+= 5;
				}
				continue;
			}
		}
	});
	var ws = XLSX.utils.aoa_to_sheet(data);
	HTMLOUT.innerHTML = XLSX.utils.sheet_to_html(ws, {editable: true});
};

const handleReadBtn = async function() {
	const o = await electron.dialog.showOpenDialog({
		title: 'Selecione um arquivo',
		filters: [{
			name: "Spreadsheets",
			extensions: EXTENSIONS
		}],
		properties: ['openFile']
	});
	if(o.filePaths.length > 0) processWb(XLSX.readFile(o.filePaths[0]));
};

const exportXlsx = async function() {
	const HTMLOUT = document.getElementById('htmlout');
	const wb = XLSX.utils.table_to_book(HTMLOUT);
	const o = await electron.dialog.showSaveDialog({
		title: 'Sarvar arquivo como',
		filters: [{
			name: "Spreadsheets",
			extensions: EXTENSIONS
		}]
	});
	//console.log(o.filePath);
	XLSX.writeFile(wb, o.filePath);
	electron.dialog.showMessageBox({ message: "Planilha exportada para " + o.filePath, buttons: ["OK"] });
};

const readBtn = document.getElementById('readBtn');
const exportBtn = document.getElementById('exportBtn');

readBtn.addEventListener('click', handleReadBtn, false);
exportBtn.addEventListener('click', exportXlsx, false);