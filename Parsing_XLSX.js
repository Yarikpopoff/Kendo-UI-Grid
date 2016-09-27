function dataProcess(bstr) {

	var workbook = XLSX.read(bstr, {type:"binary"});
	var dataJson = XLSX.utils.sheet_to_json(workbook.Sheets.Sheet1);
	// var dataJson = XLSX.utils.sheet_to_row_object_array(workbook.Sheets.Sheet1);
		
	// debugger;

	// JSON file preparation
	var dataGrid = dataJson.map(function(obj, i) {
		var newObj = {};

		// remove spaces and special characters
		Object.keys(obj).forEach(function(key, j) {
			newObj[key.replace(/[@\s]/g, '')] = obj[key];
		})
		
		// edit datetime
		newObj['PubTime'] = newObj['PubTime'].slice(0, -3);

			// Change calculation
		var n = (obj.AllEODClose - obj.AllTONLast) * 100 / obj.AllTONLast;
			
		newObj['Change'] = Math.round( n * 100 ) / 100;

			//symbol color choice
		if (newObj['Change'] > 0) {
			newObj['Color'] = 'rgb(0, 150, 149)';
		} else {
				newObj['Color'] = 'rgb(245, 255, 68);'
		}
			
		// Change to string
		newObj['Change'] += '%';
			
		// Volume formatting
		var str = '', j = 0;
		for (i = newObj.Volume.length - 1; i >= 0; i--) {
			str += newObj.Volume[i];
			j++;
			if (j == 3) {
				str += ','; 
				j = 0;
			}
		} 
		newObj['Volume'] = Array.prototype.reduceRight.call(str, function(memo, item) {
			memo += item; return memo
		}, '');
			return newObj;
	});

	$("#grid > *").remove();
	$("#grid").kendoGrid({ 
		dataSource: {
			data: dataGrid,
			pageSize: 10,
		},
		height: 500,
		// groupable: true,
        sortable: true,
		pageable: {
        // refresh: true,
        pageSizes: true,
        buttonCount: 9
        },
        selectable: true,
        columns: [{
 			template: "<div class='item' style='background-color:#: Color #;'></div><div class='company'>#: Company #</div>",
 			field: "Company",
 			title: "SYMBOL",
 			width: 70
 		}, {
 			field: "PubTime",
 			title: "DATETIME",
 			width: 100
 		}, {
 			field: "Source",
 			title: "SOURSE",
 			width: 70
 		}, {
 			field: "HeadLine",
 			title: "HEADLINES",
 			width: 500,
 		}, {
 			field: "PriceNews",
 			title: "PRICE@NEWS",
 			width: 90
 		}, {
 			field: "Change",
 			title: "CHANGE%",
 			width: 80
 		}, {
 			field: "LastPrice",
 			title: "LAST PRICE",
 			width: 80
 		}, {
 			field: "Volume",
 			title: "VOLUME",
 			width: 70
 		},{
 			field: "VolumeRatio",
 			title: "VOL RATIO",
 			width: 80
 		}],
 		change: function() {
 			var selectedRows = this.select();
			$(".article").text(this.dataItem(selectedRows[0]).HeadLine);
 		} 
	});	
}

$('.xlsx-file').on('click', function(e) {
	
	/* set up XMLHttpRequest */
	$('.file-name').text('Interview_NewsDB.xlsx');
	var url = "/Interview_NewsDB.xlsx";
	var oReq = new XMLHttpRequest();
	oReq.open("GET", url, true);
	oReq.responseType = "arraybuffer";

	oReq.onload = function(e) {
  		
  		var arraybuffer = oReq.response;

  		/* convert data to binary string */
  		var data = new Uint8Array(arraybuffer);
  		var arr = new Array();
  		for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
  		var bstr = arr.join("");

  		dataProcess(bstr);

	}

	oReq.send();
});

xlsxFile.addEventListener('change', fileSelect); 

function fileSelect(e) {
	// file selection
	var file = e.target.files[0];

	$('.file-name').text(file.name);
	var reader  = new FileReader();
	reader.readAsBinaryString(file);
	reader.addEventListener('load', fileLoad);
	
	function fileLoad() {
		
		dataProcess(reader.result);

	}
}
		
