
global.$ = $;

var cfg = {
	all: {
				 export_excel:{
								  show: {
											order: "订单",
											stage: "工序",
											month: "月份",
											price: "价格",
											num: "数量",
											'status': "完款与否"
										},
								  name: "总览",
							  },
				 import_excel: {
								   map: {
											'A': 'order',
											'C': 'num',
											'E': 'price',
											'F': 'status'
										},
								   default_values: {
													   'order': '',
													   'num': 0,
													   'price': 0,
													   'status': ''
												   },
								   sheet_name: "总览",
							   },
			 },

	caijian: {
				 export_excel:{
								  show: {
											order: "订单",
											month: "月份",
											price: "价格",
											num: "数量",
											'status': "完款与否"
										},
								  name: "裁剪",
							  },
				 import_excel: {
								   map: {
											'A': 'order',
											'C': 'num',
											'E': 'price',
											'F': 'status'
										},
								   default_values: {
													   'order': '',
													   'num': 0,
													   'price': 0,
													   'status': ''
												   },
								   sheet_name: "裁剪",
							   },
			 },
	yapiao: {
				 export_excel:{
								  show: {
											order: "订单",
											month: "月份",
											price: "价格",
											num: "数量",
											'status': "完款与否"
										},
								  name: "压朴"
							  },
				 import_excel: {
								   map: {
											'A': 'order',
											'C': 'num',
											'E': 'price',
											'F': 'status'
										},
								   default_values: {
													   'order': '',
													   'num': 0,
													   'price': 0,
													   'status': ''
												   },
								  sheet_name: "压朴"
							   },
			},
	chefeng: {
				 export_excel:{
								  show: {
											order: "订单",
											month: "月份",
											price: "价格",
											num: "数量",
											'status': "完款与否"
										},
								  name: "车缝"
							  },
				 import_excel: {
								   map: {
											'A': 'order',
											'C': 'num',
											'E': 'price',
											'F': 'status'
										},
								   default_values: {
													   'order': '',
													   'num': 0,
													   'price': 0,
													   'status': ''
												   },
								  sheet_name: "车缝"
							   },
			},
	dazao: {
				 export_excel:{
								  show: {
											order: "订单",
											month: "月份",
											price: "价格",
											num: "数量",
											'status': "完款与否"
										},
								  name: "打枣"
							  },
				 import_excel: {
								   map: {
											'A': 'order',
											'C': 'num',
											'E': 'price',
											'F': 'status'
										},
								   default_values: {
													   'order': '',
													   'num': 0,
													   'price': 0,
													   'status': ''
												   },
								  sheet_name: "打枣"
							   },
			},
	fengyan: {
				 export_excel:{
								  show: {
											order: "订单",
											month: "月份",
											price: "价格",
											num: "数量",
											'status': "完款与否"
										},
								  name: "凤眼"
							  },
				 import_excel: {
								   map: {
											'A': 'order',
											'C': 'num',
											'E': 'price',
											'F': 'status'
										},
								   default_values: {
													   'order': '',
													   'num': 0,
													   'price': 0,
													   'status': ''
												   },
								  sheet_name: "凤眼"
							   },
			},
	shuixi: {
				 export_excel:{
								  show: {
											order: "订单",
											month: "月份",
											price: "价格",
											num: "数量",
											'status': "完款与否"
										},
								  name: "水洗"
							  },
				 import_excel: {
								   map: {
											'A': 'order',
											'C': 'num',
											'E': 'price',
											'F': 'status'
										},
								   default_values: {
													   'order': '',
													   'num': 0,
													   'price': 0,
													   'status': ''
												   },
								  sheet_name: "水洗"
							   },
			},
	xiuxian: {
				 export_excel:{
								  show: {
											order: "订单",
											month: "月份",
											price: "价格",
											num: "数量",
											'status': "完款与否"
										},
								  name: "修线"
							  },
				 import_excel: {
								   map: {
											'A': 'order',
											'C': 'num',
											'E': 'price',
											'F': 'status'
										},
								   default_values: {
													   'order': '',
													   'num': 0,
													   'price': 0,
													   'status': ''
												   },
								  sheet_name: "修线"
							   },
			},
};


function boot(succ, fail) {
	var cmd = require("node-cmd");
	cmd.get("mysql.server status", function(stat) {
		if (/not running/.test(stat)) {
			cmd.get("mysql.server start", function(stat) {
				if (/SUCC/.test(stat)) {
					succ(stat);
				} else {
					fail(stat);
				}
			});
		} else {
			succ(stat);
		}
	});
}


/* 调试窗口 */
function message(msg, json = false)
{
	if (json) {
		msg = JSON.stringify(msg);
	}
	$("#dialog").html(msg);
}

function is_num(s)
{
	if (s!=null && s!="")
	{
		return !isNaN(s);
	}
	return false;
}

function is_order_id(id) {
	return /^[\w\(\)\-\+]+$/.test(id) ;
}

function filter_orders(rows) {
	var result = [];
	for (i = 0; i < rows.length; i ++) {
		if (rows[i].order && is_order_id(rows[i].order)) {
			result.push(rows[i]);
		}
	}
	return result;
}

function exc2rows(map, sheet)
{
	var result = {
		'total': 0,
		'rows': []
	};

	var reg = /^([A-Z]+)(\d+)$/;

	var last_row = 1;
	var row = {};
	for (cell in sheet) {
		result.total += 1;
		var tmp = reg.exec(cell);
		if (tmp && typeof(map[tmp[1]]) != 'undefined') {

			if (last_row != tmp[2]) {
				last_row = tmp[2];
				result.rows.push(row);
				row = {};
			}
			row[map[tmp[1]]] = sheet[cell].v;
		}
	}
	if (row) {
		result.rows.push(row);
	}

	for (i = 0; i < result.rows.length; i ++) {
		for (f in map) {
			f = map[f];
			if (typeof(result.rows[i][f]) == 'undefined') {
				result.rows[i][f] = null;
			}
		}
	}

	return result;

}

function mysql_conn()
{
	var db = require('mysql').createConnection({
		host : 'localhost',
		user : 'root',
		password : '',
		database : 'weihua'
	});
	return db;
}


/*
 * 将字段转换成易于读取的形式，例如: caijian -> 裁剪
 * @param	rows,	[{stage: caijian}]
 * @param	map,	{stage: {caijian: "裁剪", xiuxian: "修线"}}
 * @return			[{stage: "裁剪"}]
 */
function map_rows(rows, map)
{
	for (i in rows) {
		for (field in rows[i]) {
			var value = rows[i][field];
			if (typeof(map[field]) != 'undefined' && typeof(map[field][value] != 'undefined')) {
				rows[i][field] = map[field][value];
			}
		}
	}
	return rows;
}

function save_rows(table, rows, fields)
{
	if (!rows.length) {
		return ;
	}
	db = mysql_conn();
	db.connect();

	var f_sql = [];
	for(f in fields) {
		f_sql.push('`' + f + '`');
	}

	var values = [];
	for (i = 0; i < rows.length; i ++) {
		var val = '(';
		for (f in fields) {
			val += "'" + ((rows[i][f]) ? rows[i][f] : fields[f]) + "',";
		}
		val = val.replace(/,$/,"") + ")";
		values.push(val);
	}

	var sql = 'INSERT INTO `' + table + '` (' + f_sql.join(",") + ') VALUES ' + values.join(",");
	db.query(sql, function(err, rows, fields) {
		if (err) {
			message(JSON.stringify(err) + ' SQL:' + sql);
		} else {
			display();
			message("导入成功!");
		};
	});

	db.end();
	
}

function sheet_from_array_of_arrays(data, opts) {
	var ws = {};
	var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	var XLSX = require('xlsx');
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell = {v: data[R][C] };
			if(cell.v == null) continue;
			var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n'; cell.z = XLSX.SSF._table[14];
				cell.v = datenum(cell.v);
			}
			else cell.t = 's';

			ws[cell_ref] = cell;
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
	return ws;
}

function clean_search()
{
	$('#month').val('');
	$('#order').val('');
}

function build_where_sql(){
	var stage = curr_stage();
	var where = ' where 1 = 1 ';
	if (stage != 'all') {
		where += ' and `stage` = \'' + stage + '\'';
	}
	var month = $('#month').val();
	var order = $('#order').val();
	if (order) {
		where += ' and `order` = \''+ order +'\'';
	}
	if (month) {
		where += ' and `month` = \''+ month +'\'';
	}
	return where;
}

function export_excel()
{
	db = mysql_conn();
	db.connect();
	var stage = curr_stage();
	var fields = cfg[stage].export_excel.show;
	var sql = "select * from `cost` " + build_where_sql();
	var path = "/Users/waiwah/Desktop/" + cfg[stage].export_excel.name + ".xlsx";
	db.query(sql, function(err, rows, db_fields) {
		if (err) {
			message(JSON.stringify(err) + ' SQL:' + sql);
			return;
		};

		rows = map_rows(rows, {stage:{caijian:"裁剪", chefeng:"车缝", shuixi:"水洗", fengyan:"凤眼",dazao:"打枣",yapiao:"压朴",xiuxian:"修线"}});

		var xlsx = require('xlsx');

		var header = [];
		for (f in fields) {
			header.push(fields[f]);
		}

		var exc_rows = [];
		exc_rows.push(header);
		for(i = 0; i < rows.length; i ++) {
			var r = [];
			for (f in fields) {
				r.push(rows[i][f]);
			}
			exc_rows.push(r);
		}
		var data = {SheetNames:["Sheet1"], Sheets:{"Sheet1":sheet_from_array_of_arrays(exc_rows,{})}};
		xlsx.writeFile(data, path);
	});

	db.end();
}


function run() {

	display('cost', ['order', 'month', 'price', 'num', 'status']);

	/* 加载模块 */
	xlsx = require('xlsx');

	/* 导出文件 */
	$("#export").click(function(){
		export_excel('cost', {order:'订单', month:'月份', price:'价格', num:'数量', 'status':'完款与否'}, './test.xlsx');
	});

	/* 切换工序 */
	$("#listView li").click(function(){
		if ( $("#listView li").hasClass("list-item-active") ) {
			$("#listView li").removeClass("list-item-active");
		}
		$(this).addClass("list-item-active");
		clean_search();
		display();
	});

	/* 搜索按钮 */
	$("#search").click(function(){
		display();
	});

	/*
	 *	文件导入
	 */
	$("#inputFile").unbind("change");
	$("#inputFile").change(function(ev) {
		var path = $(this).val();
		var workbook = xlsx.readFile(path);
		for (stage in cfg) {
			if (stage != "all") {
				import_excel(stage, workbook);
			}
		}

	});
}

function log(msg, json=false)
{
	if (json) {
		msg = JSON.stringify(msg);
	}
	var fs = require("fs");
	fs.appendFile("./my.log", msg + "\n", function(err){
		if (err) {
			message(err, true);
		}
	});
}

function import_excel(stage, workbook)
{

		var excel2mysql = cfg[stage].import_excel.map;
		var default_fields = cfg[stage].import_excel.default_values;
		var keyword = cfg[stage].import_excel.sheet_name;

		// 获取表格（通过关键字来匹配表名)
		//var worksheet = workbook.Sheets[workbook.SheetNames[1]];
		var debug = '';

		for (i = 0; i < workbook.SheetNames.length; i ++) {
			var sname = unescape(workbook.SheetNames[i].replace(/&#x([A-Za-z0-9]+);/g, "%u$1"));
			debug += "(" + sname + ":" + keyword + "," + sname.indexOf(keyword) + ')';
			debug += "(" + sname.charCodeAt(0) + "," + sname.charCodeAt(1) + ":" + keyword.charCodeAt(0) + "," + keyword.charCodeAt(1) + ')';
			if (sname.indexOf(keyword) != -1) {
				var worksheet = workbook.Sheets[workbook.SheetNames[i]];

				default_fields.month = 201606;
				default_fields.stage = stage;

				var data = exc2rows(excel2mysql, worksheet);
				data.rows = filter_orders(data.rows);
				save_rows('cost', data.rows, default_fields	);
			}
		}

}

function curr_stage()
{
	return $(".list-item-active a").attr("name");
}

function display()
{
	var table = 'cost';
	db = mysql_conn();
	db.connect();
	var stage = curr_stage();
	var fields = cfg[stage].export_excel.show;
	var sql = "select * from `cost`" + build_where_sql();
	db.query(sql, function(err, rows, db_fields) {
		if (err) {
			message(JSON.stringify(err) + ' SQL:' + sql);
			return;
		};
		rows = map_rows(rows, {stage:{caijian:"裁", chefeng:"车", shuixi:"水", fengyan:"凤眼",dazao:"打枣",yapiao:"压朴",xiuxian:"修线"}});

		var thead = '<thead><tr>';
		for (f in fields) {
			thead += '<th class="text-left">' + fields[f] + '</th>';	
		}
		thead += '</tr></thead>';

		var tbody = '<tbody class="table-hover">';
		for (i = 0; i < rows.length; i ++) {
			tbody += '<tr>';
			var v = rows[i];
			for (f in fields) {
				if (typeof(v[f]) != 'undefined') {
					tbody += '<td class="text-left">' + v[f] + '</td>';
				} else {
					tbody += '<td class="text-left">undefined</td>';
				}
			}
			tbody += '</tr>';
		}
		tbody += '</tbody>';
		$('#list').html(thead + tbody);
	});

	db.end();
}

$(document).ready(function() {
	boot(run, function(stat) {
		message("FAILED! mysql start failed, error: " + stat);
	});
});
 
