const mysql = require('mysql');
const XlsxBuilder = require('xlsx-io').XlsxBuilder;

const builder = new XlsxBuilder();

const connection = mysql.createConnection({
    host     : 'localhost',
    user     : 'Kirill',
    password : 'hunter2',
    database : 'whitebox_challenge'
  });

const outputFilename = 'output.xlsx';
const client_id = '1240'
const tables = {};

//create the initial table objects by isolating unique locale, shipping speed, and zone info from db
function start(){
    connection.connect();
    connection.query(`SELECT zone, locale, shipping_speed FROM rates WHERE client_id=${client_id} GROUP BY zone, locale, shipping_speed;`, function (error, results, fields) {
        if (error) throw error;

        for(let row of results){
            //create a table name and use as object key
            let tableName = `${row.locale} ${row.shipping_speed}`;
            if(tables[tableName] == undefined){
                //initialize object if undefined
                tables[tableName] = {
                    locale: row.locale, 
                    shipping_speed: row.shipping_speed, 
                    zones: [row.zone], 
                    columnNames: ["Start Weight", "End Weight"], //xlsx column names
                    data: [] //xlsx data rows
                };
            }else{
                //object defined, add zone info
                tables[tableName].zones.push(row.zone);
            }
        }
        //Create and append colum names from zone info
        for(let table of Object.values(tables)){
            let zoneNames = table.zones.map((val)=>{return `Zone ${val}`})
            table.columnNames = [...table.columnNames, ...zoneNames];
        }
        console.log("Created initial xlsx table sheet objects");
        queryAllZoneRates();
    });
}

function queryAllZoneRates(){
    console.log("Querying zone rates per xlsx sheet");
    for(const table of Object.values(tables)){
        //we need a conditional zone select statement based on the zones in a given xlsx table
        let zonesQuery = "";
        for(let zone of table.zones){
            zonesQuery += `SUM(IF(zone="${zone}",rate,NULL)) as "Zone ${zone}",`;
        }
        //remove trailing comma
        zonesQuery = zonesQuery.slice(0,-1);

        //We can use sql to output the exact excel table structure we need by querying for it
        const query = `SELECT start_weight, end_weight, ${zonesQuery} FROM rates WHERE client_id=${client_id} AND locale="${table.locale}" AND shipping_speed="${table.shipping_speed}" GROUP BY start_weight, end_weight;`;

        connection.query(query, function (error, results, fields) {
            if (error) throw error;
            for(row of results){
                table.data.push(Object.values(row));
            }
            shouldCreateTables();//check after every query
        });
    }
}

function shouldCreateTables(){
    let createExcel = true;
    for(const table of Object.values(tables)){
        //check if any of the tables have missing data, not ready to create an excel file yet
        if(table.data.length==0){
            createExcel=false;
        }
    }
    //all tables have data
    if(createExcel){
        createXLSX();
        connection.end();
    }
}

function createXLSX(){
    console.log('Building xlsx....')
    for(let tableName of Object.keys(tables)){
        builder.addSheet({
            name: tableName,
            columnsName: tables[tableName].columnNames,
            data: tables[tableName].data,
            dataType: 'array'
        });
    }
    builder.output(outputFilename);
    console.log(`Done! output avalialbe in ${outputFilename}`);
}

start();