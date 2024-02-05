const mysql = require("mysql2");
const fs = require("fs");
const xlsx = require('xlsx');

    const connection = mysql.createConnection({
        host: "",
        user: "Roman",
        database: "",
        password: ""
    });
    connection.connect(function(err){
        if (err) {
            return console.error("Ошибка: " + err.message);
        }
        else{
            console.log("Подключение к серверу MySQL успешно установлено");
        }
    });

    const d = new Date();
    const start = d.setDate(d.getDate() - 1);
    const ed = new Date();
    const end = ed.setDate(ed.getDate());


    function formatDate(date) {
        let d = new Date(date),
            month = '' + (d.getMonth() + 1),
            day = '' + d.getDate(),
            year = d.getFullYear(),
            hour = '00',
            minutes = '00',
            seconds = '00';

        if (month.length < 2)
            month = '0' + month;
        if (day.length < 2)
            day = '0' + day;

        return [year, month, day].join('-') + ' ' + [ hour, minutes, seconds].join(':');
    }

    //const formatDateEnd = formatDate(end);
    //const formatDateStart = formatDate(start);
    const formatDateEnd = '2023-02-05 00:00:00';
    const formatDateStart = '2023-02-05 23:59:59';

    console.log(formatDateEnd);
    console.log(formatDateStart);


const letName = ed.setDate(ed.getDate() - 1);

function formatDateName(date) {
    let d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2)
        month = '0' + month;
    if (day.length < 2)
        day = '0' + day;

    return [year, month, day].join('-');
}

const formatName = formatDateName(letName);

//*2323
    connection.query(
        `select count(*) from cdr where calldate > '${formatDateStart}' and calldate < '${formatDateEnd}' and src in ('74952491625') and disposition = 'ANSWERED' and dst = '9995' ORDER BY calldate DESC`,
        function(err, results) {
            console.log(results);
            fs.writeFile('./Docx/docx1.json', JSON.stringify(results), function (err) {
                if(err) {
                    return console.log(err);
                }
                console.log('APD CC');
                let content = JSON.parse(fs.readFileSync("./Docx/docx1.json", "utf8"));
                let newWB = xlsx.utils.book_new();
                let newWS = xlsx.utils.json_to_sheet(content);
                xlsx.utils.book_append_sheet(newWB, newWS, 'APD CC');
                xlsx.writeFile(newWB, 'docxfromdb '+ formatName +'.xlsx');
            });
        }
    );

//GIBDD
    const numbersgibdd = JSON.parse(fs.readFileSync('./Numbers/numbersgibdd.json', "utf8"));
    const elementNUmberGibdd = numbersgibdd.map(element => element.cidnum);
    connection.query(
        `select count(*) from cdr where calldate > '${formatDateStart}' and calldate < '${formatDateEnd}' and src in (${elementNUmberGibdd}) and disposition = 'ANSWERED' and dst = '9995' ORDER BY calldate DESC`,
        function (err,results) {
            console.log(results);
            fs.writeFile('./Docx/docx2.json', JSON.stringify(results), function (err) {
                if (err) {
                    return console.log(err);
                }
                console.log('GIBDD');
                let content = JSON.parse(fs.readFileSync("./Docx/docx2.json", "utf8"));
                let readWB = xlsx.readFile('docxfromdb '+ formatName +'.xlsx');
                let newWS = xlsx.utils.json_to_sheet(content);
                xlsx.utils.book_append_sheet(readWB, newWS, 'GIBDD');
                xlsx.writeFile(readWB, 'docxfromdb '+ formatName +'.xlsx');
            });
        }
    );


//CPU330
    connection.query(
        `select count(*) from cdr where calldate > '${formatDateStart}' and calldate < '${formatDateEnd}' and src in ('74822360279') and disposition = 'ANSWERED' and dst = '9995' ORDER BY calldate DESC`,
        function (err,results) {
            console.log(results);
            fs.writeFile('./Docx/docx3.json', JSON.stringify(results), function (err) {
                if (err) {
                    return console.log(err);
                }
                console.log('CPU330');
                let content = JSON.parse(fs.readFileSync("./Docx/docx3.json", "utf8"));
                let readWB = xlsx.readFile('docxfromdb '+ formatName +'.xlsx');
                let newWS = xlsx.utils.json_to_sheet(content);
                xlsx.utils.book_append_sheet(readWB, newWS, 'CPU330');
                xlsx.writeFile(readWB, 'docxfromdb '+ formatName +'.xlsx');
            });
        }
    );


//MTTS
const numbersmtts = JSON.parse(fs.readFileSync('./Numbers/numbersmtts.json', "utf8"));
const elementNUmberMtts = numbersmtts.map(element => element.cidnum);
connection.query(
    `select count(*) from cdr where calldate > '${formatDateStart}' and calldate < '${formatDateEnd}' and src in (${elementNUmberMtts}) and disposition = 'ANSWERED' and dst = '9995' ORDER BY calldate DESC`,
    function (err,results) {
        console.log(results);
        fs.writeFile('./Docx/docx4.json', JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('MTTS');
            let content = JSON.parse(fs.readFileSync("./Docx/docx4.json", "utf8"));
            let readWB = xlsx.readFile('docxfromdb '+ formatName +'.xlsx');
            let newWS = xlsx.utils.json_to_sheet(content);
            xlsx.utils.book_append_sheet(readWB, newWS, 'MTTS');
            xlsx.writeFile(readWB, 'docxfromdb '+ formatName +'.xlsx');
        });
    }
);




//avtodor
connection.query(
    `select count(*) from cdr where calldate > '${formatDateStart}' and calldate < '${formatDateEnd}' and src in ('74952499390') and disposition = 'ANSWERED' and dst = '9995' ORDER BY calldate DESC`,
    function(err, results) {
        console.log(results);
        fs.writeFile('./Docx/docx5.json', JSON.stringify(results), function (err) {
            if(err) {
                return console.log(err);
            }
            console.log('AVTRTMSO');
            let content = JSON.parse(fs.readFileSync("./Docx/docx5.json", "utf8"));
            let readWB = xlsx.readFile('docxfromdb '+ formatName +'.xlsx');
            let newWS = xlsx.utils.json_to_sheet(content);
            xlsx.utils.book_append_sheet(readWB, newWS, 'AVTRTMSO');
            xlsx.writeFile(readWB, 'docxfromdb '+ formatName +'.xlsx');
        });
    }
);

//RAS
const numbersras = JSON.parse(fs.readFileSync('./Numbers/numbersras.json', "utf8"));
const elementNUmberRas = numbersras.map(element => element.cidnum);
connection.query(
    `select count(*) from cdr where calldate > '${formatDateStart}' and calldate < '${formatDateEnd}' and src in (${elementNUmberRas}) and disposition = 'ANSWERED' and dst = '9995' ORDER BY calldate DESC`,
    function(err, results) {
        console.log(results);
        fs.writeFile('./Docx/docx6.json', JSON.stringify(results), function (err) {
            if(err) {
                return console.log(err);
            }
            console.log('RAS');
            let content = JSON.parse(fs.readFileSync("./Docx/docx6.json", "utf8"));
            let readWB = xlsx.readFile('docxfromdb '+ formatName +'.xlsx');
            let newWS = xlsx.utils.json_to_sheet(content);
            xlsx.utils.book_append_sheet(readWB, newWS, 'RAS');
            xlsx.writeFile(readWB, 'docxfromdb '+ formatName +'.xlsx');
        });
    }
);


//TSI
const numberstsi = JSON.parse(fs.readFileSync('./Numbers/numbertsi.json', "utf8"));
const elementNUmberTsi = numberstsi.map(element => element.cidnum);
connection.query(
    `select count(*) from cdr where calldate > '${formatDateStart}' and calldate < '${formatDateEnd}' and src in (${elementNUmberTsi}) and disposition = 'ANSWERED' and dst = '9995' ORDER BY calldate DESC`,
    function(err, results) {
        console.log(results);
        fs.writeFile('./Docx/docx7.json', JSON.stringify(results), function (err) {
            if(err) {
                return console.log(err);
            }
            console.log('TSI');
            let content = JSON.parse(fs.readFileSync("./Docx/docx7.json", "utf8"));
            let readWB = xlsx.readFile('docxfromdb '+ formatName +'.xlsx');
            let newWS = xlsx.utils.json_to_sheet(content);
            xlsx.utils.book_append_sheet(readWB, newWS, 'TSI');
            xlsx.writeFile(readWB, 'docxfromdb '+ formatName +'.xlsx');
        });
    }
);


//Evacuation
const numberseva = JSON.parse(fs.readFileSync('./Numbers/numberevacuation.json', "utf8"));
const elementNUmberEva = numberseva.map(element => element.cidnum);
connection.query(
    `select count(*) from cdr where calldate > '${formatDateStart}' and calldate < '${formatDateEnd}' and src in (${elementNUmberEva}) and disposition = 'ANSWERED' and dst = '9995' ORDER BY calldate DESC`,
    function(err, results) {
        console.log(results);
        fs.writeFile('./Docx/docx8.json', JSON.stringify(results), function (err) {
            if(err) {
                return console.log(err);
            }
            console.log('Evacuation');
            let content = JSON.parse(fs.readFileSync("./Docx/docx8.json", "utf8"));
            let readWB = xlsx.readFile('docxfromdb '+ formatName +'.xlsx');
            let newWS = xlsx.utils.json_to_sheet(content);
            xlsx.utils.book_append_sheet(readWB, newWS, 'Evacuation');
            xlsx.writeFile(readWB, 'docxfromdb '+ formatName +'.xlsx');
        });
    }
);



connection.query(
    `select count(*) from cdr where calldate > '${formatDateStart}' and calldate < '${formatDateEnd}'  and disposition = 'ANSWERED' and dst = '9995' ORDER BY calldate DESC`,
    function(err, results) {
        console.log(results);
    }
);


    connection.end();


