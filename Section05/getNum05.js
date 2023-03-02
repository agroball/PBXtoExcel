const mysql = require("mysql2");
const fs = require("fs");
const xlsx = require("xlsx");

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

//ras
connection.query(
    `SELECT cidnum FROM incoming WHERE description LIKE '%Ð”Ð Ðœ%'`,
    function (err,results) {
        console.log(results);
        fs.writeFile('./Numbers/numbersras.json', JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('ras numbers rewrite');
        });
    }
);

//tsi
connection.query(
    `SELECT cidnum FROM incoming WHERE description LIKE '%ÐÐš%'`,
    function (err,results) {
        console.log(results);
        fs.writeFile('./Numbers/numbertsi.json', JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('tsi numbers rewrite');
        });
    }
);


connection.query(
    `SELECT cidnum FROM incoming WHERE description LIKE '%Ð¡Ð¾Ñ‚Ð¾Ð²Ñ‹Ð¹ ÑÐ²Ð°ÐºÑƒÐ°Ñ‚Ð¾Ñ€%'`,
    function (err,results) {
        console.log(results);
        fs.writeFile('./Numbers/numberevacuation.json', JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('evacuation numbers rewrite');
        });
    }
);
//gibdd
connection.query(
    `SELECT cidnum FROM incoming WHERE description LIKE '%GIBDD%'`,
    function (err,results) {
        console.log(results);
        fs.writeFile('./Numbers/numbersgibdd.json', JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('gibdd numbers rewrite');
        });
    }
);
//mtts
connection.query(
    `SELECT cidnum FROM incoming WHERE description LIKE '%MTTS%'`,
    function (err,results) {
        console.log(results);
        fs.writeFile('./Numbers/numbersmtts.json', JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('mtts numbers rewrite');
        });
    }
);

connection.end();
