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

const formatDateEnd = formatDate(end);
const formatDateStart = formatDate(start);


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






connection.query(
    `SELECT qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, 
qs.info1 AS info1, qs.info2 AS info2,  qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, 
qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND 
qs.qevent = ac.event_id AND qs.datetime >= '${formatDateStart}' AND qs.datetime <= '${formatDateEnd}' 
AND q.queue IN (queue,'NONE') AND ac.event IN ('ABANDON', 'EXITWITHTIMEOUT','COMPLETECALLER','COMPLETEAGENT','AGENTLOGIN','AGENTLOGOFF','AGENTCALLBACKLOGIN','AGENTCALLBACKLOGOFF')
ORDER BY qs.datetime`,
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

connection.query(
    `SELECT qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, 
qs.info1 AS info1, qs.info2 AS info2, qs.info3 AS info3  FROM queue_stats AS qs, qname AS q, 
qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND 
qs.qevent = ac.event_id AND qs.datetime >= '${formatDateStart}' AND qs.datetime <= '${formatDateEnd}' AND 
q.queue IN (queue) AND ag.agent in (agent) AND ac.event IN ('COMPLETECALLER', 'COMPLETEAGENT') ORDER BY ag.agent`,
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



`SELECT ag.agent AS agent, qs.info1 AS info1,  qs.info2 AS info2 ";
FROM  queue_stats AS qs, qevent AS ac, qagent as ag, qname As q WHERE qs.qevent = ac.event_id 
AND qs.qname = q.qname_id AND ag.agent_id = qs.qagent AND qs.datetime >= '$start' 
AND qs.datetime <= '$end' AND  q.queue IN ($queue)  AND ag.agent in ($agent) AND  ac.event = 'TRANSFER'`


connection.query(
//     `SELECT  ac.event AS action,  qs.info1 AS info1,  qs.info2 AS info2,  qs.info3 AS info3
// FROM  queue_stats AS qs, qevent AS ac, qname As q, qagent as ag WHERE
// qs.qevent = ac.event_id AND qs.qname = q.qname_id AND qs.datetime >= '${formatDateStart}' AND
// qs.datetime <= '${formatDateEnd}' AND  q.queue IN (queue)  AND ag.agent in (agent) AND  ac.event IN ('ABANDON', 'EXITWITHTIMEOUT', 'TRANSFER')
// ORDER BY  ac.event,  qs.info3`,
    `SELECT ag.agent AS agent, qs.info1 AS info1,  qs.info2 AS info2
FROM  queue_stats AS qs, qevent AS ac, qagent as ag, qname As q WHERE qs.qevent = ac.event_id 
AND qs.qname = q.qname_id AND ag.agent_id = qs.qagent AND qs.datetime >= '${formatDateStart}' 
AND qs.datetime <= '${formatDateEnd}' AND  q.queue IN (queue)  AND ag.agent in (agent) AND  ac.event = 'TRANSFER'`,
    function (err,results) {
        console.log(results);
        fs.writeFile('./Docx/docx3.json', JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('abandoned');
            let content = JSON.parse(fs.readFileSync("./Docx/docx3.json", "utf8"));
            let readWB = xlsx.readFile('docxfromdb '+ formatName +'.xlsx');
            let newWS = xlsx.utils.json_to_sheet(content);
            xlsx.utils.book_append_sheet(readWB, newWS, 'abandoned');
            xlsx.writeFile(readWB, 'docxfromdb '+ formatName +'.xlsx');
        });
    }
);
