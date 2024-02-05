const mysql = require("mysql2");
const fs = require("fs");
const xlsx = require('xlsx');
const loadash = require("lodash");
const moment = require("moment");


const connection = mysql.createConnection({
    host: "",
    user: "",
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

function formatDateStart(date) {
    let d = new Date(date),
        month = '' + (d.getMonth()),
        day = '' + (d.getDate()),
        year = d.getFullYear(),
        hour = '21',
        minutes = '00',
        seconds = '00';

    if (month.length < 2)
        month = '0' + month;
    if (day.length < 2)
        day = '0' + day;

    return [year, month, day].join('-') + ' ' + [ hour, minutes, seconds].join(':');
}


function formatDateEnd(date) {
    let d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear(),
        hour = '20',
        minutes = '59',
        seconds = '59';

    if (month.length < 2)
        month = '0' + month;
    if (day.length < 2)
        day = '0' + day;

    return [year, month, day].join('-') + ' ' + [ hour, minutes, seconds].join(':');
}
const date = new Date();
const lastDayDate = new Date(date.getFullYear(), date.getMonth(), 0);
const firstDayDate = new Date(date.getFullYear(), date.getMonth(), 0);
//const mounthStart = formatDateStart(firstDayDate);
//const mounthEnd = formatDateEnd(lastDayDate);
const mounthStart = '2023-08-01 00:00:00';
const mounthEnd = '2023-08-01 23:59:59';
// console.log(mounthStart);
// console.log(mounthEnd);

const dayfirst = moment(mounthStart).add(+3, 'hours').format('YYYY-MM-DD hh:mm:ss');
const daylast = moment(mounthEnd).add(+3, 'hours').format('YYYY-MM-DD hh:mm:ss');

console.log(dayfirst);
console.log(daylast);

const letName = date.setDate(date.getDate() - 1);

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

function formatDateArr(date) {
    let d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2)
        month = '0' + month;
    if (day.length < 2)
        day = '0' + day;

    return [day, month, year].join('.');
}

function formatHourArr(date) {
    let d = new Date(date),
        hour = '' + d.getHours(),
        minutes = '00',
        seconds = '00';

    return [hour, minutes, seconds].join(':');

}

//*2323
                connection.query(
                    `select count(*) as count from cdr where calldate > '${mounthStart}' and calldate < '${mounthEnd}' and src in ('84952491625', '84952499390') and disposition = 'ANSWERED' and dst in ('2004', '2001', '2002', '2003', '2005', '2011', '2012', '2013', '2014', '2022', '2024') ORDER BY calldate DESC`,
                    function (err, results) {
                        console.log(results);
                        fs.writeFile(`${__dirname}/Docx/docx1.json`, JSON.stringify(results), function (err) {
                            if (err) {
                                return console.log(err);
                            }
                            console.log('APD CC');
                        });
                    }
                );

//GIBDD
                const numbersgibdd = JSON.parse(fs.readFileSync(`${__dirname}/Numbers/numbersgibdd.json`, "utf8"));
                const elementNUmberGibdd = numbersgibdd.map(element => element.cidnum);
                connection.query(
                    `select count(*) as count from cdr where calldate > '${mounthStart}' and calldate < '${mounthEnd}' and src in (${elementNUmberGibdd}) and disposition = 'ANSWERED' and dst in ('2004', '2001', '2002', '2003', '2005', '2011', '2012', '2013', '2014', '2022', '2024') ORDER BY calldate DESC`,
                    function (err, results) {
                        console.log(results);
                        fs.writeFile(`${__dirname}/Docx/docx2.json`, JSON.stringify(results), function (err) {
                            if (err) {
                                return console.log(err);
                            }
                            console.log('GIBDD');
                        });
                    }
                );

//MTTS
                const numbersmtts = JSON.parse(fs.readFileSync(`${__dirname}/Numbers/numbersmtts.json`, "utf8"));
                const elementNUmberMtts = numbersmtts.map(element => element.cidnum);
                connection.query(
                    `select count(*) as count from cdr where calldate > '${mounthStart}' and calldate < '${mounthEnd}' and src in (${elementNUmberMtts}) and disposition = 'ANSWERED' and dst in ('2004', '2001', '2002', '2003', '2005', '2011', '2012', '2013', '2014', '2022', '2024') ORDER BY calldate DESC`,
                    function (err, results) {
                        console.log(results);
                        fs.writeFile(`${__dirname}/Docx/docx4.json`, JSON.stringify(results), function (err) {
                            if (err) {
                                return console.log(err);
                            }
                            console.log('MTTS');
                        });
                    }
                );
//RAS
                const numbersras = JSON.parse(fs.readFileSync(`${__dirname}/Numbers/numbersras.json`, "utf8"));
                const elementNUmberRas = numbersras.map(element => element.cidnum);
                connection.query(
                    `select count(*) as count from cdr where calldate > '${mounthStart}' and calldate < '${mounthEnd}' and src in (${elementNUmberRas}) and disposition = 'ANSWERED' and dst in ('2004', '2001', '2002', '2003', '2005', '2011', '2012', '2013', '2014', '2022', '2024') ORDER BY calldate DESC`,
                    function (err, results) {
                        console.log(results);
                        fs.writeFile(`${__dirname}/Docx/docx6.json`, JSON.stringify(results), function (err) {
                            if (err) {
                                return console.log(err);
                            }
                            console.log('RAS');
                        });
                    }
                );


//TSI
                const numberstsi = JSON.parse(fs.readFileSync(`${__dirname}/Numbers/numbertsi.json`, "utf8"));
                const elementNUmberTsi = numberstsi.map(element => element.cidnum);
                connection.query(
                    `select count(*) as count from cdr where calldate > '${mounthStart}' and calldate < '${mounthEnd}' and src in (${elementNUmberTsi}) and disposition = 'ANSWERED' and dst in ('2004', '2001', '2002', '2003', '2005', '2011', '2012', '2013', '2014', '2022', '2024') ORDER BY calldate DESC`,
                    function (err, results) {
                        console.log(results);
                        fs.writeFile(`${__dirname}/Docx/docx7.json`, JSON.stringify(results), function (err) {
                            if (err) {
                                return console.log(err);
                            }
                            console.log('TSI');
                        });
                    }
                );


//Evacuation
                const numberseva = JSON.parse(fs.readFileSync(`${__dirname}/Numbers/numberevacuation.json`, "utf8"));
                const elementNUmberEva = numberseva.map(element => element.cidnum);
                connection.query(
                    `select count(*) as count from cdr where calldate > '${mounthStart}' and calldate < '${mounthEnd}' and src in (${elementNUmberEva}) and disposition = 'ANSWERED' and dst in ('2004', '2001', '2002', '2003', '2005', '2011', '2012', '2013', '2014', '2022', '2024') ORDER BY calldate DESC`,
                    function (err, results) {
                        console.log(results);
                        fs.writeFile(`${__dirname}/Docx/docx8.json`, JSON.stringify(results), function (err) {
                            if (err) {
                                return console.log(err);
                            }
                            console.log('Evacuation');
                        });
                    }
                );
connection.end();




const connectionQueue = mysql.createConnection({
    host: "",
    user: "",
    database: "",
    password: ""
});
    connectionQueue.connect(function (err) {
        if (err) {
            return console.error("Ошибка: " + err.message);
        } else {
            console.log("Подключение к серверу MySQL успешно установлено");
        }
    });

    connectionQueue.query(
        `SELECT count(*) AS count, qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, qs.info1 AS info1, qs.info2 AS info2,  qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND qs.qevent = ac.event_id AND qs.datetime >= '${dayfirst}' AND qs.datetime <= '${daylast}' AND q.queue IN ('9980','NONE') AND ac.event IN ('COMPLETECALLER','COMPLETEAGENT','AGENTLOGIN','AGENTLOGOFF','AGENTCALLBACKLOGIN','AGENTCALLBACKLOGOFF', 'TRANSFER') GROUP BY DATE(qs.datetime)`,
        function (err, results) {
            console.log(results);
            fs.writeFile(`${__dirname}/Docx/docx9.json`, JSON.stringify(results), function (err) {
                if (err) {
                    return console.log(err);
                }
                console.log('COUNTBYDAY');
            });
        }
    );

    connectionQueue.query(
        `SELECT count(*) AS count, qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, qs.info1 AS info1, qs.info2 AS info2,  qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND qs.qevent = ac.event_id AND qs.datetime >= '${mounthStart}' AND qs.datetime <= '${mounthEnd}' AND q.queue IN ('9980','NONE') AND ac.event IN ('ABANDON', 'EXITWITHTIMEOUT', 'TRANSFER') GROUP BY DATE(qs.datetime)`,
        function (err, results) {
            console.log(results);
            fs.writeFile(`${__dirname}/Docx/docx10.json`, JSON.stringify(results), function (err) {
                if (err) {
                    return console.log(err);
                }
                console.log('LOSTBYDAYS');
            });
        }
    );

connectionQueue.query(
    `SELECT count(*) AS count, qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, qs.info1 AS info1, qs.info2 AS info2,  qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND qs.qevent = ac.event_id AND qs.datetime >= '${mounthStart}' AND qs.datetime <= '${mounthEnd}' AND q.queue IN ('9980','NONE') AND ac.event IN ('COMPLETECALLER','COMPLETEAGENT','AGENTLOGIN','AGENTLOGOFF','AGENTCALLBACKLOGIN','AGENTCALLBACKLOGOFF', 'TRANSFER') GROUP BY hour( datetime )`,
    function (err, results) {
        console.log(results);
        fs.writeFile(`${__dirname}/Docx/docx11.json`, JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('COUNTBYHOUR');
        });
    }
);
connectionQueue.query(
    `SELECT count(*) AS count, qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, qs.info1 AS info1, qs.info2 AS info2,  qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND qs.qevent = ac.event_id AND qs.datetime >= '${mounthStart}' AND qs.datetime <= '${mounthEnd}' AND q.queue IN ('9980','NONE') AND ac.event IN ('ABANDON', 'EXITWITHTIMEOUT', 'TRANSFER') GROUP BY hour( datetime )`,
    function (err, results) {
        console.log(results);
        fs.writeFile(`${__dirname}/Docx/docx12.json`, JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('LOSTYHOUR');
        });
    }
);
connectionQueue.query(
    `SELECT count(*) AS count, qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, qs.info1 AS info1, qs.info2 AS info2, qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND qs.qevent = ac.event_id AND qs.datetime >= '${mounthStart}' AND qs.datetime <= '${mounthEnd}' AND q.queue IN ('9980', 'NONE') AND ag.agent in ('TMSO CPU', 'TMSO 1', 'TMSO 2', 'TMSO 3', 'TMSO') AND ac.event IN ('COMPLETECALLER', 'COMPLETEAGENT') GROUP BY hour( datetime ), ag.agent`,
    function (err, results) {
        console.log(results);
        fs.writeFile(`${__dirname}/Docx/docx13.json`, JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('byagentByhour');
        });
    }
);
connectionQueue.query(
    `SELECT count(*) AS count, qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, qs.info1 AS info1, qs.info2 AS info2, qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND qs.qevent = ac.event_id AND qs.datetime >= '${mounthStart}' AND qs.datetime <= '${mounthEnd}' AND q.queue IN ('9980', 'NONE') AND ag.agent in ('TMSO') AND ac.event IN ('COMPLETECALLER', 'COMPLETEAGENT') GROUP BY hour( datetime ), ag.agent`,
    function (err, results) {
        console.log(results);
        fs.writeFile(`${__dirname}/Docx/docx91.json`, JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('ByTMSO');
        });
    }
);
connectionQueue.query(
    `SELECT count(*) AS count, qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, qs.info1 AS info1, qs.info2 AS info2, qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND qs.qevent = ac.event_id AND qs.datetime >= '${mounthStart}' AND qs.datetime <= '${mounthEnd}' AND q.queue IN ('9980', 'NONE') AND ag.agent in ('TMSO 1') AND ac.event IN ('COMPLETECALLER', 'COMPLETEAGENT') GROUP BY hour( datetime ), ag.agent`,
    function (err, results) {
        console.log(results);
        fs.writeFile(`${__dirname}/Docx/docx92.json`, JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('ByTMSO1');
        });
    }
);
connectionQueue.query(
    `SELECT count(*) AS count, qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, qs.info1 AS info1, qs.info2 AS info2, qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND qs.qevent = ac.event_id AND qs.datetime >= '${mounthStart}' AND qs.datetime <= '${mounthEnd}' AND q.queue IN ('9980', 'NONE') AND ag.agent in ('TMSO 2') AND ac.event IN ('COMPLETECALLER', 'COMPLETEAGENT') GROUP BY hour( datetime ), ag.agent`,
    function (err, results) {
        console.log(results);
        fs.writeFile(`${__dirname}/Docx/docx93.json`, JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('ByTMSO2');
        });
    }
);
connectionQueue.query(
    `SELECT count(*) AS count, qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, qs.info1 AS info1, qs.info2 AS info2, qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND qs.qevent = ac.event_id AND qs.datetime >= '${mounthStart}' AND qs.datetime <= '${mounthEnd}' AND q.queue IN ('9980', 'NONE') AND ag.agent in ('TMSO 3') AND ac.event IN ('COMPLETECALLER', 'COMPLETEAGENT') GROUP BY hour( datetime ), ag.agent`,
    function (err, results) {
        console.log(results);
        fs.writeFile(`${__dirname}/Docx/docx94.json`, JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('ByTMSO3');
        });
    }
);
connectionQueue.query(
    `SELECT count(*) AS count, qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, qs.info1 AS info1, qs.info2 AS info2, qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND qs.qevent = ac.event_id AND qs.datetime >= '${mounthStart}' AND qs.datetime <= '${mounthEnd}' AND q.queue IN ('9980', 'NONE') AND ag.agent in ('TMSO CPU') AND ac.event IN ('COMPLETECALLER', 'COMPLETEAGENT') GROUP BY hour( datetime ), ag.agent`,
    function (err, results) {
        console.log(results);
        fs.writeFile(`${__dirname}/Docx/docx95.json`, JSON.stringify(results), function (err) {
            if (err) {
                return console.log(err);
            }
            console.log('ByTMSOCPU');
        });
    }
);
connectionQueue.end();

function Queue() {
    setTimeout(() => {

        let cdr1 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx7.json`, "utf8"));
        let cdrarr1 = cdr1.map(i => ({days: formatDateArr(i.calldate), tsi: i.count}));

        let cdr2 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx2.json`, "utf8"));
        let cdrarr2 = cdr2.map(i => ({days: formatDateArr(i.calldate), gibdd: i.count}));

        let cdr3 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx4.json`, "utf8"));
        let cdrarr3 = cdr3.map(i => ({days: formatDateArr(i.calldate), mtts: i.count}));

        let cdr4 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx6.json`, "utf8"));
        let cdrarr4 = cdr4.map(i => ({days: formatDateArr(i.calldate), ras: i.count}));

        let cdr5 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx1.json`, "utf8"));
        let cdrarr5 = cdr5.map(i => ({days: formatDateArr(i.calldate), apdcc: i.count}));

        let cdr6 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx8.json`, "utf8"));
        let cdrarr6 = cdr6.map(i => ({days: formatDateArr(i.calldate), evacuation: i.count}));

        const cdrresult1 = cdrarr1.map(item => {
            const itemWithSameDay = cdrarr2.find(({days, count}) => days === item.days)
            if(itemWithSameDay){
                item.gibdd = itemWithSameDay.gibdd
            }
            return item
        })

        const cdrresult2 = cdrresult1.map(item => {
            const itemWithSameDay = cdrarr3.find(({days, count}) => days === item.days)
            if(itemWithSameDay){
                item.mtts = itemWithSameDay.mtts
            }
            return item
        })

        const cdrresult3 = cdrresult2.map(item => {
            const itemWithSameDay = cdrarr4.find(({days, count}) => days === item.days)
            if(itemWithSameDay){
                item.ras = itemWithSameDay.ras
            }
            return item
        })

        const cdrresult4 = cdrresult3.map(item => {
            const itemWithSameDay = cdrarr5.find(({days, count}) => days === item.days)
            if(itemWithSameDay){
                item.apdcc = itemWithSameDay.apdcc
            }
            return item
        })

        const cdrresult5 = cdrresult4.map(item => {
            const itemWithSameDay = cdrarr6.find(({days, count}) => days === item.days)
            if(itemWithSameDay){
                item.evacuation = itemWithSameDay.evacuation
            }
            return item
        })

        let newWB = xlsx.utils.book_new();
        let newWS = xlsx.utils.json_to_sheet(cdrresult5);
        xlsx.utils.book_append_sheet(newWB, newWS, 'WHO');
        xlsx.writeFile(newWB, `${__dirname}/docxfromdb ${formatName}.xlsx`);

        //////////////////////////

        let content1 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx9.json`, "utf8"));
        let arr1 = content1.map(i => ({days: formatDateArr(i.datetime), cnt: i.count}));

        let content2 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx10.json`, "utf8"));
        let arr2 = content2.map(i => ({days: formatDateArr(i.datetime), count: i.count}));

        let content3 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx11.json`, "utf8"));
        let newarr3 = content3.map(i => ({hour: formatHourArr(i.datetime), cnt: i.count}));

        let content4 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx12.json`, "utf8"));
        let newarr4 = content4.map(i => ({hour: formatHourArr(i.datetime), count: i.count}));

        let content5 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx91.json`, "utf8"));
        let newarr5 = content5.map(i => ({hour: formatHourArr(i.datetime), c: i.count}));

        let content6 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx92.json`, "utf8"));
        let newarr6 = content6.map(i => ({hour: formatHourArr(i.datetime), co: i.count}));

        let content7 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx93.json`, "utf8"));
        let newarr7 = content7.map(i => ({hour: formatHourArr(i.datetime), cou: i.count}));

        let content8 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx94.json`, "utf8"));
        let newarr8 = content8.map(i => ({hour: formatHourArr(i.datetime), coun: i.count}));

        let content9 = JSON.parse(fs.readFileSync(`${__dirname}/Docx/docx95.json`, "utf8"));
        let newarr9 = content9.map(i => ({hour: formatHourArr(i.datetime), count: i.count}));

        const result = arr1.map(item => {
            const itemWithSameDay = arr2.find(({days, count}) => days === item.days)
            if (itemWithSameDay) {
                item.count = itemWithSameDay.count
            }
            return item
        })

        const result2 = newarr3.map(item => {
            const itemWithSameHour = newarr4.find(({hour, count}) => hour === item.hour)
            if(itemWithSameHour){
                item.count = itemWithSameHour.count
            }
            return item
        })

        const result3 = newarr5.map(item => {
            const itemWithSameHour = newarr6.find(({hour, count}) => hour === item.hour)
            if(itemWithSameHour){
                item.co = itemWithSameHour.co
            }
            return item
        })

        const result4 = result3.map(item => {
            const itemWithSameHour = newarr7.find(({hour, count}) => hour === item.hour)
            if(itemWithSameHour){
                item.cou = itemWithSameHour.cou
            }
            return item
        })

        const result5 = result4.map(item => {
            const itemWithSameHour = newarr8.find(({hour, count}) => hour === item.hour)
            if(itemWithSameHour){
                item.coun = itemWithSameHour.coun
            }
            return item
        })

        const result6 = result5.map(item => {
            const itemWithSameHour = newarr9.find(({hour, count}) => hour === item.hour)
            if(itemWithSameHour){
                item.count = itemWithSameHour.count
            }
            return item
        })

        let readWB = xlsx.readFile(`${__dirname}/docxfromdb ${formatName}.xlsx`);
        let newWss = xlsx.utils.json_to_sheet(result);
        xlsx.utils.book_append_sheet(readWB, newWss, 'ByDays');
        let newWs = xlsx.utils.json_to_sheet(result2);
        xlsx.utils.book_append_sheet(readWB, newWs, 'BYHOUR');
        let newW3 = xlsx.utils.json_to_sheet(result6);
        xlsx.utils.book_append_sheet(readWB, newW3, 'BYAGENT');
        xlsx.writeFile(readWB, `${__dirname}/docxfromdb ${formatName}.xlsx`);
    }, 50000)

}
Queue();
