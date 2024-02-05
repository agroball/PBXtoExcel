const moment = require('moment');

function formatDateStart(date) {
    let d = new Date(date),
        month = '' + d.getMonth(),
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


function formatDateEnd(date) {
    let d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear(),
        hour = '23',
        minutes = '59',
        seconds = '59';

    if (month.length < 2)
        month = '0' + month;
    if (day.length < 2)
        day = '0' + day;

    return [year, month, day].join('-') + ' ' + [ hour, minutes, seconds].join(':');
}
const date = new Date();
const firstDayDate = new Date(date.getFullYear(), date.getMonth() -1, 0);
const lastDayDate = new Date(date.getFullYear(), date.getMonth(), 0);
//const mounthStart = moment.utc(formatDateStart(firstDayDate));
//const mounthEnd = formatDateEnd(lastDayDate);
const mounthStart = '2023-08-01 00:00:00';
const mounthEnd = '2023-08-01 23:59:59';
//console.log(mounthStart);
//console.log(mounthEnd);

//const dayfirst = moment.utc(firstDayDate, 'YYYY:MM:DD hh:mm:ss');
// const dayfirst = moment.utc(moment().year(date.getFullYear()).month(date.getMonth() - 1).date(1).hour(0).minutes(0).seconds(0).format('YYYY-MM-DD hh:mm:ss'));
// const daylast = moment.utc().year(date.getFullYear()).month(date.getMonth()).date(0).hour(23).minutes(59).seconds(59).format('YYYY-MM-DD hh:mm:ss');

const dayfirst = moment(mounthStart).utc();
const daylast = moment(mounthEnd).utc();
console.log(dayfirst);
console.log(daylast);
