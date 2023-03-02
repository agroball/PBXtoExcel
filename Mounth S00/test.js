function formatDateStart(date) {
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

// добавить 23 59 59

const date = new Date();
const lastDayDate = new Date(date.getFullYear(), date.getMonth(), 0);
const firstDayDate = new Date(date.getFullYear(), date.getMonth() - 1, 1);
const mounthEnd = formatDateEnd(lastDayDate);
const mounthStart = formatDateStart(firstDayDate);
console.log(mounthStart);
console.log(mounthEnd);




