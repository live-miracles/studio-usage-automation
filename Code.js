const CALENDAR_TAB = 'Calendar';

const PROJECT_SAMSKRITI = [
    'Hara Hara Mahadev',
    'Nirvana Shatakam',
    'Kalaripayattu',
    'Guru Paduka Stotram',
];
const REGIONS = ['India', 'Europe', 'APAC', 'PS'];
const PROGRAMS = ['Other', 'Step 7', '7 Day', 'Satsang', ...PROJECT_SAMSKRITI];
const PROGRAM_SESSIONS = {
    'Step 7': 5,
    '7 Day': 7,
    'Hara Hara Mahadev': 4,
    'Nirvana Shatakam': 6,
    Kalaripayattu: 9,
    'Guru Paduka Stotram': 7,
    Satsang: 1,
    Other: 1,
};
const LANGUAGES = [
    'English',
    'Hindi',
    'Tamil',
    'Telugu',
    'Kannada',
    'Marathi',
    'Malayalam',
    'Bangla',
    'Russian',
    'Spanish',
    'German',
    'Mandarin',
    'French',
    'Italian',
    'Arabic',
];

const SESSION_TYPES = {
    live: 'live',
    dryrun: 'dryrun',
    maintainance: 'maintainance',
    translation: 'translation',
};

function assertThrow(condition, message) {
    if (!condition) {
        throw new Error('Assertion failed: ' + (message || 'No message'));
    }
}

function containsKeyword(text, keyword) {
    const regex = new RegExp(`\\b${keyword}\\w*\\b`, 'i');
    return regex.test(text);
}

function containsKeywordPrefix(text, keyword, minLength = 3) {
    const key = keyword.toLowerCase();
    const words = text.toLowerCase().split(/\s+/); // split text into words
    for (const word of words) {
        if (word.length >= minLength && key.startsWith(word)) {
            return true;
        }
    }
    return false;
}

class Program {
    constructor(date, room, title, sessions) {
        this.date = date;
        this.room = room.split(' - ')[0].trim();
        this.title = title;
        this.sessions = sessions;

        for (let prog of PROGRAMS) {
            if (containsKeyword(title, prog)) {
                this.type = prog;
                break;
            }
        }
        if (!this.type) {
            this.type = 'Other';
        }

        for (let lang of LANGUAGES) {
            if (containsKeywordPrefix(title, lang)) {
                this.lang = lang;
                break;
            }
        }
        if (!this.lang) {
            this.lang = 'English';
        }

        if (containsKeywordPrefix(title, 'APAC', 4)) {
            this.region = 'APAC';
        } else if (containsKeywordPrefix(title, 'Europe', 2)) {
            this.region = 'Europe';
        } else if (this.lang === 'Spanish') {
            this.region = 'APAC';
        } else if (['Russian', 'German', 'Italian', 'French', 'Arabic'].includes(this.lang)) {
            this.region = 'Europe';
        } else if (PROJECT_SAMSKRITI.includes(this.type)) {
            this.region = 'PS';
        } else {
            this.region = 'India';
        }
    }

    static isDryRun(text) {
        return /dry[-\s]?run|test|mic check|setup/i.test(text.toLowerCase());
    }
    static isMaintainance(text) {
        return /maintenance/i.test(text.toLowerCase());
    }
    static isTranslation(text) {
        return /translation/i.test(text.toLowerCase());
    }
}

function parseTime(timeStr) {
    const match = timeStr.match(/^(\d{1,2})([:.]?(\d{1,2}))?$/);
    if (!match) return null;

    const hour = parseInt(match[1]);
    const minute = match[3] ? parseInt(match[3], 10) : 0;

    return [hour, minute];
}

function parseTimeRange(str) {
    const match = str.match(/(\d{1,2}[:.]\d{2})\s*-\s*(\d{1,2}[:.]\d{2})/);
    if (!match) return null;

    return [match[1], match[2]];
}

class Session {
    constructor(startTime, endTime, type) {
        const startParts = parseTime(startTime);
        const endParts = parseTime(endTime);
        assertThrow(startParts && endParts, 'Invalid time format: ' + startTime + ' - ' + endTime);

        this.start = startParts[0] * 60 + startParts[1];
        this.end = endParts[0] * 60 + endParts[1];
        this.type = type;
    }
    get duration() {
        let duration = this.end - this.start;
        if (duration < 0) {
            duration += 24 * 60; // Handle overnight sessions
        }
        return duration;
    }
}

function getStats(startDate, endDate) {
    const data = getFilteredDataInRange(startDate, endDate);
    const parsedData = getParsedData(data);
    const roomStats = getRoomStats(parsedData);
    const programStats = getProgramStats(parsedData);
    return {
        roomStats: roomStats,
        programStats: programStats,
    };
}

function getFilteredDataInRange(startDate, endDate) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calendar');

    const fullData = sheet.getDataRange().getValues();
    const headers = fullData[0]; // First row (headers)
    const start = new Date(startDate).setHours(0, 0, 0, 0);
    const end = new Date(endDate).setHours(0, 0, 0, 0);

    // Filter data rows (skip headers)
    const filtered = fullData.slice(1).filter((row) => {
        const cellDate = new Date(row[0]);
        if (isNaN(cellDate)) return false;
        const day = cellDate.setHours(0, 0, 0, 0);
        return day >= start && day <= end;
    });

    return [headers, ...filtered];
}

function getRoomNames() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calendar');
    const headerRow = sheet.getRange(1, 2, 1, sheet.getLastColumn() - 1).getValues()[0];
    return headerRow.map((name) => name.split(' - ')[0].trim());
}

function getParsedData(data) {
    const headers = data[0];
    const programs = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowDate = row[0];

        for (let col = 1; col < row.length; col++) {
            const cell = row[col];

            if (!cell || typeof cell !== 'string') continue;

            const cellEvents = cell.split(/^\s*(?:&)?\s*$/gm); // Paragraphs

            for (const eventText of cellEvents) {
                const lines = eventText
                    .trim()
                    .split(/\r?\n/)
                    .map((line) => line.trim());
                if (lines.length === 0) continue;

                let title = '';
                const sessions = [];

                for (const line of lines) {
                    const timeMatch = parseTimeRange(line);
                    if (!timeMatch) {
                        title = line; // First non-time line is title
                        break;
                    }
                }

                for (const line of lines) {
                    const timeRange = parseTimeRange(line);

                    if (timeRange) {
                        const isDryRun = Program.isDryRun(line) || Program.isDryRun(title);
                        const isMaintainance =
                            Program.isMaintainance(line) || Program.isMaintainance(title);
                        let type = SESSION_TYPES.live;
                        type = isDryRun ? SESSION_TYPES.dryrun : type;
                        type = isMaintainance ? SESSION_TYPES.maintainance : type;
                        sessions.push(new Session(timeRange[0], timeRange[1], type));
                    }
                }

                if (title || sessions.length > 0) {
                    programs.push(new Program(rowDate, headers[col], title, sessions));
                }
            }
        }
    }

    return programs;
}

function getRoomStats(programs) {
    const stats = {};
    getRoomNames().forEach((room) => {
        stats[room] = {
            days: new Set(),
            liveCount: 0,
            dryRunCount: 0,
            liveMin: 0,
            dryRunMin: 0,
        };
    });

    for (const prog of programs) {
        const { date, room, sessions } = prog;
        assertThrow(
            date && room && sessions,
            'Program date, room, and sessions should be defined:\n' + JSON.stringify(prog),
        );
        assertThrow(stats[room], 'Room not found ' + room);

        const roomStats = stats[room];
        roomStats.days.add(date.toString());

        for (const session of sessions) {
            if (session.type == SESSION_TYPES.dryrun) {
                roomStats.dryRunCount += 1;
                roomStats.dryRunMin += session.duration;
            } else if (
                session.type == SESSION_TYPES.live ||
                session.type == SESSION_TYPES.translation
            ) {
                roomStats.liveCount += 1;
                roomStats.liveMin += session.duration;
            }
        }
    }

    for (const room in stats) {
        stats[room].daysUsed = stats[room].days.size;
        delete stats[room].days;
    }
    console.log('Room Stats', stats);
    return stats;
}

function getProgramStats(programs) {
    const stats = {};
    REGIONS.forEach((region) => (stats[region] = {}));

    for (const prog of programs) {
        const { region, type, lang } = prog;
        assertThrow(
            region && type && lang,
            'Program region, type, and lang should be defined:\n' + JSON.stringify(prog),
        );

        // Skip if no live session

        if (!stats[region][lang]) stats[region][lang] = {};
        if (!stats[region][lang][prog.type]) stats[region][lang][type] = { count: 0, min: 0 };

        const hasLiveSession = prog.sessions.some((s) => s.type === SESSION_TYPES.live);
        if (hasLiveSession) {
            stats[region][lang][type].count++;
        }
        prog.sessions.forEach((session) => {
            stats[region][lang][type].min += session.duration;
        });
    }

    console.log('Program Stats', stats);
    return stats;
}

function getRoomReport(startDate, endDate) {
    const roomStats = getStats(startDate, endDate).roomStats;
    const rooms = Object.keys(roomStats);
    const table = [['', ...rooms]];

    const metrics = ['Days Used', 'Live Count', 'Dry Run Count', 'Live Hours', 'Dry Run Hours'];
    for (const metric of metrics) {
        const row = [metric];
        for (const room of rooms) {
            const stats = roomStats[room];
            switch (metric) {
                case 'Days Used':
                    row.push(stats.daysUsed || 0);
                    break;
                case 'Live Count':
                    row.push(stats.liveCount);
                    break;
                case 'Dry Run Count':
                    row.push(stats.dryRunCount);
                    break;
                case 'Live Hours':
                    row.push(parseInt(stats.liveMin / 60));
                    break;
                case 'Dry Run Hours':
                    row.push(parseInt(stats.dryRunMin / 60));
                    break;
            }
        }
        table.push(row);
    }

    console.log('Room Report', table);
    return table;
}

function getProgramCountReport(startDate, endDate, regionFilter = null) {
    const programStats = getStats(startDate, endDate).programStats;
    if (regionFilter) {
        Object.keys(programStats).forEach(
            (region) => region === regionFilter || delete programStats[region],
        );
    }

    // const types = PROGRAMS.filter((type) => {
    //     for (const region of Object.keys(programStats)) {
    //         for (const lang of Object.keys(programStats[region])) {
    //             if (programStats[region][lang][type]) {
    //                 return true;
    //             }
    //         }
    //     }
    //     return false;
    // });
    const types = PROGRAMS;
    const table = [['Region', 'Language', ...types]];

    Object.keys(programStats).forEach((region) => {
        const regionTable = [];
        Object.keys(programStats[region]).forEach((lang) => {
            const row = [region, lang, ...Array(types.length).fill(0)];
            Object.keys(programStats[region][lang]).forEach((type) => {
                row[table[0].indexOf(type)] = Math.round(
                    programStats[region][lang][type].count / PROGRAM_SESSIONS[type],
                );
            });
            regionTable.push(row);
        });
        regionTable.sort((a, b) => {
            const sumA = a.filter((item) => typeof item === 'number').reduce((a, b) => a + b, 0);
            const sumB = b.filter((item) => typeof item === 'number').reduce((a, b) => a + b, 0);
            return sumB - sumA;
        });
        regionTable.forEach((row) => {
            for (let i = 0; i < row.length; i++) {
                if (row[i] === 0) {
                    row[i] = '';
                }
            }
        });

        table.push(...regionTable);
    });
    if (regionFilter) {
        table.forEach((row) => row.shift());
    }

    console.log('Program Count Report', table);
    return table;
}

function getProgramHourReport(startDate, endDate) {
    const programStats = getStats(startDate, endDate).programStats;
    const table = [['Region', ...PROGRAMS]];

    Object.keys(programStats).forEach((region) => {
        const row = [region, ...Array(PROGRAMS.length).fill(0)];
        Object.keys(programStats[region]).forEach((lang) => {
            Object.keys(programStats[region][lang]).forEach((type) => {
                row[table[0].indexOf(type)] += Math.round(
                    programStats[region][lang][type].min / 60,
                );
            });
        });

        for (let i = 0; i < row.length; i++) {
            if (row[i] === 0) {
                row[i] = '';
            }
        }

        table.push(row);
    });

    console.log('Program Hour Report', table);
    return table;
}
