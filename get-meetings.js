// Fetches the published Outlook ICS feed and prints today's non-all-day
// meetings to stdout as JSON for create-devlog.ps1.
//
// Reads the feed URL from the OUTLOOK_ICS_URL environment variable (User-scope,
// set during setup; the URL is a secret and must not be committed).
//
// ical-expander resolves recurrence rules against the VTIMEZONE blocks that
// Exchange embeds in the feed, so Windows timezone names ("Eastern Standard
// Time") and DST transitions are handled without any name mapping.

const IcalExpander = require('ical-expander');

const url = process.env.OUTLOOK_ICS_URL;
if (!url) {
    console.error('OUTLOOK_ICS_URL is not set');
    process.exit(1);
}

const dayStart = new Date();
dayStart.setHours(0, 0, 0, 0);
const dayEnd = new Date(dayStart.getTime() + 24 * 60 * 60 * 1000);

function toRow(icalEvent, startDate, endDate) {
    if (startDate.isDate) return null; // all-day event
    const start = startDate.toJSDate();
    if (start < dayStart || start >= dayEnd) return null; // starts outside today

    // Exchange publishes cancelled occurrences as normal events with a
    // "Canceled:" title prefix (STATUS stays CONFIRMED).
    if (/^canceled:/i.test(icalEvent.summary || '')) return null;

    let organizer = '';
    const orgProp = icalEvent.component.getFirstProperty('organizer');
    if (orgProp) {
        organizer = orgProp.getParameter('cn')
            || String(orgProp.getFirstValue()).replace(/^mailto:/i, '');
    }

    return {
        start: start.toISOString(),
        end: endDate.toJSDate().toISOString(),
        subject: icalEvent.summary || '',
        organizer: organizer,
        location: icalEvent.location || '',
        description: icalEvent.description || ''
    };
}

async function main() {
    const res = await fetch(url);
    if (!res.ok) throw new Error('ICS fetch failed: HTTP ' + res.status);
    const ics = await res.text();

    const expander = new IcalExpander({ ics, maxIterations: 1000 });
    const { events, occurrences } = expander.between(dayStart, dayEnd);

    // A modified occurrence of a recurring meeting can show up twice: once as
    // an exception event and once as the master's expanded occurrence. Process
    // exceptions first and dedupe on UID + start time.
    const seen = new Set();
    const rows = [];
    function add(icalEvent, startDate, endDate) {
        const row = toRow(icalEvent, startDate, endDate);
        if (!row) return;
        const uid = icalEvent.component.getFirstPropertyValue('uid') || row.subject;
        const key = uid + '|' + row.start;
        if (seen.has(key)) return;
        seen.add(key);
        rows.push(row);
    }
    for (const e of events) add(e, e.startDate, e.endDate);
    for (const o of occurrences) add(o.item, o.startDate, o.endDate);

    rows.sort((a, b) => new Date(a.start) - new Date(b.start));
    process.stdout.write(JSON.stringify(rows));
}

main().catch(err => {
    console.error(err.message || String(err));
    process.exit(1);
});
