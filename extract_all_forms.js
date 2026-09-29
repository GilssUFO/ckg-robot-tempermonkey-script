const fs = require('fs');
const path = require('path');

const dir = path.join(__dirname, 'Asset/CKG Sekolah');
const files = fs.readdirSync(dir).filter(f => f.startsWith('Form') || f.startsWith('1.') || f.startsWith('2.') || f.startsWith('3.'));

console.log(`Analyzing ${files.length} form files for Survey Titles and Questions:\n`);

files.forEach(f => {
    const html = fs.readFileSync(path.join(dir, f), 'utf8');
    // Extract title
    const titleMatch = html.match(/class="[^"]*(?:sd-title|sv-title)[^"]*"[^>]*aria-label="([^"]*)"/i) ||
                       html.match(/class="[^"]*(?:sd-title|sv-title)[^"]*"[^>]*>([\s\S]*?)<\//i);
    const title = titleMatch ? (titleMatch[1] || titleMatch[2]).replace(/<[^>]+>/g, '').trim() : 'Unknown';

    // Extract questions
    const qMatches = html.match(/<div[^>]*class="[^"]*(?:sd-question__title|sv-question__title)[^"]*"[^>]*>([\s\S]*?)<\/div>/gi) || [];
    const questions = qMatches.map(q => {
        return q.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    });

    console.log(`FILE: ${f}`);
    console.log(`  TITLE: ${title}`);
    console.log(`  QUESTIONS (${questions.length}):`);
    questions.forEach(q => console.log(`    * ${q}`));
    console.log('----------------------------------------------------');
});
