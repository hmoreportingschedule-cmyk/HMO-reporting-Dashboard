import { NextResponse } from 'next/server';

export const dynamic = 'force-dynamic';

const sheetUrls = [
  'https://docs.google.com/spreadsheets/d/1P1Ul-jXOfFfhuQLTKeQ-zOynKnCmywH-_gjZ9nJ8tO0/export?format=csv',
  'https://docs.google.com/spreadsheets/d/1S2tEIyaN8p-yu4Vd_GVumqBqwzgTzM3zlms4DwCJr00/export?format=csv',
  'https://docs.google.com/spreadsheets/d/1yWVgL9IVGrQFElLNeO8X_UGAIDSAGF7P8M31gGtoki8/export?format=csv'
];

export async function GET() {
  try {
    let allRows = [];

    for (const url of sheetUrls) {
      try {
        const response = await fetch(url, { cache: 'no-store' });
        if (!response.ok) continue;
        
        const csvText = await response.text();
        
        // Agar Google ka login page ya error HTML aa jaye toh skip kar dein
        if (csvText.includes('<!DOCTYPE html>') || csvText.includes('Google Accounts')) {
          console.error('Sheet is private or requires authentication:', url);
          continue;
        }

        const lines = csvText.split('\n').filter(line => line.trim() !== '');
        if (lines.length === 0) continue;
        
        const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
        
        for (let i = 1; i < lines.length; i++) {
          const currentLine = lines[i].split(',');
          let rowObj = {};
          for (let j = 0; j < headers.length; j++) {
            rowObj[headers[j]] = currentLine[j] ? currentLine[j].trim().replace(/^"|"$/g, '') : '';
          }
          allRows.push(rowObj);
        }
      } catch (err) {
        console.error('Error parsing sheet:', err);
      }
    }

    return NextResponse.json({ success: true, data: allRows });
  } catch (error) {
    console.error('API Crash Error:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
