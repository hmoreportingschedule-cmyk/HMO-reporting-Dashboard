import { BigQuery } from '@google-cloud/bigquery';
import { NextResponse } from 'next/server';

const bigquery = new BigQuery({
  projectId: 'reporting-hmo',
});

export async function GET() {
  try {
    const query = `SELECT * FROM \`reporting-hmo.hmo_data.deeni_kaam_live\` LIMIT 100`;
    const [rows] = await bigquery.query(query);

    return NextResponse.json({ success: true, data: rows });
  } catch (error) {
    console.error('BigQuery Error:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
