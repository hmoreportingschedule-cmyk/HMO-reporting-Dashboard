import { BigQuery } from '@google-cloud/bigquery';
import { NextResponse } from 'next/server';

export async function GET() {
  try {
    const bigquery = new BigQuery({
      projectId: 'reporting-hmo',
      credentials: {
        client_email: process.env.GOOGLE_CLIENT_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n') : undefined,
      },
    });

    const query = `SELECT * FROM \`reporting-hmo.hmo_data.deeni_kaam_live\` LIMIT 100`;
    const [rows] = await bigquery.query(query);

    return NextResponse.json({ success: true, data: rows });
  } catch (error) {
    console.error('BigQuery Error:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
