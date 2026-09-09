'use client';
import { useState, useEffect } from 'react';

// Google Sheets ke CSV export URLs yahan direct rakh diye hain
const sheetUrls = [
  'https://docs.google.com/spreadsheets/d/1P1Ul-jXOfFfhuQLTKeQ-zOynKnCmywH-_gjZ9nJ8tO0/export?format=csv',
  'https://docs.google.com/spreadsheets/d/1S2tEIyaN8p-yu4Vd_GVumqBqwzgTzM3zlms4DwCJr00/export?format=csv',
  'https://docs.google.com/spreadsheets/d/1yWVgL9IVGrQFElLNeO8X_UGAIDSAGF7P8M31gGtoki8/export?format=csv'
];

export default function Dashboard() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    async function fetchAllSheets() {
      try {
        let allRows = [];

        for (const url of sheetUrls) {
          try {
            const response = await fetch(url);
            if (!response.ok) continue;

            const csvText = await response.text();
            if (csvText.includes('<!DOCTYPE html>') || csvText.includes('Google Accounts')) {
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
            console.error('Error fetching sheet:', err);
          }
        }

        setData(allRows);
      } catch (err) {
        setError(err.message);
      } finally {
        setLoading(false);
      }
    }

    fetchAllSheets();
  }, []);

  if (loading) {
    return (
      <div className="flex items-center justify-center min-h-screen bg-gray-950 text-white">
        <p className="text-lg">Loading Deeni Kaam data from Google Sheets...</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex items-center justify-center min-h-screen bg-gray-950 text-red-500 p-4">
        <p>Error loading data: {error}</p>
      </div>
    );
  }

  return (
    <div className="p-8 bg-gray-950 min-h-screen text-white">
      <h1 className="text-2xl font-bold mb-6">Deeni Kaam Dashboard</h1>
      <div className="overflow-x-auto">
        <table className="w-full border-collapse border border-gray-800">
          <thead>
            <tr className="bg-gray-900">
              {data.length > 0 && Object.keys(data[0]).map((key) => (
                <th key={key} className="border border-gray-800 p-2 text-left">{key}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, index) => (
              <tr key={index} className="hover:bg-gray-900">
                {Object.values(row).map((val, idx) => (
                  <td key={idx} className="border border-gray-800 p-2">{val}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
