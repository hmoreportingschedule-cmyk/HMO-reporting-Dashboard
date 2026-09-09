'use client';
import { useState, useEffect } from 'react';

export default function Dashboard() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    async function fetchData() {
      try {
        const res = await fetch('/api/deeni-data');
        const json = await res.json();
        if (json.success) {
          setData(json.data);
        } else {
          setError(json.error);
        }
      } catch (err) {
        setError('Failed to fetch data');
      } finally {
        setLoading(false);
      }
    }
    fetchData();
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
      <div className="flex items-center justify-center min-h-screen bg-gray-950 text-red-500">
        <p>Error: {error}</p>
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
