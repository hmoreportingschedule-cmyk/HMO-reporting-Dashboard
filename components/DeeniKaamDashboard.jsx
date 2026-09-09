import React, { useState, useEffect } from 'react';

export default function DeeniKaamDashboard() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    fetch('/api/deeni-data')
      .then((res) => res.json())
      .then((result) => {
        if (result.success) {
          setData(result.data);
        } else {
          setError(result.error);
        }
        setLoading(false);
      })
      .catch((err) => {
        console.error('Error fetching Deeni Kaam data:', err);
        setError(err.message);
        setLoading(false);
      });
  }, []);

  if (loading) {
    return (
      <div style={{ padding: '20px', textAlign: 'center', fontSize: '18px' }}>
        Loading Deeni Kaam live data from BigQuery...
      </div>
    );
  }

  if (error) {
    return (
      <div style={{ padding: '20px', color: 'red' }}>
        Error loading data: {error}
      </div>
    );
  }

  return (
    <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>
      <h2 style={{ color: '#333', marginBottom: '20px' }}>
        Deeni Kaam Dashboard (Live BigQuery Data)
      </h2>

      <div style={{ overflowX: 'auto' }}>
        <table
          style={{
            width: '100%',
            borderCollapse: 'collapse',
            backgroundColor: '#fff',
            boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
          }}
        >
          <thead>
            <tr style={{ backgroundColor: '#f4f4f4', borderBottom: '2px solid #ddd' }}>
              <th style={{ padding: '12px', textAlign: 'left' }}>Source Name</th>
              <th style={{ padding: '12px', textAlign: 'left' }}>Month</th>
              <th style={{ padding: '12px', textAlign: 'left' }}>Region</th>
              <th style={{ padding: '12px', textAlign: 'right' }}>Report Value</th>
            </tr>
          </thead>
          <tbody>
            {data.map((row, index) => (
              <tr key={index} style={{ borderBottom: '1px solid #eee' }}>
                <td style={{ padding: '10px' }}>{row.Source_Name || '-'}</td>
                <td style={{ padding: '10px' }}>{row.Month || '-'}</td>
                <td style={{ padding: '10px' }}>{row.Region || '-'}</td>
                <td style={{ padding: '10px', textAlign: 'right' }}>
                  {row.Report_Value !== undefined ? row.Report_Value : '-'}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
