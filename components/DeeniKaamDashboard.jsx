import { useEffect, useState } from 'react';

export default function DeeniKaamDashboard() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetch('/api/deeni-data')
      .then((res) => res.json())
      .then((result) => {
        if (result.success) {
          setData(result.data);
        }
        setLoading(false);
      })
      .catch((err) => {
        console.error('Error fetching data:', err);
        setLoading(false);
      });
  }, []);

  if (loading) return <div>Loading Deeni Kaam data...</div>;

  return (
    <div>
      <h2>Deeni Kaam Dashboard (Live from BigQuery)</h2>
      <ul>
        {data.map((row, index) => (
          <li key={index}>
            {row.Source_Name} - {row.Month} ({row.Report_Value})
          </li>
        ))}
      </ul>
    </div>
  );
}
