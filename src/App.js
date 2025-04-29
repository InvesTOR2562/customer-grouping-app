import React, { useState } from "react";
import * as XLSX from "xlsx";

export default function App() {
  const [file, setFile] = useState(null);
  const [groupingResult, setGroupingResult] = useState(null);
  const [error, setError] = useState(null);

  const handleFileUpload = async (e) => {
    try {
      const uploadedFile = e.target.files[0];
      setFile(uploadedFile);

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

          const results = jsonData.map((row) => {
            const checks = {
              revenue: row.Revenue > 1000000,
              netProfit: row.NetProfit > 700000,
              deRatio: row["D/E"] < 2,
            };

            const passedGroup =
              checks.revenue && checks.netProfit && checks.deRatio;

            const group = passedGroup ? "กลุ่ม 1" : "กลุ่ม 2";

            return {
              ...row,
              group,
              checks,
            };
          });

          setGroupingResult(results);
          setError(null);
        } catch (err) {
          setError("ไม่สามารถอ่านข้อมูลจากไฟล์ Excel ได้ กรุณาตรวจสอบรูปแบบไฟล์");
        }
      };
      reader.readAsArrayBuffer(uploadedFile);
    } catch (err) {
      setError("เกิดข้อผิดพลาดในการอัปโหลดไฟล์");
    }
  };

  const downloadResult = () => {
    if (!groupingResult) return;
    const worksheet = XLSX.utils.json_to_sheet(groupingResult);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "GroupedClients");
    XLSX.writeFile(workbook, "grouping_result.xlsx");
  };

  return (
    <div style={{ padding: '2rem' }}>
      <h2>อัปโหลดไฟล์เพื่อตรวจสอบและจัดกลุ่มลูกค้า</h2>
      <input
        type="file"
        accept=".xlsx,.xls"
        onChange={handleFileUpload}
        style={{ marginBottom: '1rem' }}
      />
      {error && <p style={{ color: 'red' }}>{error}</p>}
      <button onClick={downloadResult} disabled={!groupingResult}>
        ดาวน์โหลดผลลัพธ์
      </button>

      {groupingResult && (
        <div style={{ marginTop: '2rem' }}>
          <h3>สรุปผลการจัดกลุ่ม</h3>
          <table border="1" cellPadding="5" cellSpacing="0">
            <thead>
              <tr>
                {Object.keys(groupingResult[0]).map((key) => (
                  <th key={key}>{key}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {groupingResult.map((row, index) => (
                <tr key={index}>
                  {Object.values(row).map((value, i) => (
                    <td key={i}>
                      {typeof value === "object"
                        ? JSON.stringify(value)
                        : value}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
