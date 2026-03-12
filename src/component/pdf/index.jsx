import React, {useState, useEffect} from "react";

export default function PdfViewer() {
  const [pdfUrl, setPdfUrl] = useState(null);

  useEffect(() => {
    return () => {
      if (pdfUrl) {
        URL.revokeObjectURL(pdfUrl);
      }
    };
  }, [pdfUrl]);

  const handleFileChange = (e) => {
    const file = e.target.files[0];

    if (!file) return;

    if (file.type !== "application/pdf") {
      alert("Please select a valid PDF file");
      return;
    }

    const objectUrl = URL.createObjectURL(file);
    setPdfUrl(objectUrl);
  };

  return (
    <div style={{height: "100vh", display: "flex", flexDirection: "column"}}>
      {/* Upload Section */}
      <div style={{padding: "10px", borderBottom: "1px solid #ddd"}}>
        <input
          type="file"
          accept="application/pdf"
          onChange={handleFileChange}
        />
      </div>

      {/* Preview Section */}
      <div style={{flex: 1}}>
        {pdfUrl ? (
          <iframe
            src={pdfUrl}
            title="PDF Preview"
            width="100%"
            height="100%"
            style={{border: "none"}}
          />
        ) : (
          <div style={{padding: "20px"}}>No PDF selected</div>
        )}
      </div>
    </div>
  );
}
