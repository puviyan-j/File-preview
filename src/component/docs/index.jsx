import React, {useRef} from "react";
import JSZip from "jszip";
import {DOMParser} from "xmldom";

export default function DocxViewer() {
  const containerRef = useRef(null);

  const handleFileChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.endsWith(".docx")) {
      alert("Please upload a .docx file");
      return;
    }

    const arrayBuffer = await file.arrayBuffer();

    const zip = await JSZip.loadAsync(arrayBuffer);

    const documentXml = await zip.file("word/document.xml").async("text");

    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(documentXml, "text/xml");

    const paragraphs = xmlDoc.getElementsByTagName("w:p");

    let html = "";

    for (let i = 0; i < paragraphs.length; i++) {
      const runs = paragraphs[i].getElementsByTagName("w:r");

      let paragraphHtml = "";

      for (let j = 0; j < runs.length; j++) {
        const run = runs[j];

        const textNode = run.getElementsByTagName("w:t")[0];
        if (!textNode) continue;

        let text = textNode.textContent;

        const bold = run.getElementsByTagName("w:b").length > 0;
        const italic = run.getElementsByTagName("w:i").length > 0;

        if (bold) text = `<strong>${text}</strong>`;
        if (italic) text = `<em>${text}</em>`;

        paragraphHtml += text;
      }

      html += `<p>${paragraphHtml}</p>`;
    }

    containerRef.current.innerHTML = html;
  };

  return (
    <div style={{height: "100vh", display: "flex", flexDirection: "column"}}>
      <div style={{padding: "10px", borderBottom: "1px solid #ddd"}}>
        <input type="file" accept=".docx" onChange={handleFileChange} />
      </div>

      <div
        ref={containerRef}
        style={{
          flex: 1,
          overflow: "auto",
          padding: "20px",
          background: "#f5f5f5",
        }}
      />
    </div>
  );
}
