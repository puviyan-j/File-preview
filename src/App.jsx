import "./App.css";
import PdfViewer from "./component/pdf";
import DocxViewer from "./component/docs";
import XlsxViewer from "./component/xlsx";

import {BrowserRouter, Routes, Route} from "react-router-dom";
import Home from "./component/home";
import Sfile from "./component/csv";

function App() {
  return (
    <>
      <BrowserRouter>
        <Routes>
          <Route path="/" element={<Home />} />
          <Route path="/pdf" element={<PdfViewer />} />
          <Route path="/docs" element={<DocxViewer />} />
          <Route path="/xlsx" element={<XlsxViewer />} />
          <Route path="/svg" element={<Sfile />} />
        </Routes>
      </BrowserRouter>
    </>
  );
}

export default App;
