import React from "react";
import {Link} from "react-router-dom";

function Home() {
  return (
    <div>
      <p>
        <Link to="/pdf">pdf</Link>
      </p>
      <p>
        <Link to="/docs">docx</Link>
      </p>
      <p>
        <Link to="/xlsx">xlsx</Link>
      </p>
      <p>
        <Link to="/svg">svg</Link>
      </p>
    </div>
  );
}

export default Home;
