import React from "react";

function Sfile() {
  return (
    <div>
      <svg
        width="512"
        height="512"
        viewBox="0 0 512 512"
        xmlns="http://www.w3.org/2000/svg"
      >
        <g
          fill="none"
          stroke="#4755f3"
          stroke-width="15"
          stroke-linecap="round"
          stroke-linejoin="round"
        >
          <path d="M60 200 L256 50 L452 200" />
          <path
            d="
      M60 200
      V430
      Q60 460 90 460
      H422
      Q452 460 452 430
      V200"
          />

          <path d="M60 200 L220 330" />
          <path d="M452 200 L292 330" />
        </g>

        <g
          fill="none"
          stroke="#4755f3"
          stroke-width="15"
          stroke-linecap="butt"
          stroke-linejoin="miter"
        >
          <path
            d="
    M232 210
    L165 210
    Q130 210 130 240
    L130 255"
          />
          <path
            d="
    M280 210
    L347 210
    Q382 210 382 240
    L382 255"
          />

          <path d="M220 330 L245 330" />
          <path d="M267 330 L292 330" />
        </g>
      </svg>
    </div>
  );
}

export default Sfile;
