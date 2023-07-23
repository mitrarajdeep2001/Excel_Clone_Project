// Default cell properties
const defaultProperties = {
  text: "",
  "font-weight": "",
  "font-style": "",
  "text-decoration": "",
  "text-align": "left",
  "background-color": "white",
  color: "black",
  "font-family": "Arial",
  "font-size": 14,
};

const cellData = {
  sheet1: {},
};

const selectedSheet = "sheet1";
const totalSheets = 1;

// Create Column Name and Row Number
(function () {
  const column = document.querySelector(".column-name-container");
  const row = document.querySelector(".row-name-container");
  for (let i = 1; i <= 100; i++) {
    column.innerHTML += `<div class="col-name col-code-${insertColumnName(
      i
    )}" id="colId-${i}">${insertColumnName(i)}</div>`;
    row.innerHTML += `<div class="row-name" id="rowId-${i}">${i}</div>`;
  }
})();

// Insert Column Name
function insertColumnName(n) {
  let str = "";
  while (n > 0) {
    const rem = n % 26;
    if (rem === 0) {
      str = "Z" + str;
      n = Math.floor(n / 26) - 1;
    } else {
      str = String.fromCharCode(rem - 1 + 65) + str;
      n = Math.floor(n / 26);
    }
  }
  return str;
}

// Create Cells
(function () {
  for (let i = 1; i <= 100; i++) {
    const row = document.createElement("div");
    row.className = "cell-row";
    for (let j = 1; j <= 100; j++) {
      row.innerHTML += `<div class="input-cell" contenteditable="false" id="row-${i}-col-${j}" data="${insertColumnName(
        j
      )}"></div>`;
    }
    document.querySelector(".input-cell-container").append(row);
  }
})();

// Make align icons selectable and applicable to the cells
(function () {
  const alignment = document.querySelectorAll(".align-icon");
  alignment.forEach((e) => {
    e.addEventListener("click", function () {
      document
        .querySelector(".align-icon.selected")
        .classList.remove("selected");
      this.classList.add("selected");
      if (
        this.classList.contains("selected") &&
        this.innerText === "format_align_left"
      ) {
        document
          .querySelector(".selected-cell")
          .classList.remove("cell-text-align-center");
        document
          .querySelector(".selected-cell")
          .classList.remove("cell-text-align-right");
        document
          .querySelector(".selected-cell")
          .classList.add("cell-text-align-left");
      }
      if (
        this.classList.contains("selected") &&
        this.innerText === "format_align_center"
      ) {
        document
          .querySelector(".selected-cell")
          .classList.remove("cell-text-align-right");
        document
          .querySelector(".selected-cell")
          .classList.remove("cell-text-align-left");
        document
          .querySelector(".selected-cell")
          .classList.add("cell-text-align-center");
      }
      if (
        this.classList.contains("selected") &&
        this.innerText === "format_align_right"
      ) {
        document
          .querySelector(".selected-cell")
          .classList.remove("cell-text-align-left");
        document
          .querySelector(".selected-cell")
          .classList.remove("cell-text-align-center");
        document
          .querySelector(".selected-cell")
          .classList.add("cell-text-align-right");
      }
    });
  });
})();

// Make style icons selectable and applicable to the cells
(function () {
  const style = document.querySelectorAll(".style-icon");
  style.forEach((e) => {
    e.addEventListener("click", function () {
      this.classList.toggle("selected");
      if (
        this.classList.contains("selected") &&
        this.innerText === "format_bold"
      ) {
        document
          .querySelector(".selected-cell")
          .classList.add("cell-text-bold");
      }
      if (
        !this.classList.contains("selected") &&
        this.innerText === "format_bold"
      ) {
        document
          .querySelector(".selected-cell")
          .classList.remove("cell-text-bold");
      }
      if (
        this.classList.contains("selected") &&
        this.innerText === "format_italic"
      ) {
        document
          .querySelector(".selected-cell")
          .classList.add("cell-text-italic");
      }
      if (
        !this.classList.contains("selected") &&
        this.innerText === "format_italic"
      ) {
        document
          .querySelector(".selected-cell")
          .classList.remove("cell-text-italic");
      }
      if (
        this.classList.contains("selected") &&
        this.innerText === "format_underlined"
      ) {
        document
          .querySelector(".selected-cell")
          .classList.add("cell-text-underline");
      }
      if (
        !this.classList.contains("selected") &&
        this.innerText === "format_underlined"
      ) {
        document
          .querySelector(".selected-cell")
          .classList.remove("cell-text-underline");
      }
    });
  });
})();

// Make cell text background color icon selectable and change background color
(function () {
  document
    .getElementById("input-cell-bg-color")
    .addEventListener("change", function () {
      document.querySelector(".selected-cell").style.backgroundColor =
        this.value;
    });
})();

// Make cell text color icon selectable and change background color
(function () {
  document
    .getElementById("input-cell-text-color")
    .addEventListener("change", function () {
      document.querySelector(".selected-cell").style.color = this.value;
    });
})();

// Make fonts applicable to the cells
(function () {
  const font = document.querySelector(".font-family-selector");
  font.addEventListener("change", function () {
    document.querySelector(".selected-cell").style.fontFamily = this.value;
  });
})();

// Make font size applicable to the cells
(function () {
  const fontSize = document.querySelector(".font-size");
  fontSize.addEventListener("change", function () {
    document.querySelector(".selected-cell").style.fontSize = `${this.value}px`;
  });
})();

// Make cells focus and contenteditable and also set the menu icon values to default on change of selected cell
(function () {
  const cells = document.querySelectorAll(".input-cell");
  cells.forEach((e) => {
    e.addEventListener("click", function (event) {
      document
        .querySelector(".selected-cell")
        .classList.remove("selected-cell");
      this.classList.add("selected-cell");
      getRowIdColName(this, this.getAttribute("data"));
    });
    e.addEventListener("dblclick", function () {
      this.setAttribute("contenteditable", "true");
      this.focus();
      document.querySelector(".font-family-selector").value = "Arial";
      document.querySelector(".font-size").value = "14";
      document
        .querySelector(".align-icon.selected")
        .classList.remove("selected");
      document.getElementById("format-align-left").classList.add("selected");
      document.querySelector("#format-bold").classList.remove("selected");
      document.querySelector("#format-italic").classList.remove("selected");
      document.querySelector("#format-underlined").classList.remove("selected");
    });
  });
})();

// Get row id, column name and display it on formula editor
function getRowIdColName(cell, colName) {
  const rowId = +cell.getAttribute("id").split("-")[1];
  document.querySelector(".formula-editor.display-cell").style.border =
    "2px solid #107C41";
  document.querySelector(
    ".formula-editor.display-cell"
  ).innerText = `${colName}${rowId}`;
    // copyToClipboard(cell)
}

// Make column name scroll in x-axis along with cells
(function () {
  document
    .querySelector(".input-cell-container")
    .addEventListener("scroll", function () {
      document.querySelector(".column-name-container").scrollLeft =
        this.scrollLeft;
    });
})();

// Make row number scroll in y-axis along with cells
(function () {
  document
    .querySelector(".input-cell-container")
    .addEventListener("scroll", function () {
      document.querySelector(".row-name-container").scrollTop = this.scrollTop;
    });
})();

