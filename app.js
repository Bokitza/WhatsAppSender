const dropArea = document.getElementById("drag");
const phones = document.getElementById("phones");

var jsoned = null;
var merged = [];
var reader = null;
var j = 0;
sheet = null;

["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
  dropArea.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
  e.preventDefault();
  e.stopPropagation();
}

var files = [];
function handleDrop(e) {
  e.target.innerText = "קובץ הועלה בהצלחה";
  e.target.style.background = "lightgreen";
  //Things against some kinda bugs... boilerplate
  e.stopPropagation();
  e.preventDefault();
  //convert the target into an array of files
  for (let i = 0; i < e.dataTransfer.files.length; i++) {
    files.push(e.dataTransfer.files[i]);
  }

  //pops the header selector buttons!
  //const uploadBtn=document.getElementById("upload");
  //document.body.appendChild(headerHelper);
  //document.body.appendChild(headers);
}

//reading and appending the files to Merged array
function createLink() {
  var reader = new FileReader();
  reader.readAsArrayBuffer(files[j]);

  reader.onload = function (elem) {
    var data = new Uint8Array(elem.target.result);
    var workbook = XLSX.read(data, { type: "array" });
    var sheet = workbook.Sheets[workbook.SheetNames[0]];
    const keys = Object.keys(sheet);
    keys.forEach((e) => {
      if (e[0] !== "C") {
        return;
      }
      if (e[1] === "1") {
        return;
      }
      const value = sheet[e].v.split("-").join("");
      var link = document.createElement("a");
      link.setAttribute("href", value);
      link.setAttribute("target", `_blank`);
      phones.appendChild(link);
      link.click();
      link.remove();

      //("afterbegin", `<a target="_blank" href=${keysOf[0]}> ${keysOf[0].slice(30,48)} </a> <br>`)
    });
  };
}

function create() {
  var newBook = XLSX.utils.book_new(),
    ws = XLSX.utils.aoa_to_sheet(merged);
  XLSX.utils.book_append_sheet(newBook, ws, "mytest");
  XLSX.writeFile(newBook, "testMerge.xls");
}

document.addEventListener("dragenter", (e) => {
  dropArea.innerHTML = "גרור קובץ הנה";
});

dropArea.addEventListener("drop", handleDrop);
