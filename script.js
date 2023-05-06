
let thead = document.querySelector("#table-heading");
let tbody = document.querySelector("#table-body");

let currposooutput = document.querySelector("#curr-pos");

let boldbtn = document.querySelector("#bold-btn");
let italicbtn = document.querySelector("#italic-btn");
let underlinebtn = document.querySelector("#underline-btn");
let txtcolorbtn = document.querySelector("#text-color");
let fillcolorbtn = document.querySelector("#fill-color");


let leftalignbtn = document.querySelector("#align-left-btn");
let centeralignbtn = document.querySelector("#align-center-btn");
let rightalignbtn = document.querySelector("#align-right-btn");

let fontsizebtn = document.querySelector("#font-size-select");
let fontstylebtn = document.querySelector("#font-style-select");

let copybtn = document.querySelector("#copy-btn");
let cutbtn = document.querySelector("#cut-btn");
let pastebtn = document.querySelector("#paste-btn");

let borderbtmbtn = document.querySelector("#border-btm-btn");
let bordertopbtn = document.querySelector("#border-top-btn");
let borderleftbtn = document.querySelector("#border-left-btn");
let borderrightbtn = document.querySelector("#border-right-btn");
let borderouterbtn = document.querySelector("#border-outer-btn");

let downloadbtn = document.querySelector("#download-btn");


let addsheetbtn = document.querySelector("#add-sheet-btn");
let sheet_no = 2;
let currentsheetno = 1;
let cutobj = {};

let tr = document.createElement("tr");
let th = document.createElement("th");
th.innerText = "◢";
th.style.fontSize = "15px";
tr.appendChild(th);


for (let col = 0; col < 26; col++) {
    let th = document.createElement("th");
    th.innerText = String.fromCharCode(col + 65);
    tr.appendChild(th);
}

thead.appendChild(tr);


for (let row = 0; row < 100; row++) {

    let tr = document.createElement("tr");
    let th = document.createElement("th");
    th.innerText = (row + 1);
    tr.appendChild(th);


    for (let col = 0; col < 26; col++) {
        let td = document.createElement("td");
        td.setAttribute("contenteditable", true);
        td.setAttribute("onfocus", "getcurrentcell(event)");
        td.setAttribute("oninput", "getcurrentcell(event)");
        td.setAttribute("spellcheck", false);
        td.setAttribute("id", `${String.fromCharCode(col + 65)}${row + 1}`);
        tr.appendChild(td);
    }


    tbody.appendChild(tr);
}

// making matrix data to store all the information of every cell
let array_of_matrixdata = [];
(() => {

    array_of_matrixdata = JSON.parse(localStorage.getItem('array_of_matrixdata')) || [];
    if (array_of_matrixdata.length > 0) {
        printExcelSheet(array_of_matrixdata[0]);
    }
    let n = array_of_matrixdata.length;
    for (let i = 0; i < n - 1; i++) {

        let sheetdiv = document.querySelector(".sheet-add");
        let sheettab = document.createElement("span");
        sheettab.innerText = `Sheet ${sheet_no}`;
        sheet_no++;
        sheetdiv.insertBefore(sheettab, sheetdiv.lastElementChild);

        document.querySelectorAll(".sheet-add>span:not(#add-sheet-btn)").forEach(function (sheet) {
            sheet.addEventListener("click", (event) => {
                // console.log(event.target);
                document.querySelectorAll(".sheet-add>span:not(#add-sheet-btn)").forEach(function (sh) {
                    sh.classList.remove("active-sheet");
                });
                event.target.setAttribute("class", "active-sheet");
                currentsheetno = Number(event.target.innerText.split(" ")[1]);
                console.log("currentsheetno", currentsheetno);
                printExcelSheet(array_of_matrixdata[currentsheetno - 1]);
            })
        });

    }

})();

function makeMatrixdata() {
    let matrixdata = new Array(100);

    for (let i = 0; i < 100; i++) {
        matrixdata[i] = new Array(26);
        for (let j = 0; j < 26; j++) {
            matrixdata[i][j] = {};
        }
    }

    array_of_matrixdata.push(matrixdata);
    localStorage.setItem("array_of_matrixdata", JSON.stringify(array_of_matrixdata));
    console.log(array_of_matrixdata);
}

if (array_of_matrixdata.length === 0) {
    makeMatrixdata();
    // printExcelSheet(array_of_matrixdata[0]);
}


// console.log(matrixdata);

function updateJSON(cell) {
    let obj = {
        text: cell.innerText,
        style: cell.style.cssText,
        id: cell.getAttribute("id")
    }

    let row = Number(cell.id.substring(1)) - 1;
    let column = cell.id.charCodeAt(0) - 65;
    //  console.log(row, column);
    array_of_matrixdata[currentsheetno - 1][row][column] = obj;

    localStorage.setItem("array_of_matrixdata", JSON.stringify(array_of_matrixdata));


    console.log(array_of_matrixdata);
}

let currele;

function getcurrentcell(event) {
    currele = event.target;
    let myid = currele.getAttribute("id");
    currposooutput.value = myid;
    console.log(currele);

    updateJSON(currele);
    //   fontsizebtn.value = currele.style.fontSize;
}

boldbtn.addEventListener("click", () => {

    if (currele.style.fontWeight == "bold") {
        currele.style.fontWeight = "normal";
    }
    else {
        currele.style.fontWeight = "bold";
    }

    console.log("bold clicked");
    updateJSON(currele);
});

italicbtn.addEventListener("click", () => {

    if (currele.style.fontStyle == "italic") {
        currele.style.fontStyle = "normal";
    }
    else {
        currele.style.fontStyle = "italic";
    }
    console.log("italic clicked");
    updateJSON(currele);
});

underlinebtn.addEventListener("click", () => {

    if (currele.style.textDecoration == "underline") {
        currele.style.textDecoration = "none";
    }
    else {
        currele.style.textDecoration = "underline";
    }
    console.log("underline clicked");
    updateJSON(currele);
});

txtcolorbtn.addEventListener("input", () => {

    currele.style.color = txtcolorbtn.value;

    console.log("txt-color active");
    updateJSON(currele);
});

fillcolorbtn.addEventListener("input", () => {

    currele.style.backgroundColor = fillcolorbtn.value;

    console.log("txt-color active");
    updateJSON(currele);
});

leftalignbtn.addEventListener("click", () => {

    currele.style.textAlign = "left";
    updateJSON(currele);

    // console.log("left-alignbtn active");
});

centeralignbtn.addEventListener("click", () => {
    currele.style.textAlign = "center";
    console.log("center-alignbtn active");
    updateJSON(currele);
});


rightalignbtn.addEventListener("click", () => {
    currele.style.textAlign = "right";
    console.log("right-alignbtn active");
    updateJSON(currele);
});

fontsizebtn.addEventListener("change", () => {
    currele.style.fontSize = `${fontsizebtn.value}px`;
    updateJSON(currele);

});

fontstylebtn.addEventListener("change", () => {
    // console.log(f);
    currele.style.fontFamily = fontstylebtn.value;
    updateJSON(currele);

});

cutbtn.addEventListener("click", () => {
    cutobj = {
        innertxt: currele.innerText,
        styles: currele.style.cssText
    }
    currele.innerText = "";
    currele.style = "";
    console.log(cutobj);

    updateJSON(currele);

});

copybtn.addEventListener("click", () => {
    cutobj = {
        innertxt: currele.innerText,
        styles: currele.style.cssText
    }

    console.log(cutobj);

});

pastebtn.addEventListener("click", () => {

    currele.innerText = cutobj.innertxt;
    currele.style = cutobj.styles;

    updateJSON(currele);
});

borderbtmbtn.addEventListener("click", () => {
    if (currele.style.borderBottom == "2px solid black") {
        currele.style.borderBottom = "1px solid lightgray";
    }
    else {
        currele.style.borderBottom = "2px solid black";
    }
    updateJSON(currele);
});

bordertopbtn.addEventListener("click", () => {
    if (currele.style.borderTop == "2px solid black") {
        currele.style.borderTop = "1px solid lightgray";
    }
    else {
        currele.style.borderTop = "2px solid black";
    }
    updateJSON(currele);
});

borderleftbtn.addEventListener("click", () => {
    if (currele.style.borderLeft == "2px solid black") {
        currele.style.borderLeft = "1px solid lightgray";
    }
    else {
        currele.style.borderLeft = "2px solid black";
    }
    updateJSON(currele);
});

borderrightbtn.addEventListener("click", () => {
    if (currele.style.borderRight == "2px solid black") {
        currele.style.borderRight = "1px solid lightgray";
    }
    else {
        currele.style.borderRight = "2px solid black";
    }
    updateJSON(currele);
});

borderouterbtn.addEventListener("click", () => {
    if (currele.style.border == "2px solid black") {
        currele.style.border = "1px solid lightgray";
    }
    else {
        currele.style.border = "2px solid black";
    }
    updateJSON(currele);
});

addsheetbtn.addEventListener("click", () => {
    let sheetdiv = document.querySelector(".sheet-add");
    let sheettab = document.createElement("span");
    sheettab.innerText = `Sheet ${sheet_no}`;
    sheet_no++;
    sheetdiv.insertBefore(sheettab, sheetdiv.lastElementChild);

    document.querySelectorAll(".sheet-add>span:not(#add-sheet-btn)").forEach(function (sheet) {
        sheet.addEventListener("click", (event) => {
            // console.log(event.target);
            document.querySelectorAll(".sheet-add>span:not(#add-sheet-btn)").forEach(function (sh) {
                sh.classList.remove("active-sheet");
            });
            event.target.setAttribute("class", "active-sheet");
            currentsheetno = Number(event.target.innerText.split(" ")[1]);
            console.log("currentsheetno", currentsheetno);
            printExcelSheet(array_of_matrixdata[currentsheetno - 1]);
        })
    });

    makeMatrixdata();
});


function downloadJSON() {

    //need to be converted to string format
    let jsonmatrixstring = JSON.stringify(matrixdata);

    let blob = new Blob([jsonmatrixstring], { type: "application/json" });

    let link = document.createElement("a");

    link.href = URL.createObjectURL(blob);

    link.download = "matrixdata.json"; // filename by which this will be downloaded

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);


}


downloadbtn.addEventListener("click", downloadJSON);


document.querySelector("#upload-btn").addEventListener("change", (event) => {
    let file = event.target.files[0];

    if (file) {
        const reader = new FileReader();

        reader.onload = function (e) {
            const fileContent = e.target.result;

            // {id,style,text}
            // Parse the JSON file content and process the data
            try {
                const jsondata = JSON.parse(fileContent);
                console.log(jsondata);

                printExcelSheet(jsondata);

            } catch (error) {
                console.error("Error parsing JSON file:", error);
            }
        };

        reader.readAsText(file);
    }
});


function printExcelSheet(jsondata) {
    console.log(jsondata);
    // jsondata.forEach((row) => {
    //     row.forEach((col) => {
    //         if (col.id) {
    //             let cell = document.getElementById(`${col.id}`)
    //             cell.innerText = col.text;
    //             cell.style.cssText = col.style;
    //         }


    //     });
    // });

    for (let i = 0; i < 100; i++) {
        for (let j = 0; j < 26; j++) {
            let cell = document.getElementById(`${String.fromCharCode(j + 65)}${i + 1}`);
            if (jsondata[i][j].id) {
                cell.innerText = jsondata[i][j].text;
                cell.style.cssText = jsondata[i][j].style;
            }
            else {
                cell.innerHTML = "";
                cell.style.cssText = "";
            }

        }
    }
}


