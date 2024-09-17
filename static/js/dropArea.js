const dropArea1 = document.getElementById("dp1");
const dropArea2 = document.getElementById("dp2");

const inputFile1 = document.getElementById("input-1");
const inputFile2 = document.getElementById("input-2");

const fileContent1 = document.getElementById("file-1");
const fileName1 = document.getElementById("file-name-1");
const fileContent2 = document.getElementById("file-2");
const fileName2 = document.getElementById("file-name-2");


// INIT FILE 1

inputFile1.addEventListener("change", uploadFile1);

dropArea1.addEventListener("dragover", function(e){
    e.preventDefault();
});

dropArea1.addEventListener("drop", function(e) {
    e.preventDefault();
    let files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.endsWith('.xlsx')) {
        inputFile1.files = files;
        uploadFile1();
    } else {
        alert("Por favor, sube solo archivos excel");
    }
});

function uploadFile1(){
    fileContent1.style.display = "flex"
    fileName1.innerText = inputFile1.files[0].name;
}

// INIT FILE 2

inputFile2.addEventListener("change", uploadFile2);

dropArea2.addEventListener("dragover", function(e){
    e.preventDefault();
});

dropArea2.addEventListener("drop", function(e) {
    e.preventDefault();
    let files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.endsWith('.xlsx')) {
        inputFile2.files = files;
        uploadFile2();
    } else {
        alert("Por favor, sube solo archivos excel");
    }
});

function uploadFile2(){
    fileContent2.style.display = "flex"
    fileName2.innerText = inputFile2.files[0].name;
}
