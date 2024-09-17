const downloadBtn1 = document.getElementById("download-btn-1");
const downloadBtn2 = document.getElementById("download-btn-2");

downloadBtn1.addEventListener("click", getSeparations);
downloadBtn2.addEventListener("click", getBills);

function getSeparations(){
    const fd = new FormData();
    fd.append("file", document.getElementById("input-1").files[0]);
    update("Separaciones", fd)
}

function getBills(){
    const fd = new FormData();
    fd.append("file", document.getElementById("input-2").files[0]);
    update("Facturas", fd)
}


function update(option, fd){
    fetch(`/upload/${option}`, {
        method: 'POST',
        body: fd
    })
    .then(response => {
        if (response.ok) {
            return response.blob();  // Convertir la respuesta a un blob
        } else {
            throw new Error('Failed to download file');
        }
    })
    .then(blob => {
        const url = window.URL.createObjectURL(new Blob([blob]));
        const link = document.createElement('a');
        link.href = url;
        link.setAttribute('download', option === "Separaciones" ? 'datos_filtrados.xlsx' : "facturas.zip" );  // El nombre del archivo descargado
        document.body.appendChild(link);
        link.click();
        link.parentNode.removeChild(link);
    })
    .catch(error => {
        console.error('Error:', error);
    });
}
